'use strict';

/**
 * Erzeugt ZUGFeRD 2.3 / Factur-X XML (CII - Cross Industry Invoice)
 * konforme zur EN 16931.
 *
 * Unterstützte Profile:
 *   MINIMUM   – Pflichtfelder, keine Positionsdetails
 *   BASIC WL  – Pflichtfelder + Zahlungsinformationen, keine Positionsdetails
 *   EN 16931  – Vollständige Rechnung mit Positionen (Comfort)
 *   EXTENDED  – Wie EN 16931 + optionale Felder
 */

// Profil-URNs und Dateinamen gemäß ZUGFeRD 2.3 Spezifikation
const PROFILES = {
  MINIMUM:   { urn: 'urn:factur-x.eu:1p0:minimum',   level: 'MINIMUM' },
  BASIC_WL:  { urn: 'urn:factur-x.eu:1p0:basicwl',   level: 'BASIC WL' },
  EN16931:   { urn: 'urn:cen.eu:en16931:2017#compliant#urn:factur-x.eu:1p0:en16931', level: 'EN 16931' },
  EXTENDED:  { urn: 'urn:cen.eu:en16931:2017#conformant#urn:factur-x.eu:1p0:extended', level: 'EXTENDED' },
};

/**
 * Parst einen deutschen Dezimalbetrag ("1.234,56 €" → 1234.56)
 */
function parseAmount(str) {
  if (!str) return 0;
  const cleaned = str.replace(/[^\d,.-]/g, '').replace('.', '').replace(',', '.');
  return parseFloat(cleaned) || 0;
}

/**
 * Formatiert eine Zahl auf 2 Dezimalstellen im XML (immer Punkt als Trenner)
 */
function fmt(n) {
  return Number(n).toFixed(2);
}

/**
 * Wandelt deutsches Datum "DD.MM.YYYY" → "YYYYMMDD" (ZUGFeRD-Format)
 */
function parseDate(dateStr) {
  if (!dateStr) return '';
  const m = dateStr.match(/(\d{2})\.(\d{2})\.(\d{4})/);
  if (!m) return dateStr.replace(/-/g, '');
  return `${m[3]}${m[2]}${m[1]}`;
}

/**
 * Extrahiert den MwSt-Satz als Zahl aus "19 %" → 19, "7%" → 7, "0 %" → 0
 */
function parseTaxRate(str) {
  if (!str) return 0;
  const m = str.match(/(\d+)/);
  return m ? parseInt(m[1], 10) : 0;
}

/**
 * Gibt den ZUGFeRD Steuer-Kategorie-Code zurück
 * S = Standard, Z = Zero/Null, E = Exempt/Steuerfrei
 */
function taxCategory(rate) {
  if (rate === 0) return 'Z';   // 0% wird als Zero Rate behandelt
  return 'S';                    // 7% und 19% = Standard
}

/**
 * Baut den XML-String für eine Rechnungsposition (nur EN16931 / EXTENDED)
 */
function buildLineItem(pos, index) {
  const qty       = parseAmount(pos.menge);
  const unitPrice = parseAmount(pos.einzelpreis);
  const lineNet   = parseAmount(pos.betrag);
  const taxRate   = parseTaxRate(pos.mwst_satz);
  const catCode   = taxCategory(taxRate);

  // Einheit: Std. → HUR, Psch. → LS, Stk. → C62, sonstige → ZZ
  const unitMap = { 'Std.': 'HUR', 'h': 'HUR', 'Stk.': 'C62', 'Stück': 'C62',
                    'Psch.': 'LS', 'pauschal': 'LS', 'km': 'KMT', 'Tag': 'DAY',
                    'Tage': 'DAY', 'Mon.': 'MON', 'Monat': 'MON' };
  const unitCode = unitMap[pos.einheit] || 'ZZ';

  return `
    <ram:IncludedSupplyChainTradeLineItem>
      <ram:AssociatedDocumentLineDocument>
        <ram:LineID>${index + 1}</ram:LineID>
      </ram:AssociatedDocumentLineDocument>
      <ram:SpecifiedTradeProduct>
        <ram:Name>${escXml(pos.beschreibung_titel || pos.beschreibung || `Position ${index + 1}`)}</ram:Name>
        ${pos.beschreibung_detail ? `<ram:Description>${escXml(pos.beschreibung_detail)}</ram:Description>` : ''}
      </ram:SpecifiedTradeProduct>
      <ram:SpecifiedLineTradeAgreement>
        <ram:NetPriceProductTradePrice>
          <ram:ChargeAmount>${fmt(unitPrice)}</ram:ChargeAmount>
        </ram:NetPriceProductTradePrice>
      </ram:SpecifiedLineTradeAgreement>
      <ram:SpecifiedLineTradeDelivery>
        <ram:BilledQuantity unitCode="${unitCode}">${fmt(qty)}</ram:BilledQuantity>
      </ram:SpecifiedLineTradeDelivery>
      <ram:SpecifiedLineTradeSettlement>
        <ram:ApplicableTradeTax>
          <ram:TypeCode>VAT</ram:TypeCode>
          <ram:CategoryCode>${catCode}</ram:CategoryCode>
          <ram:RateApplicablePercent>${taxRate}.00</ram:RateApplicablePercent>
        </ram:ApplicableTradeTax>
        <ram:SpecifiedTradeSettlementLineMonetarySummation>
          <ram:LineTotalAmount>${fmt(lineNet)}</ram:LineTotalAmount>
        </ram:SpecifiedTradeSettlementLineMonetarySummation>
      </ram:SpecifiedLineTradeSettlement>
    </ram:IncludedSupplyChainTradeLineItem>`;
}

/**
 * Escaping für XML-Textinhalte
 */
function escXml(str) {
  if (!str) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

/**
 * Berechnet Steuer-Aufschlüsselung aus Positionen
 * Gibt Array von { rate, basis, betrag, category } zurück
 */
function calcTaxBreakdown(positions) {
  const groups = {};
  for (const pos of positions) {
    const rate = parseTaxRate(pos.mwst_satz);
    const net  = parseAmount(pos.betrag);
    const key  = String(rate);
    if (!groups[key]) groups[key] = { rate, basis: 0, betrag: 0, category: taxCategory(rate) };
    groups[key].basis  += net;
    groups[key].betrag += net * rate / 100;
  }
  return Object.values(groups).sort((a, b) => b.rate - a.rate);
}

// ─────────────────────────────────────────────────────────────────────────────
/**
 * Hauptfunktion – erzeugt ZUGFeRD XML als String
 *
 * @param {object} data  – Rechnungsdaten (aus Office.js Content Controls)
 * @param {string} profileKey – 'MINIMUM' | 'BASIC_WL' | 'EN16931' | 'EXTENDED'
 * @returns {string}  XML-String (UTF-8)
 */
function generateZugferdXml(data, profileKey = 'EN16931') {
  const profile = PROFILES[profileKey] || PROFILES.EN16931;
  const isDetailed = ['EN16931', 'EXTENDED'].includes(profileKey);

  const positions = (data.positions || []).filter(p =>
    p.beschreibung_titel || p.beschreibung || p.betrag
  );

  const taxGroups   = calcTaxBreakdown(positions);
  const nettoTotal  = parseAmount(data.summe_netto);
  const bruttoTotal = parseAmount(data.summe_brutto);
  const taxTotal    = bruttoTotal - nettoTotal;

  const reDatum   = parseDate(data.rechnung_datum);
  const lieferDat = parseDate(data.lieferdatum || data.rechnung_datum);
  const faellig   = parseDate(data.faelligkeitsdatum);

  // Verkäufer (Aussteller) – aus fest konfigurierten Serverdaten
  const seller = data.seller || {};

  const xml = `<?xml version="1.0" encoding="UTF-8"?>
<rsm:CrossIndustryInvoice
  xmlns:rsm="urn:un:unece:uncefact:data:standard:CrossIndustryInvoice:100"
  xmlns:qdt="urn:un:unece:uncefact:data:standard:QualifiedDataType:100"
  xmlns:ram="urn:un:unece:uncefact:data:standard:ReusableAggregateBusinessInformationEntity:100"
  xmlns:xs="http://www.w3.org/2001/XMLSchema"
  xmlns:udt="urn:un:unece:uncefact:data:standard:UnqualifiedDataType:100">

  <!-- ── Dokumentkopf ───────────────────────────────────────────────── -->
  <rsm:ExchangedDocumentContext>
    <ram:GuidelineSpecifiedDocumentContextParameter>
      <ram:ID>${profile.urn}</ram:ID>
    </ram:GuidelineSpecifiedDocumentContextParameter>
  </rsm:ExchangedDocumentContext>

  <rsm:ExchangedDocument>
    <ram:ID>${escXml(data.rechnung_nummer)}</ram:ID>
    <ram:TypeCode>380</ram:TypeCode><!-- 380 = Rechnung -->
    <ram:IssueDateTime>
      <udt:DateTimeString format="102">${reDatum}</udt:DateTimeString>
    </ram:IssueDateTime>
    ${data.leistungsmonat ? `<ram:IncludedNote>
      <ram:Content>Leistungsmonat: ${escXml(data.leistungsmonat)}</ram:Content>
    </ram:IncludedNote>` : ''}
  </rsm:ExchangedDocument>

  <!-- ── Handels-Transaktion ────────────────────────────────────────── -->
  <rsm:SupplyChainTradeTransaction>

    <!-- Positionen (nur EN16931 / EXTENDED) -->
    ${isDetailed ? positions.map((p, i) => buildLineItem(p, i)).join('\n') : '<!-- Positionen nicht im gewählten Profil -->'}

    <!-- ── Kopfdaten der Transaktion ──────────────────────────────── -->
    <ram:ApplicableHeaderTradeAgreement>

      <!-- Verkäufer -->
      <ram:SellerTradeParty>
        <ram:Name>${escXml(seller.name)}</ram:Name>
        <ram:PostalTradeAddress>
          <ram:PostcodeCode>${escXml(seller.plz)}</ram:PostcodeCode>
          <ram:LineOne>${escXml(seller.strasse)}</ram:LineOne>
          <ram:CityName>${escXml(seller.ort)}</ram:CityName>
          <ram:CountryID>${escXml(seller.land || 'DE')}</ram:CountryID>
        </ram:PostalTradeAddress>
        ${seller.ust_id ? `<ram:SpecifiedTaxRegistration>
          <ram:ID schemeID="VA">${escXml(seller.ust_id)}</ram:ID>
        </ram:SpecifiedTaxRegistration>` : ''}
        ${seller.steuernummer ? `<ram:SpecifiedTaxRegistration>
          <ram:ID schemeID="FC">${escXml(seller.steuernummer)}</ram:ID>
        </ram:SpecifiedTaxRegistration>` : ''}
      </ram:SellerTradeParty>

      <!-- Käufer -->
      <ram:BuyerTradeParty>
        <ram:Name>${escXml(data.empfaenger_firma)}</ram:Name>
        <ram:PostalTradeAddress>
          <ram:PostcodeCode>${escXml((data.empfaenger_plz_ort || '').split(' ')[0])}</ram:PostcodeCode>
          <ram:LineOne>${escXml(data.empfaenger_strasse)}</ram:LineOne>
          <ram:CityName>${escXml((data.empfaenger_plz_ort || '').split(' ').slice(1).join(' '))}</ram:CityName>
          <ram:CountryID>${escXml(data.empfaenger_land_code || 'DE')}</ram:CountryID>
        </ram:PostalTradeAddress>
        ${data.empfaenger_ust_id ? `<ram:SpecifiedTaxRegistration>
          <ram:ID schemeID="VA">${escXml(data.empfaenger_ust_id.replace(/[^A-Z0-9]/g, ''))}</ram:ID>
        </ram:SpecifiedTaxRegistration>` : ''}
      </ram:BuyerTradeParty>

      ${data.empfaenger_bestellnr ? `<ram:BuyerOrderReferencedDocument>
        <ram:IssuerAssignedID>${escXml(data.empfaenger_bestellnr)}</ram:IssuerAssignedID>
      </ram:BuyerOrderReferencedDocument>` : ''}
    </ram:ApplicableHeaderTradeAgreement>

    <!-- ── Lieferinformationen ─────────────────────────────────────── -->
    <ram:ApplicableHeaderTradeDelivery>
      ${lieferDat ? `<ram:ActualDeliverySupplyChainEvent>
        <ram:OccurrenceDateTime>
          <udt:DateTimeString format="102">${lieferDat}</udt:DateTimeString>
        </ram:OccurrenceDateTime>
      </ram:ActualDeliverySupplyChainEvent>` : ''}
    </ram:ApplicableHeaderTradeDelivery>

    <!-- ── Zahlungs- und Steuerinformationen ──────────────────────── -->
    <ram:ApplicableHeaderTradeSettlement>
      <ram:PaymentReference>${escXml(data.rechnung_nummer)}</ram:PaymentReference>
      <ram:InvoiceCurrencyCode>${escXml(data.waehrung || 'EUR')}</ram:InvoiceCurrencyCode>

      <!-- Bankverbindung -->
      ${seller.iban ? `<ram:SpecifiedTradeSettlementPaymentMeans>
        <ram:TypeCode>58</ram:TypeCode><!-- 58 = SEPA Überweisung -->
        <ram:PayeePartyCreditorFinancialAccount>
          <ram:IBANID>${escXml(seller.iban.replace(/\s/g, ''))}</ram:IBANID>
        </ram:PayeePartyCreditorFinancialAccount>
        ${seller.bic ? `<ram:PayeeSpecifiedCreditorFinancialInstitution>
          <ram:BICID>${escXml(seller.bic)}</ram:BICID>
        </ram:PayeeSpecifiedCreditorFinancialInstitution>` : ''}
      </ram:SpecifiedTradeSettlementPaymentMeans>` : ''}

      <!-- Steueraufschlüsselung -->
      ${taxGroups.map(g => `
      <ram:ApplicableTradeTax>
        <ram:CalculatedAmount>${fmt(g.betrag)}</ram:CalculatedAmount>
        <ram:TypeCode>VAT</ram:TypeCode>
        <ram:BasisAmount>${fmt(g.basis)}</ram:BasisAmount>
        <ram:CategoryCode>${g.category}</ram:CategoryCode>
        <ram:RateApplicablePercent>${g.rate}.00</ram:RateApplicablePercent>
      </ram:ApplicableTradeTax>`).join('')}

      <!-- Zahlungsziel -->
      ${faellig ? `<ram:SpecifiedTradePaymentTerms>
        <ram:DueDateDateTime>
          <udt:DateTimeString format="102">${faellig}</udt:DateTimeString>
        </ram:DueDateDateTime>
        ${data.zahlungsziel ? `<ram:Description>${escXml(data.zahlungsziel)}</ram:Description>` : ''}
      </ram:SpecifiedTradePaymentTerms>` : ''}

      <!-- Gesamtbeträge -->
      <ram:SpecifiedTradeSettlementHeaderMonetarySummation>
        <ram:LineTotalAmount>${fmt(nettoTotal)}</ram:LineTotalAmount>
        <ram:TaxBasisTotalAmount>${fmt(nettoTotal)}</ram:TaxBasisTotalAmount>
        <ram:TaxTotalAmount currencyID="${escXml(data.waehrung || 'EUR')}">${fmt(taxTotal)}</ram:TaxTotalAmount>
        <ram:GrandTotalAmount>${fmt(bruttoTotal)}</ram:GrandTotalAmount>
        <ram:DuePayableAmount>${fmt(bruttoTotal)}</ram:DuePayableAmount>
      </ram:SpecifiedTradeSettlementHeaderMonetarySummation>
    </ram:ApplicableHeaderTradeSettlement>

  </rsm:SupplyChainTradeTransaction>
</rsm:CrossIndustryInvoice>`;

  return xml;
}

module.exports = { generateZugferdXml, PROFILES };
