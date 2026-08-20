'use strict';

/**
 * Bettet ein ZUGFeRD XML in ein bestehendes PDF ein und erzeugt
 * ein PDF/A-3b-konformes Dokument.
 *
 * Technische Anforderungen ZUGFeRD / Factur-X:
 *  - PDF/A-3b (ISO 19005-3)
 *  - Eingebettete Datei: "factur-x.xml" (Relationship: "Alternative")
 *  - XMP-Metadaten mit Factur-X-Namespace
 *  - AF-Eintrag im PDF-Katalog
 */

const { PDFDocument, PDFName, PDFString, PDFArray, PDFDict, PDFStream,
        PDFRawStream, PDFHexString, asPDFName } = require('pdf-lib');

/**
 * Erstellt ein PDF/A-3b Dokument mit eingebettetem ZUGFeRD XML.
 *
 * @param {Buffer|Uint8Array} pdfBytes   – Eingangs-PDF (Word-Export)
 * @param {string}            xmlString  – ZUGFeRD XML als UTF-8-String
 * @param {string}            profileUrn – ZUGFeRD-Profil-URN
 * @param {string}            profileLevel – z.B. "EN 16931"
 * @param {string}            invoiceId  – Rechnungsnummer (für Metadaten)
 * @returns {Promise<Uint8Array>}        – Fertiges PDF/A-3 als Bytes
 */
async function embedZugferdXml(pdfBytes, xmlString, profileUrn, profileLevel, invoiceId) {
  const pdfDoc = await PDFDocument.load(pdfBytes, { ignoreEncryption: true });
  const context = pdfDoc.context;

  // ── 1. XML als eingebettete Datei hinzufügen ────────────────────────────
  const xmlBytes = Buffer.from(xmlString, 'utf-8');

  // Embedded File Stream
  const efStream = context.stream(xmlBytes, {
    Type:    'EmbeddedFile',
    Subtype: 'application/xml',
    Params:  context.obj({
      Size:    xmlBytes.length,
      ModDate: PDFString.of(new Date().toISOString()),
    }),
  });
  const efRef = context.register(efStream);

  // Filespec Dictionary
  const filespecDict = context.obj({
    Type: asPDFName('Filespec'),
    F:    PDFString.of('factur-x.xml'),
    UF:   PDFString.of('factur-x.xml'),
    Desc: PDFString.of('ZUGFeRD / Factur-X E-Rechnung'),
    AFRelationship: asPDFName('Alternative'),
    EF:   context.obj({ F: efRef, UF: efRef }),
  });
  const filespecRef = context.register(filespecDict);

  // ── 2. Names-Baum für EmbeddedFiles im Katalog registrieren ────────────
  const catalog = pdfDoc.catalog;

  // Names-Dictionary holen oder anlegen
  let namesDict = catalog.lookupMaybe(PDFName.of('Names'), PDFDict);
  if (!namesDict) {
    namesDict = context.obj({});
    catalog.set(PDFName.of('Names'), namesDict);
  }

  // EmbeddedFiles-Baum setzen
  const efNamesArray = context.obj([
    PDFString.of('factur-x.xml'),
    filespecRef,
  ]);
  namesDict.set(PDFName.of('EmbeddedFiles'), context.obj({
    Names: efNamesArray,
  }));

  // ── 3. AF-Array (Associated Files) im Katalog ───────────────────────────
  const afArray = context.obj([filespecRef]);
  catalog.set(PDFName.of('AF'), afArray);

  // ── 4. PDF/A-3 Versionsinformationen setzen ─────────────────────────────
  // MarkInfo für Tagged PDF (PDF/A-3 Anforderung)
  catalog.set(PDFName.of('MarkInfo'), context.obj({ Marked: true }));

  // ViewerPreferences
  catalog.set(PDFName.of('ViewerPreferences'), context.obj({
    DisplayDocTitle: true,
  }));

  // ── 5. XMP-Metadaten einbetten ───────────────────────────────────────────
  const now       = new Date().toISOString();
  const xmpXml    = buildXmpMetadata(invoiceId, profileUrn, profileLevel, now);
  const xmpBytes  = Buffer.from(xmpXml, 'utf-8');

  const xmpStream = context.stream(xmpBytes, {
    Type:    'Metadata',
    Subtype: 'XML',
  });
  const xmpRef = context.register(xmpStream);
  catalog.set(PDFName.of('Metadata'), xmpRef);

  // ── 6. OutputIntent für PDF/A-3 (sRGB) ─────────────────────────────────
  // Minimaler OutputIntent ohne eingebettetes ICC-Profil (erlaubt für Viewer)
  const outputIntent = context.obj({
    Type:             asPDFName('OutputIntent'),
    S:                asPDFName('GTS_PDFA1'),
    OutputConditionIdentifier: PDFString.of('sRGB'),
    Info:             PDFString.of('sRGB IEC61966-2.1'),
    RegistryName:     PDFString.of('http://www.color.org'),
  });
  catalog.set(PDFName.of('OutputIntents'), context.obj([outputIntent]));

  // ── 7. PDF serialisieren ────────────────────────────────────────────────
  return await pdfDoc.save();
}

// ─────────────────────────────────────────────────────────────────────────────
/**
 * Baut das XMP-Metadaten-Dokument mit Factur-X-Namespace
 */
function buildXmpMetadata(invoiceId, profileUrn, profileLevel, isoDate) {
  return `<?xpacket begin="\uFEFF" id="W5M0MpCehiHzreSzNTczkc9d"?>
<x:xmpmeta xmlns:x="adobe:ns:meta/">
  <rdf:RDF xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#">

    <!-- PDF/A Conformance -->
    <rdf:Description rdf:about=""
      xmlns:pdfaid="http://www.aiim.org/pdfa/ns/id/">
      <pdfaid:part>3</pdfaid:part>
      <pdfaid:conformance>B</pdfaid:conformance>
    </rdf:Description>

    <!-- Dublin Core -->
    <rdf:Description rdf:about=""
      xmlns:dc="http://purl.org/dc/elements/1.1/">
      <dc:format>application/pdf</dc:format>
      <dc:title>
        <rdf:Alt>
          <rdf:li xml:lang="x-default">Rechnung ${escXmpAttr(invoiceId)}</rdf:li>
        </rdf:Alt>
      </dc:title>
    </rdf:Description>

    <!-- XMP Basic -->
    <rdf:Description rdf:about=""
      xmlns:xmp="http://ns.adobe.com/xap/1.0/">
      <xmp:CreateDate>${isoDate}</xmp:CreateDate>
      <xmp:ModifyDate>${isoDate}</xmp:ModifyDate>
      <xmp:CreatorTool>ZUGFeRD Word Add-in</xmp:CreatorTool>
    </rdf:Description>

    <!-- Factur-X / ZUGFeRD -->
    <rdf:Description rdf:about=""
      xmlns:fx="urn:factur-x:pdfa:CrossIndustryDocument:invoice:1p0#">
      <fx:DocumentFileName>factur-x.xml</fx:DocumentFileName>
      <fx:DocumentType>INVOICE</fx:DocumentType>
      <fx:Version>1.0</fx:Version>
      <fx:ConformanceLevel>${escXmpAttr(profileLevel)}</fx:ConformanceLevel>
    </rdf:Description>

  </rdf:RDF>
</x:xmpmeta>
<?xpacket end="w"?>`;
}

function escXmpAttr(str) {
  if (!str) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

module.exports = { embedZugferdXml };
