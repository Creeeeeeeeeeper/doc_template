// Generates sample-template.docx in the project root.
// Run: node scripts/make-sample-template.js
import PizZip from "pizzip";
import { writeFileSync } from "node:fs";

const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Default Extension="jpeg" ContentType="image/jpeg"/>
  <Default Extension="jpg" ContentType="image/jpeg"/>
  <Default Extension="gif" ContentType="image/gif"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>`;

const rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

const documentRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId100" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId101" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>`;

const styles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>
    <w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/><w:qFormat/>
  </w:style>
</w:styles>`;

const settings = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`;

// Full namespace declarations are required so that <wp:inline>, r:embed, a14:* etc.
// produced by the image module are valid when Word opens the file.
const NS =
  ' xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"' +
  ' xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"' +
  ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"' +
  ' xmlns:o="urn:schemas-microsoft-com:office:office"' +
  ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"' +
  ' xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"' +
  ' xmlns:v="urn:schemas-microsoft-com:vml"' +
  ' xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"' +
  ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"' +
  ' xmlns:w10="urn:schemas-microsoft-com:office:word"' +
  ' xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"' +
  ' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"' +
  ' xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"' +
  ' xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex"' +
  ' xmlns:wpg="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingGroup"' +
  ' xmlns:wpi="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingInk"' +
  ' xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"' +
  ' xmlns:wps="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingShape"' +
  ' mc:Ignorable="w14 w15 w16se wp14"';

const sectPr =
  '<w:sectPr>' +
  '<w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
  '<w:cols w:space="720"/>' +
  '<w:docGrid w:linePitch="360"/>' +
  '</w:sectPr>';

const document = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document${NS}>
  <w:body>
    <w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="36"/></w:rPr><w:t>员工档案</w:t></w:r></w:p>
    <w:p><w:r><w:t xml:space="preserve">姓名：</w:t></w:r><w:r><w:t>{@name}</w:t></w:r></w:p>
    <w:p><w:r><w:t xml:space="preserve">职位：</w:t></w:r><w:r><w:t>{@title}</w:t></w:r></w:p>
    <w:p><w:r><w:t xml:space="preserve">入职日期：</w:t></w:r><w:r><w:t>{@joinDate}</w:t></w:r></w:p>
    <w:p><w:r><w:t xml:space="preserve">个人简介：</w:t></w:r></w:p>
    <w:p><w:r><w:t>{@bio}</w:t></w:r></w:p>
    <w:p><w:r><w:t xml:space="preserve">照片：</w:t></w:r></w:p>
    <w:p><w:r><w:t>{%avatar}</w:t></w:r></w:p>
    ${sectPr}
  </w:body>
</w:document>`;

const zip = new PizZip();
zip.file("[Content_Types].xml", contentTypes);
zip.folder("_rels").file(".rels", rels);
zip.folder("word").file("document.xml", document);
zip.folder("word/_rels").file("document.xml.rels", documentRels);
zip.folder("word").file("styles.xml", styles);
zip.folder("word").file("settings.xml", settings);

const out = zip.generate({ type: "nodebuffer" });
writeFileSync("sample-template.docx", out);
console.log("Wrote sample-template.docx (" + out.length + " bytes)");
