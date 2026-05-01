// Smoke test: exercise both fill-mode and edit-mode pipelines end-to-end.
// Run: node scripts/smoke-test.js
import { readFileSync, writeFileSync } from "node:fs";
import {
  PizZip,
  extractFieldsFromZip,
  parseParagraphs,
  renderFilled,
  buildTemplate,
} from "../src/docx.js";

let pass = 0;
let fail = 0;
function check(label, ok) {
  console.log(`  ${ok ? "PASS" : "FAIL"}  ${label}`);
  ok ? pass++ : fail++;
}

// --------------------------------------------------------------
// Test 1: Fill mode using sample-template.docx
// --------------------------------------------------------------
console.log("\n=== Test 1: fill mode ===");
{
  const templateBytes = readFileSync("sample-template.docx");
  const imageBytes = readFileSync("src-tauri/icons/128x128.png");

  const zip = new PizZip(templateBytes);
  const fields = extractFieldsFromZip(zip);
  console.log("Detected fields:", fields.map((f) => `${f.name}(${f.type})`).join(", "));

  const values = {
    name: { text: "张三", font: "微软雅黑", size: 16, color: "#1F4E79" },
    title: { text: "高级工程师", font: "Arial", size: 12, color: "#0F0F0F" },
    joinDate: { text: "2024-03-15", font: "Calibri", size: 12, color: "#666666" },
    bio: {
      text: "热爱开发\n喜欢造轮子\n现居北京。",
      font: "宋体",
      size: 11,
      color: "#222222",
    },
    avatar: { bytes: imageBytes },
  };

  const out = renderFilled(templateBytes, fields, values);
  writeFileSync("smoke-output.docx", out);

  const outZip = new PizZip(out);
  const outXml = outZip.file("word/document.xml").asText();

  check("fields detected (5)", fields.length === 5);
  check("{@name} placeholder removed", !outXml.includes("{@name}"));
  check("{%avatar} placeholder removed", !outXml.includes("{%avatar}"));
  check("text replacement present", outXml.includes("张三"));
  check("multiline -> <w:br/>", outXml.includes("热爱开发") && outXml.includes("<w:br/>"));
  check("custom font applied", outXml.includes('w:ascii="微软雅黑"'));
  check("custom color applied", outXml.includes('w:val="1F4E79"'));
  check("image drawing inserted", outXml.includes("<w:drawing>"));
  check(
    "image media file present",
    Object.keys(outZip.files).some((k) => /^word\/media\//.test(k)),
  );
  check(
    "wp namespace declared on root",
    /xmlns:wp="http:\/\/schemas\.openxmlformats\.org\/drawingml\/2006\/wordprocessingDrawing"/.test(
      outXml,
    ),
  );
  check(
    "r namespace declared on root",
    /xmlns:r="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships"/.test(
      outXml,
    ),
  );
  check("sectPr preserved", outXml.includes("<w:sectPr>"));
  check("styles.xml present", !!outZip.file("word/styles.xml"));
}

// --------------------------------------------------------------
// Test 2: Edit-mode roundtrip
//  plain docx → add placeholders → save → load → fill → verify
// --------------------------------------------------------------
console.log("\n=== Test 2: edit mode roundtrip ===");
{
  // Build a complete minimum docx (styles.xml + settings.xml + full
  // namespaces + sectPr) so the outputs Word can actually open.
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
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
    '<w:cols w:space="720"/><w:docGrid w:linePitch="360"/></w:sectPr>';

  const plainDocXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document${NS}>
  <w:body>
    <w:p><w:r><w:t>合同书</w:t></w:r></w:p>
    <w:p><w:r><w:t xml:space="preserve">甲方：张三</w:t></w:r></w:p>
    <w:p><w:r><w:t xml:space="preserve">乙方：李四</w:t></w:r></w:p>
    <w:p><w:r><w:t>签字：</w:t></w:r></w:p>
    <w:p><w:r><w:t>(此处放签名图)</w:t></w:r></w:p>
    ${sectPr}
  </w:body>
</w:document>`;

  const plainContentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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

  const plainRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

  const plainDocRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId100" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId101" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>`;

  const plainStyles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>
</w:styles>`;

  const plainSettings = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`;

  const plainZip = new PizZip();
  plainZip.file("[Content_Types].xml", plainContentTypes);
  plainZip.folder("_rels").file(".rels", plainRels);
  plainZip.folder("word").file("document.xml", plainDocXml);
  plainZip.folder("word/_rels").file("document.xml.rels", plainDocRels);
  plainZip.folder("word").file("styles.xml", plainStyles);
  plainZip.folder("word").file("settings.xml", plainSettings);
  const plainBytes = plainZip.generate({ type: "uint8array" });

  // 2a) Parse paragraphs and edit them
  const parseZip = new PizZip(plainBytes);
  const paragraphs = parseParagraphs(parseZip);
  console.log(`Detected ${paragraphs.length} paragraphs`);

  check("found 5 body paragraphs", paragraphs.filter((p) => p.originalText).length === 5);

  // Edit paragraph 1 ("甲方：张三") -> "甲方：{@partyA}"
  // Edit paragraph 2 ("乙方：李四") -> "乙方：{@partyB}"
  // Edit paragraph 4 ("(此处放签名图)") -> "{%signature}"
  paragraphs[1].currentText = "甲方：{@partyA}";
  paragraphs[1].dirty = true;
  paragraphs[2].currentText = "乙方：{@partyB}";
  paragraphs[2].dirty = true;
  paragraphs[4].currentText = "{%signature}";
  paragraphs[4].dirty = true;

  // 2b) Save as template (with fieldMeta so fields.json is written)
  const fieldMeta = new Map([
    [
      "partyA",
      {
        type: "text",
        description: "甲方公司名",
        defaultFont: "微软雅黑",
        defaultSize: 12,
        defaultSizeLabel: "小四",
        defaultColor: "#1F4E79",
      },
    ],
    [
      "partyB",
      {
        type: "text",
        description: "乙方公司名",
        defaultFont: "微软雅黑",
        defaultSize: 12,
        defaultSizeLabel: "小四",
        defaultColor: "#1F4E79",
      },
    ],
    ["signature", { type: "image", description: "签名图" }],
  ]);
  const templateBytes = buildTemplate(plainBytes, paragraphs, fieldMeta);
  writeFileSync("smoke-edited-template.docx", templateBytes);

  const templateZip = new PizZip(templateBytes);
  const templateXml = templateZip.file("word/document.xml").asText();
  const templateCT = templateZip.file("[Content_Types].xml").asText();

  check("partyA placeholder injected", templateXml.includes("{@partyA}"));
  check("partyB placeholder injected", templateXml.includes("{@partyB}"));
  check("signature image placeholder injected", templateXml.includes("{%signature}"));
  check("untouched paragraph preserved", templateXml.includes("合同书"));
  check("sectPr preserved through edit", templateXml.includes("<w:sectPr>"));
  check("template/fields.json written", !!templateZip.file("template/fields.json"));
  check(
    "ContentTypes declares json extension",
    /Extension="json"/i.test(templateCT),
  );

  // 2c) Re-load the saved template, detect fields, fill in values
  const fields = extractFieldsFromZip(new PizZip(templateBytes));
  console.log("Re-detected fields:", fields.map((f) => `${f.name}(${f.type})`).join(", "));
  check("re-detected 3 fields", fields.length === 3);

  const imageBytes = readFileSync("src-tauri/icons/128x128.png");
  const values = {
    partyA: { text: "甲方有限公司", font: "微软雅黑", size: 12, color: "#1F4E79" },
    partyB: { text: "乙方有限公司", font: "微软雅黑", size: 12, color: "#1F4E79" },
    signature: { bytes: imageBytes },
  };
  const filled = renderFilled(templateBytes, fields, values);
  writeFileSync("smoke-edited-filled.docx", filled);

  const filledZip = new PizZip(filled);
  const filledXml = filledZip.file("word/document.xml").asText();

  check("filled: partyA replaced", filledXml.includes("甲方有限公司"));
  check("filled: partyB replaced", filledXml.includes("乙方有限公司"));
  check("filled: signature image inserted",
    filledXml.includes("<w:drawing>") &&
      Object.keys(filledZip.files).some((k) => /^word\/media\//.test(k)),
  );
  check("filled: no leftover {@ placeholders", !/\{@\w+\}/.test(filledXml));
  check("filled: no leftover {% placeholders", !/\{%\w+\}/.test(filledXml));
  check(
    "filled: editor-only fields.json stripped",
    !filledZip.file("template/fields.json"),
  );

  // Structural completeness — needed for Word to actually open the file
  check("template: styles.xml preserved", !!templateZip.file("word/styles.xml"));
  check("template: settings.xml preserved", !!templateZip.file("word/settings.xml"));
  check("template: sectPr preserved", templateXml.includes("<w:sectPr>"));
  check(
    "template: w14 namespace declared (mc:Ignorable references it)",
    /xmlns:w14=/.test(templateXml),
  );
  check("filled: styles.xml preserved", !!filledZip.file("word/styles.xml"));
  check("filled: settings.xml preserved", !!filledZip.file("word/settings.xml"));
  check("filled: sectPr preserved", filledXml.includes("<w:sectPr>"));
}

console.log(`\n${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
