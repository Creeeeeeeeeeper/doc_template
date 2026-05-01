// Inspect any docx
import PizZip from "pizzip";
import { readFileSync } from "node:fs";
import { DOMParser } from "xmldom";

const file = process.argv[2] || "smoke-edited-template.docx";
console.log(`\n========== ${file} ==========`);

const buf = readFileSync(file);
const zip = new PizZip(buf);
console.log("Files:");
for (const name of Object.keys(zip.files)) {
  if (zip.files[name].dir) continue;
  console.log(`  ${name}`);
}

console.log("\n[Content_Types].xml:");
console.log(zip.file("[Content_Types].xml")?.asText());

console.log("\nword/_rels/document.xml.rels:");
console.log(zip.file("word/_rels/document.xml.rels")?.asText());

console.log("\nword/document.xml:");
console.log(zip.file("word/document.xml")?.asText());

const errors = [];
const parser = new DOMParser({
  errorHandler: { error: (m) => errors.push(m), fatalError: (m) => errors.push(m) },
});
for (const name of Object.keys(zip.files)) {
  if (zip.files[name].dir) continue;
  if (!name.endsWith(".xml") && !name.endsWith(".rels")) continue;
  errors.length = 0;
  parser.parseFromString(zip.file(name).asText(), "text/xml");
  if (errors.length) console.log(`\nXML errors in ${name}:`, errors);
}
console.log("\n(no XML errors logged above => all parts well-formed)");
