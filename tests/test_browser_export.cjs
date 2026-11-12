// Run with: node tests/test_browser_export.cjs path/to/xlsx.bundle.js
// Uses the real browser library, including XLSX serialization and re-reading.
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const assert = require("node:assert/strict");

const libraryPath = process.argv[2] || require.resolve("xlsx-js-style/dist/xlsx.bundle.js");
const context = vm.createContext({console});
vm.runInContext(fs.readFileSync(libraryPath, "utf8"), context, {filename: libraryPath});
const XLSX = context.XLSX;
const elements = new Map();
for (const selector of ["#file-a", "#file-b", "#merge", "#status"]) elements.set(selector, {
  files: [], classList: {remove() {}, add() {}},
  addEventListener(event, callback) { this[event] = callback; }
});

function inputFile(name, sheets) {
  const book = XLSX.utils.book_new();
  for (const [title, rows] of Object.entries(sheets)) {
    XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet(rows), title);
  }
  const bytes = XLSX.write(book, {type: "array", bookType: "xlsx"});
  return {name, async arrayBuffer() { return bytes; }};
}

elements.get("#file-a").files = [inputFile("left.xlsx", {
  Hotels: [["Hotel name", "Stars", "Email"], ["One Hotel", "4*", "sales@one.test"]],
  Empty: [],
  "Blank row": [[]],
  "Null cell": [[null]],
  "Header only": [["Hotel name", "Stars", "Email"]]
})];
elements.get("#file-b").files = [inputFile("right.xlsx", {
  Hotels: [["Hotel name", "Stars", "Email"], ["One Hotel", "4*", "groups@one.test"]]
})];

let output, serialized, filename;
XLSX.writeFile = (book, name, options) => {
  output = book;
  serialized = XLSX.write(book, {...options, type: "array"});
  filename = name;
};
context.window = context;
context.document = {querySelector: selector => elements.get(selector)};
const html = fs.readFileSync(path.join(__dirname, "../index.html"), "utf8");
vm.runInContext(html.match(/<script>\s*([\s\S]*?)<\/script>/)[1], context);

(async () => {
  await elements.get("#merge").click();
  assert.ok(serialized, elements.get("#status").textContent);
  assert.equal(filename, "left_merged.xlsx");
  const reopened = XLSX.read(serialized, {type: "array"});
  for (const title of ["Empty", "Blank row", "Null cell"]) {
    const name = `Source 1 - left.xlsx - ${title}`.slice(0, 31);
    assert.ok(output.Sheets[name], `Empty source tab was dropped: ${name}`);
    assert.equal(output.Sheets[name]["!ref"], undefined);
    assert.equal(output.Sheets[name]["!autofilter"], undefined);
    assert.ok(reopened.SheetNames.includes(name));
    assert.equal(reopened.Sheets[name]["!autofilter"], undefined);
  }
  // A header with no data still has a valid range and keeps its filter.
  const headerName = "Source 1 - left.xlsx - Header only".slice(0, 31);
  assert.equal(output.Sheets[headerName]["!autofilter"].ref, "A1:C1");
  assert.equal(reopened.Sheets["Merged Hotels"]["!autofilter"].ref, "A1:G4");
  assert.equal(reopened.Sheets["Merged Hotels"].A2.v, "One Hotel");
  assert.equal(reopened.Sheets["Merged Hotels"].B2.v, "4*");
  assert.equal(reopened.Sheets["Merged Hotels"].C2.v, "sales@one.test");
  assert.equal(reopened.Sheets["Merged Hotels"].D2.v, "groups@one.test");
  assert.match(elements.get("#status").textContent, /^Downloaded /);
  assert.equal(elements.get("#merge").disabled, false);
  process.stdout.write("Browser XLSX export regression passed (empty tabs, headers, values and filters).\n");
})().catch(error => { console.error(error); process.exitCode = 1; });
