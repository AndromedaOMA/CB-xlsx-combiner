// Exercise the unchanged upload/click flow with a small SheetJS adapter.
// The adapter captures workbook cells/styles; Python separately verifies XLSX serialization.
const fs = require("node:fs");
const vm = require("node:vm");
const assert = require("node:assert/strict");
const path = require("node:path");
const fixtures = JSON.parse(fs.readFileSync(process.argv[2], "utf8"));
const encodeCell = ({r, c}) => {
  let column = "";
  for (let number = c + 1; number; number = Math.floor((number - 1) / 26))
    column = String.fromCharCode(65 + (number - 1) % 26) + column;
  return `${column}${r + 1}`;
};
const elements = new Map();
for (const selector of ["#file-a", "#file-b", "#merge", "#status"]) elements.set(selector, {
  files: [], classList: {remove() {}, add() {}},
  addEventListener(event, callback) { this[event] = callback; }
});
fixtures.forEach((fixture, index) => {
  elements.get(index ? "#file-b" : "#file-a").files = [{
    name: fixture.name,
    async arrayBuffer() { return {SheetNames: Object.keys(fixture.sheets), Sheets: fixture.sheets}; }
  }];
});
let output;
const XLSX = {
  read: value => value,
  utils: {
    sheet_to_json: sheet => sheet,
    encode_cell: encodeCell,
    encode_range: ({s, e}) => `${encodeCell(s)}:${encodeCell(e)}`,
    aoa_to_sheet(rows) {
      const sheet = {rows: rows.map(row => [...row])};
      rows.forEach((row, r) => row.forEach((v, c) => {
        if (v !== null && v !== undefined) sheet[encodeCell({r, c})] = {v};
      }));
      return sheet;
    },
    book_new: () => ({SheetNames: [], Sheets: {}}),
    book_append_sheet(book, sheet, name) { book.SheetNames.push(name); book.Sheets[name] = sheet; }
  },
  writeFile: book => { output = book; }
};
const html = fs.readFileSync(path.join(__dirname, "../index.html"), "utf8");
const script = html.match(/<script>\s*([\s\S]*?)<\/script>/)[1];
vm.runInNewContext(script, {document: {querySelector: selector => elements.get(selector)}, window: {XLSX}, XLSX});
(async () => {
  await elements.get("#merge").click();
  assert.ok(output, elements.get("#status").textContent);
  const sheet = output.Sheets["Merged Hotels"];
  const fills = {};
  sheet.rows.forEach((_, r) => {
    const color = sheet[encodeCell({r, c: 0})]?.s?.fill?.fgColor?.rgb;
    if (color) fills[r] = color;
  });
  process.stdout.write(JSON.stringify({rows: sheet.rows, fills, firstSheet: output.SheetNames[0]}));
})().catch(error => { console.error(error); process.exitCode = 1; });
