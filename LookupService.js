/**
 * LookupService - 共用存取工具
 */
const LookupService = (() => {
  function mustSheet(name) {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
    if (!sh) throw new Error(`Sheet not found: ${name}`);
    return sh;
  }

  function headers(sheet) {
    return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  }

  function headerIndex(headerRow) {
    const map = {};
    headerRow.forEach((h, i) => { map[h] = i; });
    return map;
  }

  function readObjects(sheet) {
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    const idx = headerIndex(data[0]);
    return data.slice(1).map((r) => {
      const obj = {};
      Object.keys(idx).forEach((k) => { obj[k] = r[idx[k]]; });
      return obj;
    });
  }

  function enumValues(category) {
    const rows = readObjects(mustSheet('Enums'));
    return rows.filter((r) => r.category === category && r.is_active === true).map((r) => r.value);
  }

  function genId(prefix) {
    return `${prefix}_${Utilities.getUuid().slice(0, 8)}_${Date.now()}`;
  }

  return { mustSheet, headers, headerIndex, readObjects, enumValues, genId };
})();
