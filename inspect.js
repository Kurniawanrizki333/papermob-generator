const ExcelJS = require('exceljs');

async function inspect() {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile('/Users/venturo/Downloads/Tamplate coding 1 (2).xlsx');
  
  const sheet = workbook.worksheets[0];
  
  for (let r = 1; r <= 10; r++) {
    const row = sheet.getRow(r);
    console.log(`Row ${r} - height: ${row.height}`);
    const cells = [];
    for (let c = 1; c <= 10; c++) {
      const cell = sheet.getCell(r, c);
      cells.push({
        col: c,
        val: cell.value,
        fill: cell.fill ? (cell.fill.fgColor ? cell.fill.fgColor.argb || cell.fill.fgColor.theme : 'pattern-no-fg') : 'no-fill',
        width: sheet.getColumn(c).width
      });
    }
    console.log(`  Cells:`, cells.slice(0, 5));
  }
}

inspect().catch(console.error);
