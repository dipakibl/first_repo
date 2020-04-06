const path = require('path');
const XLSX = require('xlsx')
const filePath = path.resolve(__dirname, 'OLD MIS.xlsx');
const workbook = XLSX.readFile(filePath);
const sheet_name_list = workbook.SheetNames;
const xlData = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);
const newsheet = XLSX.utils.json_to_sheet(xlData);
const neworkBook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(neworkBook, newsheet);
XLSX.writeFile(neworkBook, "newexcel.xlsx");