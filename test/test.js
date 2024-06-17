import * as XLSX from 'xlsx/xlsx.mjs';

/* load 'fs' for readFile and writeFile support */
import * as fs from 'fs';
XLSX.set_fs(fs);

/* load 'stream' for stream support */
import { Readable } from 'stream';
XLSX.stream.set_readable(Readable);

/* load the codepage support library for extended support with older formats  */
import * as cpexcel from 'xlsx/dist/cpexcel.full.mjs';
XLSX.set_cptable(cpexcel);

if(process.argv.length < 3) {
    console.log(`Exiting analysis... need to provide at least one file`)
    process.exit();
}

/*
 *  rows = array of objects with similar formatting
 *  const worksheet = XLSX.utils.json_to_sheet(rows)
 *  const workbook = XLSX.utils.book_new();
 *  XLSX.utils.book_append_sheet(workbook, worksheet, "Name of sheet");
 *  
 */

// default uses the first file provided
const filePath = process.argv[2];
const resultFileName = 'result.xlsx'
console.log(`##### Starting Analysis on ${filePath}`)

var wb = XLSX.readFile(filePath);
var ws = wb.Sheets[wb.SheetNames[0]]

// var range = XLSX.utils.decode_range(ws['!ref']); // get the range
// for(var R = range.s.r; R <= range.e.r; ++R) {
  // for(var C = range.s.c; C <= range.e.c; ++C) {
    // /* find the cell object */
    // var cellref = XLSX.utils.encode_cell({c:C, r:R}); // construct A1 reference for cell
    // if(!ws[cellref]) continue; // if cell doesn't exist, move on
    // var cell = ws[cellref];
    // console.log(cell);
// 
    // /* if the cell is a text cell with the old string, change it */
    // if(!(cell.t == 's' || cell.t == 'str')) continue; // skip if cell is not text
  // }
// }

console.log(ws);
