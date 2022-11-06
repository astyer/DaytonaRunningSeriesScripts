"use es6";
// pkg sortResults.js -t host

import * as fs from "fs";
import { read, utils, writeFile, set_fs } from "xlsx/xlsx.mjs";

const [, , inputResultsExcelPath, outputResultsExcelPath] = process.argv;

const fileBuffer = fs.readFileSync(inputResultsExcelPath);
const inputWorkbook = read(fileBuffer);
const inputSheet = inputWorkbook.Sheets[inputWorkbook.SheetNames[0]];
const sheetRange = utils.decode_range(inputSheet["!ref"]);

const outputPath = outputResultsExcelPath || "Results.xlsx";
const outputWorkbook = utils.book_new();
outputWorkbook.title = outputPath.slice(0, -5);

for (let row = 0; row <= sheetRange.e.r; row++) {
  const position = inputSheet[utils.encode_cell({ r: row, c: 0 })].v;
  const name = inputSheet[utils.encode_cell({ r: row, c: 1 })].v;
  const sex = inputSheet[utils.encode_cell({ r: row, c: 2 })].v;
  // console.log("row", row, ":", position, name, sex);
}

set_fs(fs);
writeFile(inputWorkbook, outputPath);
