"use es6";
// pkg sortResults.js -t host

// import * as fs from "fs";
const XLSX = require("xlsx");

const [, , inputResultsExcelPath, outputResultsExcelPath] = process.argv;

const inputWorkbook = XLSX.readFile(inputResultsExcelPath);
const inputSheet = inputWorkbook.Sheets[inputWorkbook.SheetNames[0]];
const sheetRange = XLSX.utils.decode_range(inputSheet["!ref"]);

const outputPath = outputResultsExcelPath || "Results.xlsx";
const outputWorkbook = XLSX.utils.book_new();
outputWorkbook.title = outputPath.slice(0, -5);

const resultSheets = {
  overall: {
    male: { name: "maleOverall", data: [] },
    female: { name: "femaleOverall", data: [] },
  },
  fortyPlus: {
    male: { data: [], name: "male40+" },
    female: { data: [], name: "female40+" },
  },
  fiftyPlus: {
    male: { data: [], name: "male50+" },
    female: { data: [], name: "female50+" },
  },
  sixtyPlus: {
    male: { data: [], name: "male60+" },
    female: { data: [], name: "female60+" },
  },
  under10: {
    male: { data: [], name: "maleUnder10" },
    female: { data: [], name: "femaleUnder10" },
  },
  under15: {
    male: { data: [], name: "male10-14" },
    female: { data: [], name: "female10-14" },
  },
  under20: {
    male: { data: [], name: "male15-19" },
    female: { data: [], name: "female15-19" },
  },
  under25: {
    male: { data: [], name: "male20-24" },
    female: { data: [], name: "female20-24" },
  },
  under30: {
    male: { data: [], name: "male25-29" },
    female: { data: [], name: "female25-29" },
  },
  under35: {
    male: { data: [], name: "male30-34" },
    female: { data: [], name: "female30-34" },
  },
  under40: {
    male: { data: [], name: "male35-39" },
    female: { data: [], name: "female35-39" },
  },
  under45: {
    male: { data: [], name: "male40-44" },
    female: { data: [], name: "female40-44" },
  },
  under50: {
    male: { data: [], name: "male45-49" },
    female: { data: [], name: "female45-49" },
  },
  under55: {
    male: { data: [], name: "male50-54" },
    female: { data: [], name: "female50-54" },
  },
  under60: {
    male: { data: [], name: "male55-59" },
    female: { data: [], name: "female55-59" },
  },
  under65: {
    male: { data: [], name: "male60-64" },
    female: { data: [], name: "female60-64" },
  },
  under70: {
    male: { data: [], name: "male65-69" },
    female: { data: [], name: "female65-69" },
  },
  under75: {
    male: { data: [], name: "male70-74" },
    female: { data: [], name: "female70-74" },
  },
  under80: {
    male: { data: [], name: "male75-79" },
    female: { data: [], name: "female75-79" },
  },
  under85: {
    male: { data: [], name: "male80-84" },
    female: { data: [], name: "female80-84" },
  },
  under90: {
    male: { data: [], name: "male85-89" },
    female: { data: [], name: "female85-89" },
  },
  over89: {
    male: { data: [], name: "male90+" },
    female: { data: [], name: "female90+" },
  },
};

for (let row = 0; row <= sheetRange.e.r; row++) {
  const position = inputSheet[XLSX.utils.encode_cell({ r: row, c: 0 })].v;
  const name = inputSheet[XLSX.utils.encode_cell({ r: row, c: 1 })].v;
  const sex = inputSheet[XLSX.utils.encode_cell({ r: row, c: 2 })].v;
  const age = inputSheet[XLSX.utils.encode_cell({ r: row, c: 3 })].v;
  const rowData = [position, name, sex, age];

  const sexKey = sex == "M" ? "male" : "female";

  // overall results
  resultSheets.overall[sexKey].data.push(rowData);

  // masters points earners
  if (age >= 40) {
    resultSheets.fortyPlus[sexKey].data.push(rowData);
  }
  if (age >= 50) {
    resultSheets.fiftyPlus[sexKey].data.push(rowData);
  }
  if (age >= 60) {
    resultSheets.sixtyPlus[sexKey].data.push(rowData);
  }

  // age groups
  if (age < 10) {
    resultSheets.under10[sexKey].data.push(rowData);
  } else if (age < 15) {
    resultSheets.under15[sexKey].data.push(rowData);
  } else if (age < 20) {
    resultSheets.under20[sexKey].data.push(rowData);
  } else if (age < 25) {
    resultSheets.under25[sexKey].data.push(rowData);
  } else if (age < 30) {
    resultSheets.under30[sexKey].data.push(rowData);
  } else if (age < 35) {
    resultSheets.under35[sexKey].data.push(rowData);
  } else if (age < 40) {
    resultSheets.under40[sexKey].data.push(rowData);
  } else if (age < 45) {
    resultSheets.under45[sexKey].data.push(rowData);
  } else if (age < 50) {
    resultSheets.under50[sexKey].data.push(rowData);
  } else if (age < 55) {
    resultSheets.under55[sexKey].data.push(rowData);
  } else if (age < 60) {
    resultSheets.under60[sexKey].data.push(rowData);
  } else if (age < 65) {
    resultSheets.under65[sexKey].data.push(rowData);
  } else if (age < 70) {
    resultSheets.under70[sexKey].data.push(rowData);
  } else if (age < 75) {
    resultSheets.under75[sexKey].data.push(rowData);
  } else if (age < 80) {
    resultSheets.under80[sexKey].data.push(rowData);
  } else if (age < 85) {
    resultSheets.under85[sexKey].data.push(rowData);
  } else if (age < 90) {
    resultSheets.under90[sexKey].data.push(rowData);
  } else {
    resultSheets.over89[sexKey].data.push(rowData);
  }
}

for (let sheet of Object.values(resultSheets)) {
  outputWorkbook.SheetNames.push(sheet.male.name);
  outputWorkbook.SheetNames.push(sheet.female.name);
  outputWorkbook.Sheets[sheet.male.name] = XLSX.utils.aoa_to_sheet(sheet.male.data);
  outputWorkbook.Sheets[sheet.female.name] = XLSX.utils.aoa_to_sheet(sheet.female.data);
}

XLSX.writeFile(outputWorkbook, outputPath);
