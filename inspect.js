import XLSX from 'xlsx';
import fs from 'fs';

const file = 'Desempenho-individual (4).xls';
const wb = XLSX.readFile(file);

// Check Sheet: Desempenho Individual - Resumo
const wsResumo = wb.Sheets['Desempenho Individual - Resumo'];
const rowsResumo = XLSX.utils.sheet_to_json(wsResumo, { header: 1, defval: '' });
console.log("Resumo sheet, number of students:", rowsResumo.length - 9); // Headers end at row 8

// Check Sheet: Desempenho Individual
const wsIndiv = wb.Sheets['Desempenho Individual'];
const rowsIndiv = XLSX.utils.sheet_to_json(wsIndiv, { header: 1, defval: '' });
console.log("Individual sheet total rows:", rowsIndiv.length);

let currentStudent = null;
let questionsCount = 0;
const studentCounts = {};

for (let i = 4; i < rowsIndiv.length; i++) {
  const row = rowsIndiv[i];
  if (row[0] === 'Aluno' && row[1] === 'Matrícula') {
    if (currentStudent) {
      studentCounts[currentStudent] = questionsCount;
    }
    const infoRow = rowsIndiv[i + 1];
    currentStudent = infoRow ? infoRow[0] : null;
    questionsCount = 0;
    i += 2; // skip Aluno headers and Aluno details
  } else if (row[0] && row[0] !== 'Identificador') {
    questionsCount++;
  }
}
if (currentStudent) {
  studentCounts[currentStudent] = questionsCount;
}

console.log("Questions per student in Desempenho Individual:");
console.log(studentCounts);

// Check Desemp. Ind. por Registro sheet
const wsReg = wb.Sheets['Desemp. Ind. por Registro'];
const rowsReg = XLSX.utils.sheet_to_json(wsReg, { header: 1, defval: '' });
console.log("\nDesemp. Ind. por Registro sheet total rows:", rowsReg.length);
const regCounts = {};
for (let i = 5; i < rowsReg.length; i++) {
  const row = rowsReg[i];
  const student = row[8]; // Aluno
  if (student) {
    regCounts[student] = (regCounts[student] || 0) + 1;
  }
}
console.log("Rows per student in Desemp. Ind. por Registro:");
console.log(regCounts);
