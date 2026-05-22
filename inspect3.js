import XLSX from 'xlsx';
import fs from 'fs';

const file = 'Desempenho-individual (4).xls';
const wb = XLSX.readFile(file);

// Exact replica of processarDados from App.jsx
const processarDados = (wb) => {
  let resumoRows = null;
  let registroRows = null;
  let individualRows = null;

  wb.SheetNames.forEach((name) => {
    const sheet = wb.Sheets[name];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
    if (!rows || rows.length === 0) return;

    let hasIdentificador = false;
    let hasMarcacao = false;
    let hasGabarito = false;
    let hasDesempenhoAvaliacao = false;
    let hasAlunoMatriculaHeader = false;
    let totalAlunoCountFirstCol = 0;

    rows.forEach((row, idx) => {
      if (!row) return;
      if (row[0] === 'Aluno') totalAlunoCountFirstCol++;

      if (idx < 25) {
        const rowStr = row.map(cell => (cell || '').toString().toLowerCase());
        if (rowStr.includes('identificador')) hasIdentificador = true;
        if (rowStr.includes('marcação respondente') || rowStr.includes('marcaçao respondente') || rowStr.includes('marcacao respondente')) {
          hasMarcacao = true;
        }
        if (rowStr.includes('gabarito')) hasGabarito = true;
        if (rowStr.includes('desempenho na avaliação') || rowStr.includes('desempenho na avaliacao')) {
          hasDesempenhoAvaliacao = true;
        }
        if (rowStr.includes('aluno') && rowStr.includes('matrícula')) {
          hasAlunoMatriculaHeader = true;
        }
      }
    });

    if (hasMarcacao && hasGabarito && hasIdentificador && totalAlunoCountFirstCol <= 2) {
      registroRows = rows;
    } else if (totalAlunoCountFirstCol > 2 && hasIdentificador) {
      individualRows = rows;
    } else if (hasDesempenhoAvaliacao || (hasAlunoMatriculaHeader && !hasIdentificador)) {
      resumoRows = rows;
    }
  });

  if (!resumoRows) {
    const ws = wb.Sheets[wb.SheetNames[0]];
    resumoRows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  }
  if (!registroRows) {
    const abaRegistroName = wb.SheetNames.find((nome) => nome.includes('por Registro'));
    if (abaRegistroName) {
      const wsRegistro = wb.Sheets[abaRegistroName];
      registroRows = XLSX.utils.sheet_to_json(wsRegistro, { header: 1, defval: '' });
    }
  }
  if (!individualRows) {
    const abaIndividualName = wb.SheetNames.find((nome) => nome === 'Desempenho Individual' || nome.includes('Individual'));
    if (abaIndividualName) {
      const wsIndividual = wb.Sheets[abaIndividualName];
      individualRows = XLSX.utils.sheet_to_json(wsIndividual, { header: 1, defval: '' });
    }
  }

  const titulo = resumoRows[0] ? resumoRows[0][0] : 'ESTATÍSTICA DA APLICAÇÃO';
  const subtitulo = resumoRows[1] ? resumoRows[1][0] : 'Desempenho Individual';
  
  let statsRowIndex = 4;
  for (let i = 0; i < Math.min(10, resumoRows.length); i++) {
    const row = resumoRows[i];
    if (row && (row.includes('Desempenho na avaliação') || row.includes('Desempenho na avaliacao'))) {
      statsRowIndex = i - 1;
      break;
    }
  }
  const stats = resumoRows[statsRowIndex] || [];
  const desempenho = parseFloat((stats[0] || '0').toString().replace('%', '').replace(',', '.'));
  const acertosMedio = parseFloat((stats[1] || '0').toString().replace(',', '.'));
  const errosMedio = parseFloat((stats[2] || '0').toString().replace(',', '.'));
  const totalItens = parseInt(stats[3] || 0);
  const respondentes = parseInt(stats[4] || 0);
  
  const labelsStats = resumoRows[statsRowIndex + 1] || [];
  
  let headersRowIndex = 8;
  for (let i = 0; i < Math.min(12, resumoRows.length); i++) {
    const row = resumoRows[i];
    if (row && row[0] === 'Aluno' && row[1] === 'Matrícula') {
      headersRowIndex = i;
      break;
    }
  }
  const headers = resumoRows[headersRowIndex] || ['Aluno', 'Matrícula', 'Desempenho', 'Acertos', 'Erros', 'Tempo de realização'];
  
  const inicioAlunos = headersRowIndex + 1;
  const alunos = [];
  for (let i = inicioAlunos; i < resumoRows.length; i++) {
    const row = resumoRows[i];
    if (!row || !row[0] || typeof row[0] !== 'string') continue;
    
    const nome = row[0].trim();
    if (nome === '' || nome === 'Aluno') continue;
    
    const matricula = row[1];
    const alunoDesempenho = parseFloat((row[2] || '0').toString().replace('%', '').replace(',', '.'));
    
    alunos.push({
      nome,
      matricula: (matricula || '').toString().trim(),
      desempenho: alunoDesempenho,
      acertos: parseInt(row[3] || 0),
      erros: parseInt(row[4] || 0),
      tempo: row[5] || '',
      questoes: []
    });
  }

  let curso = '';
  let turmaName = '';
  let dataProva = '';
  let periodo = '';
  let escola = '';
  let turno = '';

  for (let r = 0; r < Math.min(6, resumoRows.length); r++) {
    const row = resumoRows[r];
    if (row) {
      row.forEach(cell => {
        if (cell) {
          const str = cell.toString().trim();
          const lower = str.toLowerCase();
          if (lower.startsWith('curso:')) {
            curso = str.substring(6).trim();
          } else if (lower.startsWith('turma:')) {
            turmaName = str.substring(6).trim();
          } else if (lower.startsWith('avaliação:') || lower.startsWith('avaliacao:')) {
            dataProva = str.substring(10).trim();
          } else if (lower.startsWith('período:') || lower.startsWith('periodo:')) {
            periodo = str.substring(8).trim();
          } else if (lower.startsWith('escola:')) {
            escola = str.substring(7).trim();
          } else if (lower.startsWith('turno:')) {
            turno = str.substring(6).trim();
          }
        }
      });
    }
  }

  const detalhesAlunos = {};
  let possuiAbaRegistro = false;

  if (registroRows && registroRows.length > 4) {
    let headerRowIndex = 4;
    for (let i = 0; i < Math.min(12, registroRows.length); i++) {
      if (registroRows[i] && registroRows[i].includes('Identificador')) {
        headerRowIndex = i;
        break;
      }
    }
    
    const headerRow = registroRows[headerRowIndex];
    const colIdx = {
      aluno: headerRow.findIndex(c => c && c.toString().toLowerCase() === 'aluno'),
      matricula: headerRow.findIndex(c => c && (c.toString().toLowerCase() === 'matrícula' || c.toString().toLowerCase() === 'matricula')),
      identificador: headerRow.findIndex(c => c && c.toString().toLowerCase() === 'identificador'),
      capacidade: headerRow.findIndex(c => c && c.toString().toLowerCase().includes('capacidade')),
      subfuncao: headerRow.findIndex(c => c && (c.toString().toLowerCase().includes('subfunção') || c.toString().toLowerCase().includes('subfuncao'))),
      padraoDesempenho: headerRow.findIndex(c => c && (c.toString().toLowerCase().includes('padrão de desempenho') || c.toString().toLowerCase().includes('padrao de desempenho'))),
      conhecimento: headerRow.findIndex(c => c && c.toString().toLowerCase().includes('conhecimento')),
      dificuldade: headerRow.findIndex(c => c && c.toString().toLowerCase().includes('dificuldade')),
      marcacao: headerRow.findIndex(c => c && (c.toString().toLowerCase().includes('marcação respondente') || c.toString().toLowerCase().includes('marcaçao respondente') || c.toString().toLowerCase().includes('marcacao respondente'))),
      gabarito: headerRow.findIndex(c => c && c.toString().toLowerCase() === 'gabarito'),
      curso: headerRow.findIndex(c => c && c.toString().toLowerCase() === 'curso'),
      turma: headerRow.findIndex(c => c && c.toString().toLowerCase() === 'turma'),
      avaliacao: headerRow.findIndex(c => c && (c.toString().toLowerCase() === 'avaliação' || c.toString().toLowerCase() === 'avaliacao')),
      escolaCol: headerRow.findIndex(c => c && c.toString().toLowerCase().includes('escola')),
      turnoCol: headerRow.findIndex(c => c && c.toString().toLowerCase().includes('turno')),
      periodoCol: headerRow.findIndex(c => c && (c.toString().toLowerCase().includes('período de agendamento') || c.toString().toLowerCase().includes('periodo de agendamento')))
    };

    if (colIdx.identificador >= 0 && colIdx.matricula >= 0) {
      possuiAbaRegistro = true;
      const primeiraLinhaDados = registroRows[headerRowIndex + 1];
      if (primeiraLinhaDados) {
        if (colIdx.turma >= 0 && primeiraLinhaDados[colIdx.turma]) turmaName = primeiraLinhaDados[colIdx.turma].toString().trim();
        if (colIdx.curso >= 0 && primeiraLinhaDados[colIdx.curso]) curso = primeiraLinhaDados[colIdx.curso].toString().trim();
        if (colIdx.avaliacao >= 0 && primeiraLinhaDados[colIdx.avaliacao]) dataProva = primeiraLinhaDados[colIdx.avaliacao].toString().trim();
        if (colIdx.escolaCol >= 0 && primeiraLinhaDados[colIdx.escolaCol]) escola = primeiraLinhaDados[colIdx.escolaCol].toString().trim();
        if (colIdx.turnoCol >= 0 && primeiraLinhaDados[colIdx.turnoCol]) turno = primeiraLinhaDados[colIdx.turnoCol].toString().trim();
        if (colIdx.periodoCol >= 0 && primeiraLinhaDados[colIdx.periodoCol]) periodo = primeiraLinhaDados[colIdx.periodoCol].toString().trim();
      }

      for (let i = headerRowIndex + 1; i < registroRows.length; i++) {
        const row = registroRows[i];
        if (!row || row.length <= Math.max(colIdx.matricula, colIdx.identificador)) continue;
        
        const matriculaAluno = row[colIdx.matricula];
        if (!matriculaAluno) continue;
        
        const chave = `${matriculaAluno.toString().trim()}`;
        if (!detalhesAlunos[chave]) {
          detalhesAlunos[chave] = [];
        }
        
        const marcacao = colIdx.marcacao >= 0 ? row[colIdx.marcacao].toString().trim() : '';
        const gabarito = colIdx.gabarito >= 0 ? row[colIdx.gabarito].toString().trim() : '';
        const acertou = marcacao === gabarito;
        
        detalhesAlunos[chave].push({
          identificador: colIdx.identificador >= 0 ? row[colIdx.identificador] : '',
          capacidade: colIdx.capacidade >= 0 ? row[colIdx.capacidade] : '',
          subfuncao: colIdx.subfuncao >= 0 ? row[colIdx.subfuncao] : '',
          padraoDesempenho: colIdx.padraoDesempenho >= 0 ? row[colIdx.padraoDesempenho] : '',
          conhecimento: colIdx.conhecimento >= 0 ? row[colIdx.conhecimento] : '',
          dificuldade: colIdx.dificuldade >= 0 ? row[colIdx.dificuldade] : '',
          marcacao: marcacao,
          gabarito: gabarito,
          acertou: acertou
        });
      }
    }
  }

  if (Object.keys(detalhesAlunos).length === 0 && individualRows) {
    let i = 0;
    while (i < individualRows.length) {
      const row = individualRows[i];
      if (row && row[0] === 'Aluno' && row[1] === 'Matrícula') {
        const infoRow = individualRows[i + 1];
        if (infoRow && infoRow[0] && infoRow[0] !== 'Aluno') {
          const matricula = infoRow[1];
          const chave = `${matricula.toString().trim()}`;
          
          const questList = [];
          let qIdx = i + 3;
          while (qIdx < individualRows.length) {
            const qRow = individualRows[qIdx];
            if (!qRow || qRow[0] === 'Aluno' || qRow[0] === '' || qRow.length === 0) {
              break;
            }
            if (qRow[0] && qRow[0] !== 'Identificador') {
              const marcacao = qRow[6] ? qRow[6].toString().trim() : '';
              const gabarito = qRow[7] ? qRow[7].toString().trim() : '';
              const acertou = marcacao === gabarito;
              questList.push({
                identificador: qRow[0],
                capacidade: qRow[1],
                subfuncao: qRow[2],
                padraoDesempenho: qRow[3],
                conhecimento: qRow[4],
                dificuldade: qRow[5],
                marcacao: marcacao,
                gabarito: gabarito,
                acertou: acertou
              });
            }
            qIdx++;
          }
          if (questList.length > 0) {
            detalhesAlunos[chave] = questList;
            possuiAbaRegistro = true;
          }
          i = qIdx - 1;
        }
      }
      i++;
    }
  }

  alunos.forEach((aluno) => {
    const chave = `${aluno.matricula}`;
    aluno.questoes = detalhesAlunos[chave] || [];
    if (aluno.questoes.length > 0) {
      aluno.acertos = aluno.questoes.filter(q => q.acertou).length;
      aluno.erros = aluno.questoes.filter(q => !q.acertou).length;
      aluno.desempenho = parseFloat(((aluno.acertos / aluno.questoes.length) * 100).toFixed(2));
    }
  });

  let maxQuestoes = 0;
  alunos.forEach(a => {
    if (a.questoes && a.questoes.length > maxQuestoes) {
      maxQuestoes = a.questoes.length;
    }
  });
  const finalTotalItens = Math.max(totalItens, maxQuestoes);

  return {
    titulo,
    totalItens: finalTotalItens,
    alunosCount: alunos.length,
    possuiAbaRegistro,
    totalQuestoesDetalhadas: Object.values(detalhesAlunos).reduce((sum, list) => sum + list.length, 0),
    primeiroAlunoQuestoesCount: alunos[0] ? alunos[0].questoes.length : 0,
    alunosSemQuestoes: alunos.filter(a => a.questoes.length === 0).map(a => a.nome)
  };
};

console.log(processarDados(wb));
