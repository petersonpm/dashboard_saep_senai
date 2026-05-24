import React, { useState, useRef, useMemo, useEffect } from 'react';
import * as XLSX from 'xlsx';
import { supabase, isSupabaseConfigured } from './supabaseClient';
import { 
  UploadCloud, 
  Award, 
  Clock, 
  Users, 
  Percent, 
  BarChart3, 
  Users2, 
  Search, 
  ChevronDown, 
  ChevronUp, 
  FileSpreadsheet, 
  Sparkles,
  CheckCircle2,
  XCircle,
  HelpCircle,
  TrendingUp,
  X,
  Target,
  Trophy,
  ArrowRight,
  Download,
  AlertCircle,
  RefreshCw,
  Settings
} from 'lucide-react';
import {
  BarChart,
  Bar,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  ResponsiveContainer,
  ScatterChart,
  Scatter,
  ZAxis,
  Cell
} from 'recharts';

export default function App() {
  const [turmas, setTurmas] = useState({});
  const [selectedTurma, setSelectedTurma] = useState(null);
  const [activeSubTab, setActiveSubTab] = useState('resumo'); // 'resumo' | 'individual' | 'registro'
  const [isDragging, setIsDragging] = useState(false);
  const [searchQuery, setSearchQuery] = useState('');
  const [filterLevel, setFilterLevel] = useState('todos'); // 'todos' | 'excelente' | 'bom' | 'regular' | 'insuficiente'
  const [expandedAluno, setExpandedAluno] = useState(null); // 'matricula'
  const [showHelpModal, setShowHelpModal] = useState(false);

  // Estados adicionados para a reestruturação das abas
  const [selectedAlunoMatricula, setSelectedAlunoMatricula] = useState(null);
  const [alunoQuestaoFilter, setAlunoQuestaoFilter] = useState('todas'); // 'todas' | 'acertos' | 'erros'
  const [registroSearch, setRegistroSearch] = useState('');
  const [registroPage, setRegistroPage] = useState(1);
  const [copiedPrompt, setCopiedPrompt] = useState(false);

  // Supabase Authentication & Cloud States
  const [session, setSession] = useState(null);
  const [authLoading, setAuthLoading] = useState(true);
  const [authError, setAuthError] = useState(null);
  const [isSignUpMode, setIsSignUpMode] = useState(false);
  const [isSavingToCloud, setIsSavingToCloud] = useState(false);

  // Auto-sync on startup if valid credentials exist
  useEffect(() => {
    // Obter sessão inicial e monitorar mudanças de autenticação
    if (isSupabaseConfigured) {
      supabase.auth.getSession().then(({ data: { session } }) => {
        setSession(session);
        setAuthLoading(false);
        if (session?.user) {
          loadUserTurmas(session.user.id);
        }
      });

      const { data: { subscription } } = supabase.auth.onAuthStateChange((_event, session) => {
        setSession(session);
        if (session?.user) {
          loadUserTurmas(session.user.id);
        } else {
          setTurmas({});
          setSelectedTurma(null);
        }
        setAuthLoading(false);
      });

      return () => subscription.unsubscribe();
    } else {
      setAuthLoading(false);
    }
  }, []);

  const loadUserTurmas = async (userId) => {
    if (!isSupabaseConfigured) return;
    try {
      const { data, error } = await supabase
        .from('turmas')
        .select('*')
        .order('created_at', { ascending: true });
      
      if (error) throw error;
      
      const loadedTurmas = {};
      data.forEach(row => {
        loadedTurmas[row.nome_key] = {
          ...row.dados,
          id_db: row.id
        };
      });
      setTurmas(loadedTurmas);
      
      const keys = Object.keys(loadedTurmas);
      if (keys.length > 0) {
        setSelectedTurma(keys[keys.length - 1]);
        setSelectedAlunoMatricula(loadedTurmas[keys[keys.length - 1]]?.alunos[0]?.matricula || null);
      }
    } catch (err) {
      console.error("Erro ao carregar turmas:", err.message);
    }
  };

  const saveTurmaToSupabase = async (nomeKey, dados) => {
    if (!isSupabaseConfigured) return;
    setIsSavingToCloud(true);
    try {
      // Obter usuário atual
      const { data: { user } } = await supabase.auth.getUser();
      if (!user) return;

      const { data, error } = await supabase
        .from('turmas')
        .insert([{
          user_id: user.id,
          nome_key: nomeKey,
          dados: dados
        }])
        .select();
      
      if (error) throw error;
      
      if (data && data[0]) {
        setTurmas(prev => ({
          ...prev,
          [nomeKey]: { ...dados, id_db: data[0].id }
        }));
      }
    } catch (err) {
      console.error("Erro ao salvar no Supabase:", err.message);
    } finally {
      setIsSavingToCloud(false);
    }
  };

  const deleteTurmaFromSupabase = async (nomeKey) => {
    const target = turmas[nomeKey];
    if (!target) return;
    
    if (isSupabaseConfigured && target.id_db) {
      try {
        const { error } = await supabase
          .from('turmas')
          .delete()
          .eq('id', target.id_db);
        
        if (error) throw error;
      } catch (err) {
        console.error("Erro ao excluir do Supabase:", err.message);
      }
    }
  };

  const fileInputRef = useRef(null);

  // Drag and Drop handlers
  const handleDragOver = (e) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const handleDragLeave = () => {
    setIsDragging(false);
  };

  const handleDrop = (e) => {
    e.preventDefault();
    setIsDragging(false);
    if (e.dataTransfer.files.length > 0) {
      processFiles(e.dataTransfer.files);
    }
  };

  const handleFileChange = (e) => {
    if (e.target.files.length > 0) {
      processFiles(e.target.files);
    }
  };

  const processFiles = async (files) => {
    const validFiles = Array.from(files).filter(
      (f) => f.name.endsWith('.xls') || f.name.endsWith('.xlsx')
    );

    if (validFiles.length === 0) return;

    const novasTurmas = { ...turmas };

    for (const file of validFiles) {
      const nomeSemExt = file.name.replace(/\.xlsx?$/, '');
      const dados = await lerPlanilha(file);
      novasTurmas[nomeSemExt] = dados;
      
      if (isSupabaseConfigured && session?.user) {
        await saveTurmaToSupabase(nomeSemExt, dados);
      }
    }

    if (!isSupabaseConfigured || !session?.user) {
      setTurmas(novasTurmas);
      const keys = Object.keys(novasTurmas);
      if (keys.length > 0) {
        const ultimaTurma = keys[keys.length - 1];
        setSelectedTurma(ultimaTurma); // Select latest loaded
        const primeiroAluno = novasTurmas[ultimaTurma]?.alunos[0]?.matricula || null;
        setSelectedAlunoMatricula(primeiroAluno);
      }
    } else {
      const keys = Object.keys(novasTurmas);
      if (keys.length > 0) {
        const ultimaTurma = keys[keys.length - 1];
        setSelectedTurma(ultimaTurma);
        const primeiroAluno = novasTurmas[ultimaTurma]?.alunos[0]?.matricula || null;
        setSelectedAlunoMatricula(primeiroAluno);
      }
    }
  };

  const formatCapacidade = (val) => {
    if (!val) return '';
    const str = val.toString().trim();
    if (/^\d+$/.test(str)) {
      return `C${str}`;
    }
    return str;
  };

  const lerPlanilha = (arquivo) => {
    return new Promise((resolve) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, { type: 'array' });
        resolve(processarDados(wb));
      };
      reader.readAsArrayBuffer(arquivo);
    });
  };

  const processarDados = (wb) => {
    // Content-based sheet classifier
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
        if (row[0] && row[0].toString().toLowerCase().trim() === 'aluno') totalAlunoCountFirstCol++;

        if (idx < 25) {
          const rowStr = row.map(cell => (cell || '').toString().toLowerCase().trim());
          if (rowStr.some(c => c === 'identificador')) hasIdentificador = true;
          if (rowStr.some(c => c.includes('marcação respondente') || c.includes('marcaçao respondente') || c.includes('marcacao respondente'))) {
            hasMarcacao = true;
          }
          if (rowStr.some(c => c === 'gabarito')) hasGabarito = true;
          if (rowStr.some(c => c === 'desempenho na avaliação' || c === 'desempenho na avaliacao')) {
            hasDesempenhoAvaliacao = true;
          }
          if (rowStr.some(c => c === 'aluno') && rowStr.some(c => c === 'matrícula' || c === 'matricula')) {
            hasAlunoMatriculaHeader = true;
          }
        }
      });

      // Classification
      if (hasMarcacao && hasGabarito && hasIdentificador && totalAlunoCountFirstCol <= 2) {
        registroRows = rows;
      } else if (totalAlunoCountFirstCol > 2 && hasIdentificador) {
        individualRows = rows;
      } else if (hasDesempenhoAvaliacao || (hasAlunoMatriculaHeader && !hasIdentificador)) {
        resumoRows = rows;
      }
    });

    // Fallback: If Resumo sheet was not classified, use the first sheet
    if (!resumoRows) {
      const ws = wb.Sheets[wb.SheetNames[0]];
      resumoRows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    }

    // Fallback: If Registro sheet was not classified, try name-based search
    if (!registroRows) {
      const abaRegistroName = wb.SheetNames.find((nome) => nome.includes('por Registro'));
      if (abaRegistroName) {
        const wsRegistro = wb.Sheets[abaRegistroName];
        registroRows = XLSX.utils.sheet_to_json(wsRegistro, { header: 1, defval: '' });
      }
    }

    // Fallback: If Individual sheet was not classified, try name-based search
    if (!individualRows) {
      const abaIndividualName = wb.SheetNames.find((nome) => nome === 'Desempenho Individual' || nome.includes('Individual'));
      if (abaIndividualName) {
        const wsIndividual = wb.Sheets[abaIndividualName];
        individualRows = XLSX.utils.sheet_to_json(wsIndividual, { header: 1, defval: '' });
      }
    }

    // Row 0: Title
    const titulo = resumoRows[0] ? resumoRows[0][0] : 'ESTATÍSTICA DA APLICAÇÃO';
    
    // Row 1: Subtitle
    const subtitulo = resumoRows[1] ? resumoRows[1][0] : 'Desempenho Individual';
    
    // Scan stats dynamically
    let statsRowIndex = 4;
    for (let i = 0; i < Math.min(10, resumoRows.length); i++) {
      const row = resumoRows[i];
      if (row && row.some(cell => cell && (cell.toString().toLowerCase().trim() === 'desempenho na avaliação' || cell.toString().toLowerCase().trim() === 'desempenho na avaliacao'))) {
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
    
    // Find headers row index
    let headersRowIndex = 8;
    for (let i = 0; i < Math.min(12, resumoRows.length); i++) {
      const row = resumoRows[i];
      if (row && row[0] && row[1] && 
          row[0].toString().toLowerCase().trim() === 'aluno' && 
          (row[1].toString().toLowerCase().trim() === 'matrícula' || row[1].toString().toLowerCase().trim() === 'matricula')) {
        headersRowIndex = i;
        break;
      }
    }
    const headers = resumoRows[headersRowIndex] || ['Aluno', 'Matrícula', 'Desempenho', 'Acertos', 'Erros', 'Tempo de realização'];
    
    const inicioAlunos = headersRowIndex + 1;
    
    // Parse students list
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

    // Metadata scanning from Resumo sheet cells
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

    // Process student detailed responses
    const detalhesAlunos = {};
    let possuiAbaRegistro = false;

    // Build question lookup database from individualRows if available
    const questionPedagogicalMap = {};
    if (individualRows) {
      let i = 0;
      while (i < individualRows.length) {
        const row = individualRows[i];
        if (row && row[0] && row[1] && 
            row[0].toString().toLowerCase().trim() === 'aluno' && 
            (row[1].toString().toLowerCase().trim() === 'matrícula' || row[1].toString().toLowerCase().trim() === 'matricula')) {
          let qIdx = i + 3;
          while (qIdx < individualRows.length) {
            const qRow = individualRows[qIdx];
            if (!qRow || qRow[0] === 'Aluno' || qRow[0] === '' || qRow.length === 0) {
              break;
            }
            if (qRow[0] && qRow[0] !== 'Identificador') {
              const id = qRow[0].toString().trim();
              if (!questionPedagogicalMap[id]) {
                questionPedagogicalMap[id] = {
                  capacidade: qRow[1] ? qRow[1].toString().trim() : '',
                  subfuncao: qRow[2] ? qRow[2].toString().trim() : '',
                  padraoDesempenho: qRow[3] ? qRow[3].toString().trim() : '',
                  conhecimento: qRow[4] ? qRow[4].toString().trim() : '',
                  dificuldade: qRow[5] ? qRow[5].toString().trim() : ''
                };
              }
            }
            qIdx++;
          }
          if (Object.keys(questionPedagogicalMap).length > 0) {
            break; // We got all questions from the first student block, that's enough
          }
        }
        i++;
      }
    }

    // Source A: Try parsing from 'por Registro' sheet
    if (registroRows && registroRows.length > 4) {
      // Find header row (usually contains 'Identificador')
      let headerRowIndex = 4;
      for (let i = 0; i < Math.min(12, registroRows.length); i++) {
        if (registroRows[i] && registroRows[i].some(cell => cell && cell.toString().toLowerCase().trim() === 'identificador')) {
          headerRowIndex = i;
          break;
        }
      }
      
      const headerRow = registroRows[headerRowIndex];
      const colIdx = {
        aluno: headerRow.findIndex(c => c && c.toString().toLowerCase().trim() === 'aluno'),
        matricula: headerRow.findIndex(c => c && (c.toString().toLowerCase().trim() === 'matrícula' || c.toString().toLowerCase().trim() === 'matricula')),
        identificador: headerRow.findIndex(c => c && c.toString().toLowerCase().trim() === 'identificador'),
        capacidade: headerRow.findIndex(c => c && c.toString().toLowerCase().includes('capacidade')),
        subfuncao: headerRow.findIndex(c => c && (c.toString().toLowerCase().includes('subfunção') || c.toString().toLowerCase().includes('subfuncao') || c.toString().toLowerCase().includes('habilidade'))),
        padraoDesempenho: headerRow.findIndex(c => c && (c.toString().toLowerCase().includes('padrão de desempenho') || c.toString().toLowerCase().includes('padrao de desempenho') || c.toString().toLowerCase().includes('descritor') || c.toString().toLowerCase().includes('padrão') || c.toString().toLowerCase().includes('padrao'))),
        conhecimento: headerRow.findIndex(c => c && (c.toString().toLowerCase().includes('conhecimento') || c.toString().toLowerCase().includes('assunto') || c.toString().toLowerCase().includes('conteúdo') || c.toString().toLowerCase().includes('conteudo'))),
        dificuldade: headerRow.findIndex(c => c && c.toString().toLowerCase().includes('dificuldade')),
        marcacao: headerRow.findIndex(c => c && (c.toString().toLowerCase().includes('marcação respondente') || c.toString().toLowerCase().includes('marcaçao respondente') || c.toString().toLowerCase().includes('marcacao respondente'))),
        gabarito: headerRow.findIndex(c => c && c.toString().toLowerCase().trim() === 'gabarito'),
        curso: headerRow.findIndex(c => c && c.toString().toLowerCase().trim() === 'curso'),
        turma: headerRow.findIndex(c => c && c.toString().toLowerCase().trim() === 'turma'),
        avaliacao: headerRow.findIndex(c => c && (c.toString().toLowerCase().trim() === 'avaliação' || c.toString().toLowerCase().trim() === 'avaliacao')),
        escolaCol: headerRow.findIndex(c => c && c.toString().toLowerCase().includes('escola')),
        turnoCol: headerRow.findIndex(c => c && c.toString().toLowerCase().includes('turno')),
        periodoCol: headerRow.findIndex(c => c && (c.toString().toLowerCase().includes('período de agendamento') || c.toString().toLowerCase().includes('periodo de agendamento')))
      };

      if (colIdx.identificador >= 0 && colIdx.matricula >= 0) {
        possuiAbaRegistro = true;
        
        // Get metadata from the first data row
        const primeiraLinhaDados = registroRows[headerRowIndex + 1];
        if (primeiraLinhaDados) {
          if (colIdx.turma >= 0 && primeiraLinhaDados[colIdx.turma]) turmaName = primeiraLinhaDados[colIdx.turma].toString().trim();
          if (colIdx.curso >= 0 && primeiraLinhaDados[colIdx.curso]) curso = primeiraLinhaDados[colIdx.curso].toString().trim();
          if (colIdx.avaliacao >= 0 && primeiraLinhaDados[colIdx.avaliacao]) dataProva = primeiraLinhaDados[colIdx.avaliacao].toString().trim();
          if (colIdx.escolaCol >= 0 && primeiraLinhaDados[colIdx.escolaCol]) escola = primeiraLinhaDados[colIdx.escolaCol].toString().trim();
          if (colIdx.turnoCol >= 0 && primeiraLinhaDados[colIdx.turnoCol]) turno = primeiraLinhaDados[colIdx.turnoCol].toString().trim();
          if (colIdx.periodoCol >= 0 && primeiraLinhaDados[colIdx.periodoCol]) periodo = primeiraLinhaDados[colIdx.periodoCol].toString().trim();
        }

        // Loop all rows after the header row
        for (let i = headerRowIndex + 1; i < registroRows.length; i++) {
          const row = registroRows[i];
          if (!row || row.length <= Math.max(colIdx.matricula, colIdx.identificador)) continue;
          
          const matriculaAluno = row[colIdx.matricula];
          if (!matriculaAluno) continue;
          
          const chave = `${matriculaAluno.toString().trim()}`;
          if (!detalhesAlunos[chave]) {
            detalhesAlunos[chave] = [];
          }
          
          const id = row[colIdx.identificador].toString().trim();
          const pedInfo = questionPedagogicalMap[id] || {};

          const marcacao = colIdx.marcacao >= 0 ? row[colIdx.marcacao].toString().trim() : '';
          const gabarito = colIdx.gabarito >= 0 ? row[colIdx.gabarito].toString().trim() : '';
          const acertou = marcacao === gabarito;
          
          detalhesAlunos[chave].push({
            identificador: id,
            capacidade: formatCapacidade(pedInfo.capacidade || (colIdx.capacidade >= 0 ? row[colIdx.capacidade] : '')),
            subfuncao: pedInfo.subfuncao || (colIdx.subfuncao >= 0 ? row[colIdx.subfuncao] : ''),
            padraoDesempenho: pedInfo.padraoDesempenho || (colIdx.padraoDesempenho >= 0 ? row[colIdx.padraoDesempenho] : ''),
            conhecimento: pedInfo.conhecimento || (colIdx.conhecimento >= 0 ? row[colIdx.conhecimento] : ''),
            dificuldade: pedInfo.dificuldade || (colIdx.dificuldade >= 0 ? row[colIdx.dificuldade] : ''),
            marcacao: marcacao,
            gabarito: gabarito,
            acertou: acertou
          });
        }
      }
    }

    // Source B: Fallback parsing from 'Desempenho Individual' (block-by-block format)
    if (Object.keys(detalhesAlunos).length === 0 && individualRows) {
      let i = 0;
      while (i < individualRows.length) {
        const row = individualRows[i];
        if (row && row[0] && row[1] && 
            row[0].toString().toLowerCase().trim() === 'aluno' && 
            (row[1].toString().toLowerCase().trim() === 'matrícula' || row[1].toString().toLowerCase().trim() === 'matricula')) {
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
                  identificador: qRow[0].toString().trim(),
                  capacidade: formatCapacidade(qRow[1]),
                  subfuncao: qRow[2] ? qRow[2].toString().trim() : '',
                  padraoDesempenho: qRow[3] ? qRow[3].toString().trim() : '',
                  conhecimento: qRow[4] ? qRow[4].toString().trim() : '',
                  dificuldade: qRow[5] ? qRow[5].toString().trim() : '',
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

    // Attach details and recalculate individual student stats to ensure absolute consistency
    alunos.forEach((aluno) => {
      const chave = `${aluno.matricula}`;
      aluno.questoes = detalhesAlunos[chave] || [];
      
      if (aluno.questoes.length > 0) {
        aluno.acertos = aluno.questoes.filter(q => q.acertou).length;
        aluno.erros = aluno.questoes.filter(q => !q.acertou).length;
        aluno.desempenho = parseFloat(((aluno.acertos / aluno.questoes.length) * 100).toFixed(2));
      }
    });

    // Safeguard: Override totalItens if more questions are parsed
    let maxQuestoes = 0;
    alunos.forEach(a => {
      if (a.questoes && a.questoes.length > maxQuestoes) {
        maxQuestoes = a.questoes.length;
      }
    });
    const finalTotalItens = Math.max(totalItens, maxQuestoes);

    // Recalculate class averages dynamically if we parsed student questions
    let finalDesempenho = desempenho;
    let finalAcertosMedio = acertosMedio;
    let finalErrosMedio = errosMedio;

    if (alunos.length > 0 && alunos[0].questoes.length > 0) {
      const totalAlunos = alunos.length;
      const somaDesempenho = alunos.reduce((acc, a) => acc + a.desempenho, 0);
      const somaAcertos = alunos.reduce((acc, a) => acc + a.acertos, 0);
      const somaErros = alunos.reduce((acc, a) => acc + a.erros, 0);
      
      finalDesempenho = parseFloat((somaDesempenho / totalAlunos).toFixed(2));
      finalAcertosMedio = parseFloat((somaAcertos / totalAlunos).toFixed(2));
      finalErrosMedio = parseFloat((somaErros / totalAlunos).toFixed(2));
    }

    // Sort students by performance
    alunos.sort((a, b) => b.desempenho - a.desempenho);
    
    return {
      titulo,
      subtitulo,
      desempenho: finalDesempenho,
      acertosMedio: finalAcertosMedio,
      errosMedio: finalErrosMedio,
      totalItens: finalTotalItens,
      respondentes,
      labelsStats,
      headers,
      alunos,
      possuiAbaRegistro,
      curso: curso || 'Curso Técnico',
      turmaName: turmaName || 'Turma Geral',
      dataProva: dataProva || 'Avaliação SAEP',
      escola,
      turno,
      periodo,
      dataProcessamento: new Date().toLocaleString('pt-BR')
    };
  };

  const currentData = useMemo(() => {
    if (!selectedTurma || !turmas[selectedTurma]) return null;
    return turmas[selectedTurma];
  }, [selectedTurma, turmas]);

  // Statistics calculations
  const statsCalculated = useMemo(() => {
    if (!currentData) return null;

    const melhorAluno = currentData.alunos[0];
    const piorAluno = currentData.alunos[currentData.alunos.length - 1];

    // Compute times
    const temposSegundos = currentData.alunos
      .map((a) => {
        const partes = a.tempo.split(':');
        if (partes.length === 3) {
          return parseInt(partes[0]) * 3600 + parseInt(partes[1]) * 60 + parseInt(partes[2]);
        }
        return 0;
      })
      .filter((t) => t > 0);

    const tempoMedioSegundos = temposSegundos.length > 0 
      ? temposSegundos.reduce((sum, val) => sum + val, 0) / temposSegundos.length 
      : 0;

    const formatarTempo = (segundosTotal) => {
      if (segundosTotal === 0) return '00:00:00';
      const h = Math.floor(segundosTotal / 3600);
      const m = Math.floor((segundosTotal % 3600) / 60);
      const s = Math.floor(segundosTotal % 60);
      return `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}:${String(s).padStart(2, '0')}`;
    };

    const tempoMedioStr = formatarTempo(tempoMedioSegundos);

    // Students above avg
    const alunosAcimaMedia = currentData.alunos.filter(
      (a) => a.desempenho >= currentData.desempenho
    ).length;

    // Passing (>= 70%)
    const alunosAprovados = currentData.alunos.filter((a) => a.desempenho >= 70).length;

    // Metrics by Area of Knowledge
    const desempenhoPorAssunto = {};
    if (currentData.possuiAbaRegistro) {
      currentData.alunos.forEach((a) => {
        a.questoes.forEach((q) => {
          const assunto = q.conhecimento || 'Outros';
          if (!desempenhoPorAssunto[assunto]) {
            desempenhoPorAssunto[assunto] = { acertos: 0, total: 0 };
          }
          desempenhoPorAssunto[assunto].total += 1;
          if (q.acertou) {
            desempenhoPorAssunto[assunto].acertos += 1;
          }
        });
      });
    }

    const dadosAssuntoChart = Object.keys(desempenhoPorAssunto).map((assunto) => {
      const item = desempenhoPorAssunto[assunto];
      return {
        assunto: assunto.length > 20 ? assunto.slice(0, 18) + '...' : assunto,
        assuntoCompleto: assunto,
        desempenho: Math.round((item.acertos / item.total) * 100),
      };
    }).sort((a, b) => b.desempenho - a.desempenho);

    // Distribution categories
    let excelente = 0, bom = 0, regular = 0, insuficiente = 0;
    currentData.alunos.forEach((a) => {
      if (a.desempenho >= 90) excelente++;
      else if (a.desempenho >= 70) bom++;
      else if (a.desempenho >= 50) regular++;
      else insuficiente++;
    });

    const dadosNivelChart = [
      { name: 'Excelente (≥90%)', quantidade: excelente, fill: '#10b981' },
      { name: 'Bom (70%-89%)', quantidade: bom, fill: '#06b6d4' },
      { name: 'Regular (50%-69%)', quantidade: regular, fill: '#f59e0b' },
      { name: 'Insuficiente (<50%)', quantidade: insuficiente, fill: '#f43f5e' }
    ];

    // Scatter data: Time vs performance
    const dadosDispersao = currentData.alunos.map((a) => {
      const partes = a.tempo.split(':');
      let minutos = 0;
      if (partes.length === 3) {
        minutos = parseInt(partes[0]) * 60 + parseInt(partes[1]) + parseInt(partes[2]) / 60;
      }
      
      let n = 'Insuficiente';
      if (a.desempenho >= 90) n = 'Excelente';
      else if (a.desempenho >= 70) n = 'Bom';
      else if (a.desempenho >= 50) n = 'Regular';

      return {
        name: a.nome,
        tempoMinutos: Math.round(minutos * 10) / 10,
        tempoStr: a.tempo,
        desempenho: a.desempenho,
        nivel: n,
        acertos: a.acertos
      };
    });

    return {
      melhorAluno,
      piorAluno,
      tempoMedioSegundos,
      tempoMedioStr,
      alunosAcimaMedia,
      alunosAprovados,
      taxaAproveitamento: Math.round((alunosAprovados / currentData.respondentes) * 100),
      dadosNivelChart,
      dadosAssuntoChart,
      dadosDispersao
    };
  }, [currentData]);

  const currentAluno = useMemo(() => {
    if (!currentData) return null;
    const matriculaToFind = selectedAlunoMatricula || currentData.alunos[0]?.matricula;
    return currentData.alunos.find(a => a.matricula === matriculaToFind) || currentData.alunos[0];
  }, [selectedAlunoMatricula, currentData]);

  const matrizQuestoes = useMemo(() => {
    if (!currentData || !currentData.possuiAbaRegistro) return [];
    
    const questMap = {};
    currentData.alunos.forEach(aluno => {
      aluno.questoes.forEach(q => {
        const id = q.identificador;
        if (!id) return;
        if (!questMap[id]) {
          questMap[id] = {
            identificador: id,
            capacidade: q.capacidade,
            conhecimento: q.conhecimento,
            subfuncao: q.subfuncao,
            padraoDesempenho: q.padraoDesempenho,
            dificuldade: q.dificuldade,
            acertos: 0,
            total: 0
          };
        }
        questMap[id].total++;
        if (q.acertou) questMap[id].acertos++;
      });
    });

    return Object.values(questMap).map(item => {
      const taxa = item.total > 0 ? parseFloat(((item.acertos / item.total) * 100).toFixed(1)) : 0;
      let status = 'dominada'; // > 70%
      if (taxa < 40) status = 'critica';
      else if (taxa <= 70) status = 'atencao';
      
      return {
        ...item,
        taxa,
        status
      };
    }).sort((a, b) => a.taxa - b.taxa);
  }, [currentData]);

  const todosRegistros = useMemo(() => {
    if (!currentData || !currentData.possuiAbaRegistro) return [];
    const logs = [];
    currentData.alunos.forEach(a => {
      a.questoes.forEach(q => {
        logs.push({
          alunoNome: a.nome,
          alunoMatricula: a.matricula,
          identificador: q.identificador,
          capacidade: q.capacidade,
          subfuncao: q.subfuncao,
          padraoDesempenho: q.padraoDesempenho,
          conhecimento: q.conhecimento,
          marcacao: q.marcacao || 'N/R',
          gabarito: q.gabarito,
          acertou: q.acertou
        });
      });
    });
    return logs;
  }, [currentData]);

  const filteredRegistros = useMemo(() => {
    const searchLower = registroSearch.toLowerCase();
    return todosRegistros.filter(r => 
      (r.alunoNome && r.alunoNome.toLowerCase().includes(searchLower)) ||
      (r.alunoMatricula && r.alunoMatricula.toString().includes(searchLower)) ||
      (r.identificador && r.identificador.toString().toLowerCase().includes(searchLower)) ||
      (r.conhecimento && r.conhecimento.toLowerCase().includes(searchLower)) ||
      (r.capacidade && r.capacidade.toLowerCase().includes(searchLower)) ||
      (r.subfuncao && r.subfuncao.toLowerCase().includes(searchLower)) ||
      (r.padraoDesempenho && r.padraoDesempenho.toLowerCase().includes(searchLower))
    );
  }, [todosRegistros, registroSearch]);

  const paginatedRegistros = useMemo(() => {
    const start = (registroPage - 1) * 25;
    return filteredRegistros.slice(start, start + 25);
  }, [filteredRegistros, registroPage]);

  const maxRegistroPages = useMemo(() => {
    return Math.ceil(filteredRegistros.length / 25);
  }, [filteredRegistros]);

  // Filtering
  const filteredAlunos = useMemo(() => {
    if (!currentData) return [];
    return currentData.alunos.filter((aluno) => {
      const matchesSearch = aluno.nome.toLowerCase().includes(searchQuery.toLowerCase()) ||
                            aluno.matricula.toString().includes(searchQuery);
      
      let matchesFilter = true;
      if (filterLevel === 'excelente') matchesFilter = aluno.desempenho >= 90;
      else if (filterLevel === 'bom') matchesFilter = aluno.desempenho >= 70 && aluno.desempenho < 90;
      else if (filterLevel === 'regular') matchesFilter = aluno.desempenho >= 50 && aluno.desempenho < 70;
      else if (filterLevel === 'insuficiente') matchesFilter = aluno.desempenho < 50;

      return matchesSearch && matchesFilter;
    });
  }, [currentData, searchQuery, filterLevel]);

  const filteredQuestoesAluno = useMemo(() => {
    if (!currentAluno || !currentAluno.questoes) return [];
    return currentAluno.questoes.filter(q => {
      if (alunoQuestaoFilter === 'acertos') return q.acertou;
      if (alunoQuestaoFilter === 'erros') return !q.acertou;
      return true;
    });
  }, [currentAluno, alunoQuestaoFilter]);

  const assuntoStats = useMemo(() => {
    if (!currentAluno || !currentAluno.questoes) return [];
    const map = {};
    currentAluno.questoes.forEach(q => {
      const assunto = q.conhecimento || 'Geral';
      if (!map[assunto]) map[assunto] = { acertos: 0, total: 0 };
      map[assunto].total++;
      if (q.acertou) map[assunto].acertos++;
    });
    return Object.entries(map).map(([assunto, val]) => ({
      assunto,
      acertos: val.acertos,
      total: val.total,
      percent: Math.round((val.acertos / val.total) * 100)
    }));
  }, [currentAluno]);

  const criticidadeStats = useMemo(() => {
    let criticas = 0, atencao = 0, dominadas = 0;
    matrizQuestoes.forEach(q => {
      if (q.status === 'critica') criticas++;
      else if (q.status === 'atencao') atencao++;
      else dominadas++;
    });
    return { criticas, atencao, dominadas };
  }, [matrizQuestoes]);

  const iaPrompt = useMemo(() => {
    if (!currentData) return '';
    
    const curso = currentData.curso || 'Curso Técnico';
    const turma = currentData.turmaName || 'Turma Geral';
    const escola = currentData.escola || 'SENAI';
    const mediaTurma = currentData.desempenho;
    const totalItens = currentData.totalItens;
    const totalAlunos = currentData.respondentes;
    
    // Critical items (< 50% success rate)
    const itensCriticos = matrizQuestoes.filter(q => q.taxa < 50);
    const itensCriticosText = itensCriticos.length > 0 
      ? itensCriticos.map(q => 
          `- Item #${q.identificador} (${q.taxa.toFixed(1)}% de acertos):
   • Conhecimento: ${q.conhecimento || 'Geral'}
   • Capacidade: ${q.capacidade || 'Não especificada'}
   • Subfunção: ${q.subfuncao || 'Não especificada'}
   • Dificuldade: ${q.dificuldade || 'Média'}`
        ).join('\n\n')
      : 'Nenhum item crítico identificado (todas as questões tiveram taxa de acertos superior a 50%).';

    // Mastered items (>= 70% success rate)
    const itensDominados = matrizQuestoes.filter(q => q.taxa >= 70);
    const itensDominadosText = itensDominados.length > 0
      ? itensDominados.map(q => 
          `- Item #${q.identificador} (${q.taxa.toFixed(1)}% de acertos): ${q.conhecimento || 'Geral'}`
        ).join('\n')
      : 'Nenhum item com taxa de acertos superior a 70%.';

    // Students needing support (< 50% score)
    const alunosCriticos = currentData.alunos.filter(a => a.desempenho < 50);
    const alunosCriticosText = alunosCriticos.length > 0
      ? alunosCriticos.map(a => `- ${a.nome} (${a.desempenho.toFixed(1)}% de aproveitamento - ${a.acertos} acertos de ${totalItens})`).join('\n')
      : 'Nenhum estudante com aproveitamento inferior a 50%.';

    return `Aja como um especialista em coordenação pedagógica, design educacional e tutor do SENAI.
Fui encarregado de realizar uma intervenção pedagógica de revisão com base nos resultados obtidos por uma turma na avaliação oficial do SAEP (Sistema de Avaliação da Educação Profissional).

Aqui estão os dados estruturados da avaliação para análise:

### DADOS GERAIS DA TURMA
- Escola: ${escola}
- Curso/Área: ${curso}
- Turma: ${turma}
- Média de Aproveitamento Coletivo: ${mediaTurma}%
- Total de Itens (Questões) Avaliados: ${totalItens}
- Número de Estudantes Participantes: ${totalAlunos}

### ITENS CRÍTICOS (Taxa de acerto inferior a 50%)
Estes são os tópicos em que a turma apresentou maior dificuldade e necessita de revisão prioritária:
${itensCriticosText}

### ITENS DOMINADOS (Taxa de acerto superior a 70%)
A turma demonstrou bom domínio nestes assuntos, exigindo menor tempo de intervenção:
${itensDominadosText}

### ESTUDANTES QUE NECESSITAM DE ATENÇÃO INDIVIDUAL (Aproveitamento inferior a 50%)
Estes estudantes estão abaixo da média esperada e precisam de suporte ou reforço adaptado:
${alunosCriticosText}

---

### SUA DIRETRIZ:
Com base no diagnóstico acima, elabore um plano de ação e revisão pedagógica extremamente prático e estruturado, contendo:

1. **Priorização Pedagógica**: Identifique as principais lacunas conceituais comuns (cruzando os itens críticos com suas respectivas capacidades e conhecimentos) e explique por que a turma pode ter errado esses pontos.
2. **Roteiro de Aula Prático (Recuperação)**: Crie uma sugestão de aula de revisão de 4 horas focada na superação dos itens críticos, utilizando metodologias ativas (como rotação por estações, aprendizagem baseada em problemas ou instrução pelos pares) alinhadas ao perfil do SENAI (foco prático e técnico).
3. **Plano de Suporte Individualizado**: Sugira estratégias de recuperação contínua para os estudantes listados na seção de atenção individual, sem desacelerar o restante do cronograma da turma.
4. **Questões de Fixação**: Crie 2 exemplos de questões conceituais inéditas no mesmo nível de complexidade das questões críticas para que eu possa aplicar em sala de aula como verificação de aprendizagem.`;
  }, [currentData, matrizQuestoes]);

  const getClassificacaoTempo = (tempoStr, tempoMedioSeg) => {
    if (tempoMedioSeg === 0) return { label: 'Normal', colorClass: 'normal', isSlow: false, isFast: false };
    const partes = tempoStr.split(':');
    if (partes.length !== 3) return { label: 'Normal', colorClass: 'normal', isSlow: false, isFast: false };
    const seg = parseInt(partes[0]) * 3600 + parseInt(partes[1]) * 60 + parseInt(partes[2]);
    
    if (seg < tempoMedioSeg * 0.8) {
      return { label: 'Rápido', colorClass: 'fast', isSlow: false, isFast: true };
    } else if (seg > tempoMedioSeg * 1.2) {
      return { label: 'Lento', colorClass: 'slow', isSlow: true, isFast: false };
    } else {
      return { label: 'Normal', colorClass: 'normal', isSlow: false, isFast: false };
    }
  };

  const getScoreClass = (score) => {
    if (score >= 90) return 'excelente';
    if (score >= 70) return 'bom';
    if (score >= 50) return 'regular';
    return 'insuficiente';
  };

  const removerTurma = async (nomeKey, e) => {
    e.stopPropagation();
    await deleteTurmaFromSupabase(nomeKey);
    const novasTurmas = { ...turmas };
    delete novasTurmas[nomeKey];
    setTurmas(novasTurmas);

    const restanteKeys = Object.keys(novasTurmas);
    if (restanteKeys.length > 0) {
      const proximaTurma = restanteKeys[0];
      setSelectedTurma(proximaTurma);
      setSelectedAlunoMatricula(novasTurmas[proximaTurma]?.alunos[0]?.matricula || null);
    } else {
      setSelectedTurma(null);
      setSelectedAlunoMatricula(null);
    }
  };

  const handleExportCSV = () => {
    if (!filteredRegistros || filteredRegistros.length === 0) return;
    
    // CSV Headers
    const csvHeaders = [
      'Aluno', 
      'Matrícula', 
      'Questão ID', 
      'Capacidade', 
      'Subfunção', 
      'Padrão de desempenho', 
      'Conhecimento', 
      'Marcado', 
      'Gabarito', 
      'Situação'
    ];
    
    // CSV Rows
    const csvRows = filteredRegistros.map(r => [
      `"${r.alunoNome.replace(/"/g, '""')}"`,
      `"${r.alunoMatricula}"`,
      `"${r.identificador}"`,
      `"${(r.capacidade || '').replace(/"/g, '""')}"`,
      `"${(r.subfuncao || '').replace(/"/g, '""')}"`,
      `"${(r.padraoDesempenho || '').replace(/"/g, '""')}"`,
      `"${(r.conhecimento || '').replace(/"/g, '""')}"`,
      `"${r.marcacao}"`,
      `"${r.gabarito}"`,
      `"${r.acertou ? 'Acerto' : 'Erro'}"`
    ]);
    
    const csvContent = [csvHeaders.join(','), ...csvRows.map(row => row.join(','))].join('\n');
    
    // Use BOM for Excel compatibility in UTF-8
    const blob = new Blob(['\ufeff' + csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.setAttribute('href', url);
    link.setAttribute('download', `registros_saep_${currentData?.turmaName || 'turma'}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  if (isSupabaseConfigured && authLoading) {
    return (
      <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', height: '100vh', background: 'var(--bg-primary)', color: 'white' }}>
        <RefreshCw className="animate-spin" size={40} style={{ color: 'var(--accent-orange)' }} />
        <p style={{ marginTop: '1rem', fontSize: '0.85rem', color: 'var(--text-secondary)' }}>Carregando sessão...</p>
      </div>
    );
  }

  if (isSupabaseConfigured && !session) {
    return (
      <div className="app-container" style={{ justifyContent: 'center', alignItems: 'center', height: '100vh', background: 'radial-gradient(circle at top right, rgba(235, 87, 36, 0.08), transparent 40%), radial-gradient(circle at bottom left, rgba(6, 182, 212, 0.08), transparent 40%), var(--bg-primary)' }}>
        <div className="glass-panel" style={{ width: '100%', maxWidth: '420px', padding: '2.5rem', borderRadius: '16px', boxShadow: '0 20px 40px rgba(0,0,0,0.4)', border: '1px solid rgba(255,255,255,0.06)' }}>
          
          <div style={{ textAlign: 'center', marginBottom: '2rem' }}>
            <div className="senai-logo-container" style={{ margin: '0 auto 1rem', display: 'flex', justifyContent: 'center' }}>
              <span>SENA</span>
              <span className="logo-letter-i">
                <span className="i-top"></span>
                <span className="i-bottom"></span>
              </span>
            </div>
            <h2 style={{ fontSize: '1.4rem', color: 'white', fontWeight: '800', fontFamily: 'Outfit' }}>Portal de Acompanhamento SAEP</h2>
            <p style={{ fontSize: '0.8rem', color: 'var(--text-secondary)', marginTop: '0.35rem' }}>
              {isSignUpMode ? 'Crie sua conta para salvar suas turmas na nuvem' : 'Entre para acessar suas planilhas de qualquer lugar'}
            </p>
          </div>

          {authError && (
            <div style={{ padding: '0.75rem', background: 'rgba(244, 63, 94, 0.08)', borderRadius: '8px', border: '1px solid rgba(244, 63, 94, 0.2)', fontSize: '0.75rem', color: 'var(--accent-rose)', marginBottom: '1.25rem' }}>
              {authError}
            </div>
          )}

          <form onSubmit={async (e) => {
            e.preventDefault();
            setAuthError(null);
            const email = e.target.email.value;
            const password = e.target.password.value;
            
            if (isSignUpMode) {
              const { error } = await supabase.auth.signUp({ email, password });
              if (error) {
                setAuthError(error.message);
              } else {
                setAuthError("Cadastro realizado! Verifique seu e-mail para confirmar a conta ou tente fazer o login.");
                setIsSignUpMode(false);
              }
            } else {
              const { error } = await supabase.auth.signInWithPassword({ email, password });
              if (error) {
                setAuthError(error.message);
              }
            }
          }} style={{ display: 'flex', flexDirection: 'column', gap: '1.25rem' }}>
            <div>
              <label style={{ display: 'block', fontSize: '0.65rem', textTransform: 'uppercase', color: 'var(--text-secondary)', fontWeight: 'bold', marginBottom: '0.35rem' }}>
                E-mail
              </label>
              <input
                name="email"
                type="email"
                required
                placeholder="professor@senai.com.br"
                style={{ width: '100%', padding: '0.7rem', background: 'var(--bg-input)', border: '1px solid var(--border-color)', borderRadius: 'var(--radius-sm)', color: 'white', fontSize: '0.8rem', outline: 'none' }}
              />
            </div>

            <div>
              <label style={{ display: 'block', fontSize: '0.65rem', textTransform: 'uppercase', color: 'var(--text-secondary)', fontWeight: 'bold', marginBottom: '0.35rem' }}>
                Senha
              </label>
              <input
                name="password"
                type="password"
                required
                placeholder="******"
                style={{ width: '100%', padding: '0.7rem', background: 'var(--bg-input)', border: '1px solid var(--border-color)', borderRadius: 'var(--radius-sm)', color: 'white', fontSize: '0.8rem', outline: 'none' }}
              />
            </div>

            <button
              type="submit"
              className="btn-select"
              style={{ width: '100%', justifyContent: 'center', padding: '0.8rem', fontSize: '0.85rem', fontWeight: 'bold', background: 'linear-gradient(135deg, var(--accent-orange), #ff8533)', boxShadow: 'var(--shadow-glow-orange)', marginTop: '0.5rem' }}
            >
              {isSignUpMode ? 'Criar Conta' : 'Entrar no Portal'}
            </button>
          </form>

          <div style={{ display: 'flex', flexDirection: 'column', gap: '0.75rem', marginTop: '1.5rem' }}>
            <div style={{ textAlign: 'center', color: 'var(--text-muted)', fontSize: '0.65rem', fontWeight: 'bold', textTransform: 'uppercase', letterSpacing: '0.05em' }}>
              — OU —
            </div>
            
            <button
              type="button"
              onClick={async () => {
                setAuthError(null);
                const { error } = await supabase.auth.signInAnonymously();
                if (error) {
                  setAuthError(error.message);
                }
              }}
              className="btn-help"
              style={{ width: '100%', justifyContent: 'center', gap: '0.5rem', background: 'rgba(255, 255, 255, 0.03)', borderColor: 'var(--border-color)', padding: '0.8rem', fontSize: '0.8rem', fontWeight: 'bold', color: 'white', cursor: 'pointer' }}
            >
              <Sparkles size={16} style={{ color: 'var(--accent-orange)' }} />
              Entrar como Convidado (Rápido)
            </button>
          </div>

          <div style={{ textAlign: 'center', marginTop: '1.5rem', fontSize: '0.75rem' }}>
            <button
              onClick={() => {
                setAuthError(null);
                setIsSignUpMode(!isSignUpMode);
              }}
              style={{ background: 'transparent', border: 'none', color: 'var(--accent-cyan)', cursor: 'pointer', textDecoration: 'underline' }}
            >
              {isSignUpMode ? 'Já tenho uma conta. Fazer Login.' : 'Não tem conta? Crie uma aqui.'}
            </button>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div className="app-container">
      {/* Sidebar Navigation */}
      {Object.keys(turmas).length > 0 && (
        <aside className="sidebar">
          <div className="sidebar-header-wrapper">
            <div className="sidebar-brand">
              <div className="senai-logo-container-small">
                <span className="logo-s">S</span>
              </div>
              <div className="sidebar-brand-text">
                <span className="saep-portal-prefix-small">Portal</span>
                <span className="saep-title-small">SAEP</span>
              </div>
            </div>
          </div>

          <div className="sidebar-menu">
            <button
              onClick={() => setActiveSubTab('resumo')}
              className={`sidebar-item ${activeSubTab === 'resumo' ? 'active' : ''}`}
              title="Desempenho Individual - Resumo"
            >
              <div className="sidebar-icon">
                <BarChart3 size={18} />
              </div>
              <span className="sidebar-label">Resumo da Turma</span>
            </button>
            
            <button
              onClick={() => {
                setActiveSubTab('individual');
                if (currentData && !selectedAlunoMatricula) {
                  setSelectedAlunoMatricula(currentData.alunos[0]?.matricula || null);
                }
              }}
              className={`sidebar-item ${activeSubTab === 'individual' ? 'active' : ''}`}
              title="Desempenho Individual"
            >
              <div className="sidebar-icon">
                <Users2 size={18} />
              </div>
              <span className="sidebar-label">Análise por Aluno</span>
            </button>
            
            <button
              onClick={() => setActiveSubTab('registro')}
              className={`sidebar-item ${activeSubTab === 'registro' ? 'active' : ''}`}
              title="Desempenho Ind. por Registro"
            >
              <div className="sidebar-icon">
                <FileSpreadsheet size={18} />
              </div>
              <span className="sidebar-label">Registros e Itens</span>
            </button>
            
            <button
              onClick={() => setActiveSubTab('relatorio')}
              className={`sidebar-item ${activeSubTab === 'relatorio' ? 'active' : ''}`}
              title="Gerar Relatório para IA"
            >
              <div className="sidebar-icon">
                <Sparkles size={18} style={{ color: 'var(--accent-orange)' }} />
              </div>
              <span className="sidebar-label" style={{ color: activeSubTab === 'relatorio' ? 'var(--accent-orange)' : 'inherit' }}>Relatório IA</span>
            </button>
          </div>

          <div className="sidebar-footer">
            <button
              onClick={() => setShowHelpModal(true)}
              className="sidebar-item help-item"
              title="Como importar planilhas?"
            >
              <div className="sidebar-icon">
                <HelpCircle size={18} />
              </div>
              <span className="sidebar-label">Ajuda</span>
            </button>
          </div>
        </aside>
      )}

      <div className={`main-layout ${Object.keys(turmas).length > 0 ? 'with-sidebar' : ''}`}>
        {/* Sticky Header */}
        <header className="header">
        <div className="container header-content">
          <div className="header-brand">
            <div className="senai-logo-container">
              <span>SENA</span>
              <span className="logo-letter-i">
                <span className="i-top"></span>
                <span className="i-bottom"></span>
              </span>
            </div>
            
            <div className="header-divider"></div>
            
            <div className="saep-brand-text">
              <span className="saep-portal-prefix">Portal</span>
              <h1 className="saep-title-new">SAEP</h1>
            </div>
          </div>
          
          <div className="header-actions" style={{ display: 'flex', alignItems: 'center', gap: '1rem' }}>
            {isSupabaseConfigured && session && (
              <div style={{ display: 'flex', alignItems: 'center', gap: '0.75rem', paddingRight: '0.75rem', borderRight: '1px solid rgba(255,255,255,0.08)' }}>
                <span style={{ fontSize: '0.75rem', color: 'var(--text-secondary)', display: 'flex', alignItems: 'center', gap: '0.35rem' }}>
                  <span style={{ width: '6px', height: '6px', borderRadius: '50%', backgroundColor: 'var(--accent-emerald)', display: 'inline-block' }}></span>
                  {session.user.is_anonymous ? 'Modo Convidado' : session.user.email}
                </span>
                
                <button
                  onClick={async () => {
                    await supabase.auth.signOut();
                  }}
                  className="btn-clear"
                  style={{ padding: '0.35rem 0.6rem', fontSize: '0.7rem', color: 'var(--accent-rose)', border: 'none', background: 'transparent', cursor: 'pointer', fontWeight: 'bold' }}
                  title="Sair da Conta"
                >
                  Sair
                </button>
              </div>
            )}

            {!isSupabaseConfigured && (
              <span style={{ fontSize: '0.65rem', color: 'var(--text-muted)', background: 'rgba(255, 255, 255, 0.02)', padding: '0.25rem 0.5rem', borderRadius: '4px', border: '1px solid var(--border-color)' }}>
                Modo Local (Sem Nuvem)
              </span>
            )}

            <button 
              onClick={() => setShowHelpModal(true)}
              className="btn-help"
            >
              <HelpCircle size={14} style={{ color: 'var(--accent-cyan)' }} />
              Como importar planilhas?
            </button>
            
            {currentData && (
              <span className="process-time">
                Atualizado: {currentData.dataProcessamento}
              </span>
            )}
          </div>
        </div>
      </header>

      {/* Main Workspace */}
      <main className="container" style={{ flexGrow: 1, padding: '2rem 0' }}>
        <input
          type="file"
          ref={fileInputRef}
          onChange={handleFileChange}
          accept=".xls,.xlsx"
          multiple
          className="hidden"
          style={{ display: 'none' }}
        />
        
        {/* Upload State */}
        {Object.keys(turmas).length === 0 ? (
          <div className="upload-wrapper animate-fade-in">
            <div className="upload-header">
              <h2>
                Análise Inteligente SAEP <Sparkles size={24} style={{ color: 'var(--accent-orange)' }} />
              </h2>
              <p>
                Importe planilhas de diagnóstico e resultados do portal SAEP para visualizar indicadores e relatórios avançados de acompanhamento pedagógico.
              </p>
            </div>

            <div
              onDragOver={handleDragOver}
              onDragLeave={handleDragLeave}
              onDrop={handleDrop}
              onClick={() => fileInputRef.current.click()}
              className={`upload-box ${isDragging ? 'dragging' : ''}`}
            >
              <div className="upload-icon-box">
                <UploadCloud size={40} />
              </div>
              <h3 className="upload-title">Arraste seus relatórios do SAEP</h3>
              <p className="upload-desc">
                Suporta múltiplos arquivos no formato <strong>.xls</strong> ou <strong>.xlsx</strong>.
              </p>
              <button className="btn-select">
                Selecionar Arquivos <ArrowRight size={14} />
              </button>
            </div>

            {isSupabaseConfigured && session && (
              <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '0.35rem', marginTop: '1.5rem', fontSize: '0.75rem', color: 'var(--text-muted)' }}>
                <CheckCircle2 size={12} style={{ color: 'var(--accent-emerald)' }} />
                Sincronização na nuvem ativa. Suas planilhas serão salvas automaticamente.
              </div>
            )}
            
            <p className="upload-footer-text">
              Processamento 100% em sand-box local. Seus dados permanecem seguros em seu navegador.
            </p>
          </div>
        ) : (
          <div className="dashboard-workspace animate-fade-in">
            {/* Top Workspace Tab list */}
            <div className="tabs-bar-container">
              <div className="file-tabs">
                {Object.keys(turmas).map((nomeKey) => (
                  <button
                    key={nomeKey}
                    onClick={() => {
                      setSelectedTurma(nomeKey);
                      setExpandedAluno(null);
                    }}
                    className={`tab-button ${selectedTurma === nomeKey ? 'active' : ''}`}
                  >
                    <FileSpreadsheet size={14} />
                    <span>{nomeKey}</span>
                    <span 
                      onClick={(e) => removerTurma(nomeKey, e)}
                      className="tab-remove"
                    >
                      <X size={12} />
                    </span>
                  </button>
                ))}
                
                <button
                  onClick={() => fileInputRef.current.click()}
                  className="tab-add"
                  title="Importar mais planilhas"
                >
                  <UploadCloud size={14} />
                </button>
              </div>

            </div>

            {/* Current loaded sheet data view */}
            {currentData && (
              <div>
                
                {/* Main Information Panel */}
                <div className="glass-panel info-panel">
                  <div className="info-content">
                    <div className="info-badges">
                      <span className="info-badge primary">SAEP Nacional</span>
                      {currentData.curso && (
                        <span className="info-badge secondary">{currentData.curso}</span>
                      )}
                    </div>
                    <h2 className="info-title">{currentData.titulo}</h2>
                    <p className="info-sub">
                      {currentData.escola && <span>Escola: <strong>{currentData.escola}</strong></span>}
                      {currentData.turmaName && <span>Turma: <strong>{currentData.turmaName}</strong></span>}
                      {currentData.turno && <span>Turno: <strong>{currentData.turno}</strong></span>}
                      {currentData.periodo && <span>Período: <strong>{currentData.periodo}</strong></span>}
                      {currentData.subtitulo && <span>Avaliação: <strong>{currentData.subtitulo}</strong></span>}
                      {currentData.dataProva && <span>Matriz: <strong>{currentData.dataProva}</strong></span>}
                    </p>
                  </div>
                  
                  {!currentData.possuiAbaRegistro && (
                    <div className="warn-panel">
                      <HelpCircle size={20} className="warn-icon" />
                      <div>
                        <h4 className="warn-title">Análise de Tópicos Desabilitada</h4>
                        <p className="warn-desc">
                          Esta planilha foi importada sem a aba detalhada <em>'por Registro'</em>. Estatísticas por assunto e detalhamento individual não estão disponíveis.
                        </p>
                      </div>
                    </div>
                  )}
                </div>

                {/* Main Views Container */}
                {activeSubTab === 'resumo' && statsCalculated && (
                  <div style={{ display: 'flex', flexDirection: 'column', gap: '1.5rem' }}>
                    
                    {/* Key Indicators Grid */}
                    <div className="kpi-grid">
                      {/* KPI 1 */}
                      <div className="glass-panel kpi-card cyan">
                        <div>
                          <p className="kpi-label">Desempenho Médio</p>
                          <h3 className="kpi-value">{currentData.desempenho.toFixed(2)}%</h3>
                          <p className="kpi-sub">Média geral da turma</p>
                        </div>
                        <div className="kpi-icon-box">
                          <Percent size={18} />
                        </div>
                      </div>

                      {/* KPI 2 */}
                      <div className="glass-panel kpi-card emerald">
                        <div>
                          <p className="kpi-label">Acertos Médios</p>
                          <h3 className="kpi-value">{currentData.acertosMedio.toFixed(2)}</h3>
                          <p className="kpi-sub">de {currentData.totalItens} itens avaliados</p>
                        </div>
                        <div className="kpi-icon-box">
                          <CheckCircle2 size={18} />
                        </div>
                      </div>

                      {/* KPI 3 */}
                      <div className="glass-panel kpi-card rose">
                        <div>
                          <p className="kpi-label">Erros Médios</p>
                          <h3 className="kpi-value">{currentData.errosMedio.toFixed(2)}</h3>
                          <p className="kpi-sub">de {currentData.totalItens} itens avaliados</p>
                        </div>
                        <div className="kpi-icon-box">
                          <XCircle size={18} />
                        </div>
                      </div>

                      {/* KPI 4 */}
                      <div className="glass-panel kpi-card violet">
                        <div>
                          <p className="kpi-label">Avaliados</p>
                          <h3 className="kpi-value">{currentData.respondentes}</h3>
                          <p className="kpi-sub">Alunos respondentes</p>
                        </div>
                        <div className="kpi-icon-box">
                          <Users size={18} />
                        </div>
                      </div>

                      {/* KPI 5 */}
                      <div className="glass-panel kpi-card slate">
                        <div>
                          <p className="kpi-label">Tempo Médio</p>
                          <h3 className="kpi-value" style={{ color: '#e2e8f0' }}>{statsCalculated.tempoMedioStr}</h3>
                          <p className="kpi-sub">Tempo médio de realização</p>
                        </div>
                        <div className="kpi-icon-box">
                          <Clock size={18} />
                        </div>
                      </div>

                      {/* KPI 6 */}
                      <div className="glass-panel kpi-card amber">
                        <div>
                          <p className="kpi-label">Melhor Nota</p>
                          <h3 className="kpi-value" style={{ fontSize: '1.25rem' }} title={statsCalculated.melhorAluno?.nome}>
                            {statsCalculated.melhorAluno?.desempenho.toFixed(1)}%
                          </h3>
                          <p className="kpi-sub" style={{ textOverflow: 'ellipsis', whiteSpace: 'nowrap', overflow: 'hidden', maxWidth: '140px' }}>
                            {statsCalculated.melhorAluno?.nome}
                          </p>
                        </div>
                        <div className="kpi-icon-box">
                          <Award size={18} />
                        </div>
                      </div>

                      {/* KPI 7 */}
                      <div className="glass-panel kpi-card cyan">
                        <div>
                          <p className="kpi-label">Acima da Média</p>
                          <h3 className="kpi-value" style={{ color: '#a5f3fc' }}>{statsCalculated.alunosAcimaMedia}</h3>
                          <p className="kpi-sub">{Math.round((statsCalculated.alunosAcimaMedia / currentData.respondentes) * 100)}% da turma</p>
                        </div>
                        <div className="kpi-icon-box">
                          <TrendingUp size={18} />
                        </div>
                      </div>

                      {/* KPI 8 */}
                      <div className="glass-panel kpi-card emerald">
                        <div>
                          <p className="kpi-label">Aproveitamento (≥70%)</p>
                          <h3 className="kpi-value" style={{ color: '#a7f3d0' }}>{statsCalculated.taxaAproveitamento}%</h3>
                          <p className="kpi-sub">{statsCalculated.alunosAprovados} de {currentData.respondentes} alunos</p>
                        </div>
                        <div className="kpi-icon-box">
                          <Target size={18} />
                        </div>
                      </div>
                    </div>

                    {/* Interactive Recharts Grid */}
                    <div className="charts-grid">
                      
                      {/* Chart 1: Distribuição de Niveis */}
                      <div className="glass-panel chart-card">
                        <div className="chart-header">
                          <h3 className="chart-title">
                            <span style={{ backgroundColor: 'var(--accent-cyan)' }}></span>
                            Distribuição dos Níveis de Desempenho
                          </h3>
                          <p className="chart-sub">Frequência absoluta de alunos por faixa de pontuação</p>
                        </div>
                        <div style={{ flexGrow: 1, minHeight: '280px' }}>
                          <ResponsiveContainer width="100%" height="100%">
                            <BarChart data={statsCalculated.dadosNivelChart} margin={{ top: 20, right: 10, left: -25, bottom: 5 }}>
                              <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.05)" />
                              <XAxis dataKey="name" stroke="#94a3b8" fontSize={10} tickLine={false} />
                              <YAxis stroke="#94a3b8" fontSize={11} tickLine={false} allowDecimals={false} />
                              <Tooltip 
                                contentStyle={{ 
                                  backgroundColor: '#0f1424', 
                                  borderColor: 'rgba(255,255,255,0.1)', 
                                  borderRadius: '10px',
                                  fontSize: '11px',
                                  color: '#fff'
                                }} 
                                cursor={{ fill: 'rgba(255,255,255,0.02)' }}
                              />
                              <Bar dataKey="quantidade" radius={[6, 6, 0, 0]}>
                                {statsCalculated.dadosNivelChart.map((entry, index) => (
                                  <Cell key={`cell-${index}`} fill={entry.fill} />
                                ))}
                              </Bar>
                            </BarChart>
                          </ResponsiveContainer>
                        </div>
                      </div>

                      {/* Chart 2: Competencias/Assuntos */}
                      <div className="glass-panel chart-card">
                        <div className="chart-header">
                          <h3 className="chart-title">
                            <span style={{ backgroundColor: 'var(--accent-violet)' }}></span>
                            Desempenho por Descritor / Assunto
                          </h3>
                          <p className="chart-sub">Percentual médio de acertos nas principais áreas mapeadas</p>
                        </div>
                        {currentData.possuiAbaRegistro ? (
                          <div style={{ flexGrow: 1, minHeight: '280px' }}>
                            <ResponsiveContainer width="100%" height="100%">
                              <BarChart 
                                data={statsCalculated.dadosAssuntoChart} 
                                layout="vertical"
                                margin={{ top: 5, right: 20, left: 30, bottom: 5 }}
                              >
                                <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.05)" />
                                <XAxis type="number" stroke="#94a3b8" fontSize={10} domain={[0, 100]} unit="%" tickLine={false} />
                                <YAxis type="category" dataKey="assunto" stroke="#94a3b8" fontSize={9} tickLine={false} width={120} />
                                <Tooltip
                                  formatter={(value) => [`${value}%`, 'Acertos Médio']}
                                  labelFormatter={(label, items) => items[0]?.payload?.assuntoCompleto || label}
                                  contentStyle={{ 
                                    backgroundColor: '#0f1424', 
                                    borderColor: 'rgba(255,255,255,0.1)', 
                                    borderRadius: '10px',
                                    fontSize: '11px',
                                    color: '#fff'
                                  }}
                                  cursor={{ fill: 'rgba(255,255,255,0.02)' }}
                                />
                                <Bar dataKey="desempenho" radius={[0, 4, 4, 0]}>
                                  {statsCalculated.dadosAssuntoChart.map((entry, index) => {
                                    let color = 'var(--accent-cyan)';
                                    if (entry.desempenho >= 75) color = 'var(--accent-emerald)';
                                    else if (entry.desempenho < 50) color = 'var(--accent-rose)';
                                    return <Cell key={`cell-${index}`} fill={color} />;
                                  })}
                                </Bar>
                              </BarChart>
                            </ResponsiveContainer>
                          </div>
                        ) : (
                          <div className="chart-placeholder">
                            <FileSpreadsheet size={32} style={{ color: 'var(--text-muted)' }} />
                            <p>
                              Aba detalhada <em>'por Registro'</em> ausente. Carregue uma planilha completa para processar a estatística por conteúdo de descritor.
                            </p>
                          </div>
                        )}
                      </div>

                      {/* Chart 3: Dispersão Tempo vs Acertos */}
                      <div className="glass-panel chart-card col-span-2">
                        <div className="chart-header">
                          <h3 className="chart-title">
                            <span style={{ backgroundColor: 'var(--accent-amber)' }}></span>
                            Dispersão Didática: Tempo de Realização vs. Desempenho
                          </h3>
                          <p className="chart-sub">Mapeia a duração de execução em minutos contra o percentual de acerto individual (ótimo para sinalizar chutes ou gargalos de tempo)</p>
                        </div>
                        <div style={{ flexGrow: 1, minHeight: '280px' }}>
                          <ResponsiveContainer width="100%" height="100%">
                            <ScatterChart margin={{ top: 20, right: 20, bottom: 10, left: -20 }}>
                              <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.05)" />
                              <XAxis 
                                type="number" 
                                dataKey="tempoMinutos" 
                                name="Tempo" 
                                unit=" min" 
                                stroke="#94a3b8" 
                                fontSize={10} 
                                tickLine={false}
                              />
                              <YAxis 
                                type="number" 
                                dataKey="desempenho" 
                                name="Nota" 
                                unit="%" 
                                stroke="#94a3b8" 
                                fontSize={10} 
                                tickLine={false}
                                domain={[0, 100]}
                              />
                              <ZAxis type="number" dataKey="acertos" range={[60, 220]} />
                              <Tooltip 
                                cursor={{ strokeDasharray: '3 3', stroke: '#475569' }} 
                                content={({ active, payload }) => {
                                  if (active && payload && payload.length) {
                                    const data = payload[0].payload;
                                    return (
                                      <div style={{ backgroundColor: '#0f1424', border: '1px solid rgba(255,255,255,0.1)', padding: '0.75rem', borderRadius: '8px', fontSize: '11px', color: '#fff' }}>
                                        <p style={{ fontWeight: 'bold', marginBottom: '0.35rem' }}>{data.name}</p>
                                        <p>Desempenho: <strong style={{ color: 'var(--accent-cyan)' }}>{data.desempenho.toFixed(1)}%</strong> ({data.acertos} acertos)</p>
                                        <p>Tempo de Prova: <strong>{data.tempoStr}</strong> ({data.tempoMinutos} min)</p>
                                        <p style={{ marginTop: '0.25rem', color: 'var(--text-secondary)' }}>Classificação: {data.nivel}</p>
                                      </div>
                                    );
                                  }
                                  return null;
                                }}
                              />
                              <Scatter name="Estudantes" data={statsCalculated.dadosDispersao}>
                                {statsCalculated.dadosDispersao.map((entry, index) => {
                                  let color = 'var(--accent-cyan)';
                                  if (entry.nivel === 'Excelente') color = 'var(--accent-emerald)';
                                  else if (entry.nivel === 'Bom') color = 'var(--accent-cyan)';
                                  else if (entry.nivel === 'Regular') color = 'var(--accent-amber)';
                                  else color = 'var(--accent-rose)';
                                  return <Cell key={`cell-${index}`} fill={color} />;
                                })}
                              </Scatter>
                            </ScatterChart>
                          </ResponsiveContainer>
                        </div>
                      </div>

                    </div>

                    {/* Simplified Ranking Table */}
                    <div className="glass-panel table-card animate-fade-in" style={{ marginTop: '0.5rem' }}>
                      <div className="chart-header" style={{ padding: '1.5rem 1.5rem 0.5rem 1.5rem' }}>
                        <h3 className="chart-title">
                          <span style={{ backgroundColor: 'var(--accent-orange)' }}></span>
                          Classificação Geral da Turma
                        </h3>
                        <p className="chart-sub">Lista ranqueada de todos os estudantes. Clique em "Ver Análise" para abrir o detalhamento individual.</p>
                      </div>
                      <div className="table-wrapper">
                        <table className="students-table">
                          <thead>
                            <tr>
                              <th className="student-rank-cell">Rank</th>
                              <th>Nome do Aluno</th>
                              <th>Matrícula</th>
                              <th style={{ textAlign: 'center' }}>Desempenho</th>
                              <th style={{ textAlign: 'center' }}>Acertos / Erros</th>
                              <th style={{ textAlign: 'center' }}>Tempo</th>
                              <th style={{ textAlign: 'center' }}>Ação</th>
                            </tr>
                          </thead>
                          <tbody>
                            {currentData.alunos.map((aluno, index) => {
                              const rank = index + 1;
                              return (
                                <tr key={aluno.matricula} className="student-row">
                                  <td className="student-rank-cell">
                                    {rank === 1 ? (
                                      <span className="badge-rank gold" title="1º Lugar">🥇</span>
                                    ) : rank === 2 ? (
                                      <span className="badge-rank silver" title="2º Lugar">🥈</span>
                                    ) : rank === 3 ? (
                                      <span className="badge-rank bronze" title="3º Lugar">🥉</span>
                                    ) : (
                                      <span className="badge-rank normal">{rank}º</span>
                                    )}
                                  </td>
                                  <td>
                                    <div className="student-name" style={{ cursor: 'default' }}>
                                      {aluno.nome}
                                    </div>
                                  </td>
                                  <td className="student-matricula">{aluno.matricula}</td>
                                  <td style={{ textAlign: 'center' }}>
                                    <span className={`badge-score ${getScoreClass(aluno.desempenho)}`}>
                                      {aluno.desempenho.toFixed(1)}%
                                    </span>
                                  </td>
                                  <td style={{ textAlign: 'center' }}>
                                    <span style={{ color: 'var(--accent-emerald)', fontWeight: 'bold' }}>{aluno.acertos}</span>
                                    <span style={{ color: 'var(--text-muted)', margin: '0 0.25rem' }}>/</span>
                                    <span style={{ color: 'var(--accent-rose)', fontWeight: 'bold' }}>{aluno.erros}</span>
                                  </td>
                                  <td style={{ textAlign: 'center' }}>
                                    <span className="badge-time">{aluno.tempo}</span>
                                  </td>
                                  <td style={{ textAlign: 'center' }}>
                                    <button
                                      onClick={() => {
                                        setSelectedAlunoMatricula(aluno.matricula);
                                        setActiveSubTab('individual');
                                        window.scrollTo({ top: 0, behavior: 'smooth' });
                                      }}
                                      className="btn-help"
                                      style={{ padding: '0.35rem 0.75rem', borderColor: 'var(--accent-cyan-border)', color: 'var(--accent-cyan)', background: 'var(--accent-cyan-glow)' }}
                                    >
                                      Ver Análise <ArrowRight size={12} style={{ marginLeft: '0.25rem' }} />
                                    </button>
                                  </td>
                                </tr>
                              );
                            })}
                          </tbody>
                        </table>
                      </div>
                    </div>

                  </div>
                )}

                {activeSubTab === 'individual' && (
                  <div style={{ display: 'flex', flexDirection: 'column', gap: '1.5rem' }}>
                    {!currentData.possuiAbaRegistro ? (
                      <div className="glass-panel animate-fade-in" style={{ padding: '3rem', textAlign: 'center', color: 'var(--text-secondary)' }}>
                        <FileSpreadsheet size={40} style={{ color: 'var(--text-muted)', marginBottom: '1rem', margin: '0 auto' }} />
                        <h3>Registros de Detalhamento Indisponíveis</h3>
                        <p style={{ maxWidth: '400px', margin: '0.5rem auto' }}>
                          Esta planilha não contém dados de gabarito detalhado ou registros de respostas. Importe uma planilha completa (contendo a aba "por Registro") para ver o detalhamento de questões e perfil individual por aluno.
                        </p>
                      </div>
                    ) : (
                      <>
                        {/* Student Selector Card */}
                        <div className="glass-panel student-selector-card animate-fade-in" style={{ padding: '1.25rem', display: 'flex', flexWrap: 'wrap', gap: '1rem', alignItems: 'center', justifyContent: 'space-between' }}>
                          <div style={{ display: 'flex', alignItems: 'center', gap: '0.75rem', flexGrow: 1, minWidth: '280px' }}>
                            <Users size={20} style={{ color: 'var(--accent-cyan)' }} />
                            <div style={{ flexGrow: 1 }}>
                              <label style={{ display: 'block', fontSize: '0.65rem', textTransform: 'uppercase', color: 'var(--text-secondary)', fontWeight: 'bold', marginBottom: '0.25rem' }}>Selecionar Estudante</label>
                              <select
                                value={currentAluno ? currentAluno.matricula : ''}
                                onChange={(e) => {
                                  setSelectedAlunoMatricula(e.target.value);
                                  setAlunoQuestaoFilter('todas'); // reset filter
                                }}
                                className="filter-select"
                                style={{ width: '100%', maxWidth: '400px' }}
                              >
                                {currentData.alunos.map((a, idx) => (
                                  <option key={a.matricula} value={a.matricula}>
                                    {idx + 1}º - {a.nome} ({a.matricula}) - {a.desempenho.toFixed(1)}%
                                  </option>
                                ))}
                              </select>
                            </div>
                          </div>

                          {/* Navigation buttons */}
                          <div style={{ display: 'flex', gap: '0.5rem' }}>
                            <button
                              onClick={() => {
                                const currentIndex = currentData.alunos.findIndex(a => a.matricula === currentAluno.matricula);
                                if (currentIndex > 0) {
                                  setSelectedAlunoMatricula(currentData.alunos[currentIndex - 1].matricula);
                                }
                              }}
                              disabled={currentData.alunos.findIndex(a => a.matricula === currentAluno.matricula) === 0}
                              className="btn-clear"
                              style={{ padding: '0.5rem 0.75rem' }}
                            >
                              Anterior
                            </button>
                            <button
                              onClick={() => {
                                const currentIndex = currentData.alunos.findIndex(a => a.matricula === currentAluno.matricula);
                                if (currentIndex < currentData.alunos.length - 1) {
                                  setSelectedAlunoMatricula(currentData.alunos[currentIndex + 1].matricula);
                                }
                              }}
                              disabled={currentData.alunos.findIndex(a => a.matricula === currentAluno.matricula) === currentData.alunos.length - 1}
                              className="btn-clear"
                              style={{ padding: '0.5rem 0.75rem' }}
                            >
                              Próximo
                            </button>
                          </div>
                        </div>

                        {currentAluno && (
                          <div className="student-profile-dashboard animate-fade-in">
                            <div className="student-dashboard-grid">
                              {/* Card 1: Avatar Profile */}
                              <div className="glass-panel student-profile-card">
                                <div className="student-avatar-box">
                                  <div className="student-avatar">
                                    {currentAluno.nome.split(' ').map(n => n[0]).slice(0, 2).join('').toUpperCase()}
                                  </div>
                                  <div>
                                    <h3 className="student-profile-name">{currentAluno.nome}</h3>
                                    <p className="student-profile-sub">Matrícula: {currentAluno.matricula}</p>
                                  </div>
                                </div>
                                <div className="student-profile-metrics-list">
                                  <div className="profile-metric-item">
                                    <span className="profile-metric-label">Rank na Turma</span>
                                    <span className="profile-metric-val">
                                      {currentData.alunos.findIndex(a => a.matricula === currentAluno.matricula) + 1}º / {currentData.alunos.length}
                                    </span>
                                  </div>
                                  <div className="profile-metric-item">
                                    <span className="profile-metric-label">Tempo de Realização</span>
                                    <span className="profile-metric-val">{currentAluno.tempo}</span>
                                  </div>
                                  <div className="profile-metric-item">
                                    <span className="profile-metric-label">Status de Performance</span>
                                    <span className={`badge-score ${getScoreClass(currentAluno.desempenho)}`}>
                                      {currentAluno.desempenho >= 90 ? 'Excelente' : currentAluno.desempenho >= 70 ? 'Bom' : currentAluno.desempenho >= 50 ? 'Regular' : 'Insuficiente'}
                                    </span>
                                  </div>
                                </div>
                              </div>

                              {/* Card 2: Visual comparison */}
                              <div className="glass-panel comparison-card" style={{ padding: '1.25rem' }}>
                                <h4 className="card-title-new" style={{ fontSize: '0.85rem', color: 'white', marginBottom: '1rem', fontWeight: 'bold' }}>
                                  Comparativo de Aproveitamento vs. Turma
                                </h4>
                                <div className="comparison-container" style={{ display: 'flex', flexDirection: 'column', gap: '1.25rem' }}>
                                  
                                  {/* Metric 1: Desempenho */}
                                  <div className="comparison-item">
                                    <div className="comparison-header" style={{ display: 'flex', justifyContent: 'space-between', fontSize: '0.75rem', marginBottom: '0.35rem' }}>
                                      <span style={{ color: 'var(--text-secondary)' }}>Aproveitamento Geral</span>
                                      <span style={{ color: 'var(--accent-cyan)', fontWeight: 'bold' }}>{currentAluno.desempenho.toFixed(1)}%</span>
                                    </div>
                                    <div className="progress-bar-bg" style={{ height: '8px', background: 'rgba(255,255,255,0.06)', borderRadius: '4px', position: 'relative' }}>
                                      <div className="progress-bar-fill student-fill" style={{ height: '100%', borderRadius: '4px', width: `${currentAluno.desempenho}%`, background: 'linear-gradient(90deg, var(--accent-cyan), var(--accent-violet))' }}></div>
                                      <div className="progress-bar-marker" style={{ position: 'absolute', top: '-4px', width: '3px', height: '16px', background: 'var(--accent-amber)', borderRadius: '2px', left: `${currentData.desempenho}%` }} title={`Média da Turma: ${currentData.desempenho.toFixed(1)}%`}></div>
                                    </div>
                                    <div className="comparison-footer" style={{ display: 'flex', justifyContent: 'space-between', fontSize: '0.65rem', color: 'var(--text-muted)', marginTop: '0.25rem' }}>
                                      <span>Aluno: {currentAluno.desempenho.toFixed(1)}%</span>
                                      <span>Média Turma: {currentData.desempenho.toFixed(1)}%</span>
                                    </div>
                                  </div>

                                  {/* Metric 2: Acertos */}
                                  <div className="comparison-item">
                                    <div className="comparison-header" style={{ display: 'flex', justifyContent: 'space-between', fontSize: '0.75rem', marginBottom: '0.35rem' }}>
                                      <span style={{ color: 'var(--text-secondary)' }}>Questões Acertadas</span>
                                      <span style={{ color: 'var(--accent-emerald)', fontWeight: 'bold' }}>{currentAluno.acertos} / {currentData.totalItens}</span>
                                    </div>
                                    <div className="progress-bar-bg" style={{ height: '8px', background: 'rgba(255,255,255,0.06)', borderRadius: '4px', position: 'relative' }}>
                                      <div className="progress-bar-fill student-fill" style={{ height: '100%', borderRadius: '4px', width: `${(currentAluno.acertos / currentData.totalItens) * 100}%`, background: 'linear-gradient(90deg, var(--accent-emerald), #059669)' }}></div>
                                      <div className="progress-bar-marker" style={{ position: 'absolute', top: '-4px', width: '3px', height: '16px', background: 'var(--accent-amber)', borderRadius: '2px', left: `${(currentData.acertosMedio / currentData.totalItens) * 100}%` }} title={`Média da Turma: ${currentData.acertosMedio.toFixed(1)} acertos`}></div>
                                    </div>
                                    <div className="comparison-footer" style={{ display: 'flex', justifyContent: 'space-between', fontSize: '0.65rem', color: 'var(--text-muted)', marginTop: '0.25rem' }}>
                                      <span>Aluno: {currentAluno.acertos} acertos</span>
                                      <span>Média Turma: {currentData.acertosMedio.toFixed(1)} acertos</span>
                                    </div>
                                  </div>

                                  {/* Metric 3: Tempo */}
                                  {(() => {
                                    const alunoSegundos = currentAluno.tempo.split(':').length === 3 
                                      ? parseInt(currentAluno.tempo.split(':')[0]) * 3600 + parseInt(currentAluno.tempo.split(':')[1]) * 60 + parseInt(currentAluno.tempo.split(':')[2])
                                      : 0;
                                    const avgSegundos = statsCalculated.tempoMedioSegundos;
                                    const timeAnalysis = getClassificacaoTempo(currentAluno.tempo, avgSegundos);
                                    const maxTempo = Math.max(alunoSegundos, avgSegundos, 1);
                                    
                                    return (
                                      <div className="comparison-item">
                                        <div className="comparison-header" style={{ display: 'flex', justifyContent: 'space-between', fontSize: '0.75rem', marginBottom: '0.35rem' }}>
                                          <span style={{ color: 'var(--text-secondary)' }}>Tempo de Realização</span>
                                          <span className={`time-speed-label ${timeAnalysis.colorClass}`} style={{ display: 'inline', margin: 0 }}>
                                            {currentAluno.tempo} ({timeAnalysis.label})
                                          </span>
                                        </div>
                                        <div className="progress-bar-bg" style={{ height: '8px', background: 'rgba(255,255,255,0.06)', borderRadius: '4px', position: 'relative' }}>
                                          <div className="progress-bar-fill student-fill" style={{ height: '100%', borderRadius: '4px', width: `${(alunoSegundos / maxTempo) * 100}%`, background: 'linear-gradient(90deg, #64748b, #475569)' }}></div>
                                          <div className="progress-bar-marker" style={{ position: 'absolute', top: '-4px', width: '3px', height: '16px', background: 'var(--accent-amber)', borderRadius: '2px', left: `${(avgSegundos / maxTempo) * 100}%` }} title={`Média da Turma: ${statsCalculated.tempoMedioStr}`}></div>
                                        </div>
                                        <div className="comparison-footer" style={{ display: 'flex', justifyContent: 'space-between', fontSize: '0.65rem', color: 'var(--text-muted)', marginTop: '0.25rem' }}>
                                          <span>Aluno: {currentAluno.tempo}</span>
                                          <span>Média Turma: {statsCalculated.tempoMedioStr}</span>
                                        </div>
                                      </div>
                                    );
                                  })()}
                                </div>
                              </div>

                              {/* Card 3: Diagnosis */}
                              {(() => {
                                const pontosFortes = assuntoStats.filter(s => s.percent >= 70);
                                const pontosAtencao = assuntoStats.filter(s => s.percent < 70);
                                return (
                                  <div className="glass-panel diagnostic-card" style={{ padding: '1.25rem' }}>
                                    <h4 className="card-title-new" style={{ fontSize: '0.85rem', color: 'white', marginBottom: '0.75rem', fontWeight: 'bold' }}>
                                      Diagnóstico Pedagógico por Tópico
                                    </h4>
                                    <div className="diagnostic-lists">
                                      <div className="diagnostic-section">
                                        <span className="diagnostic-header-label success" style={{ display: 'block', fontSize: '0.65rem', fontWeight: 'bold', textTransform: 'uppercase', color: 'var(--accent-emerald)', marginBottom: '0.4rem' }}>
                                          🟢 Fortalezas (Domínio ≥70%)
                                        </span>
                                        {pontosFortes.length === 0 ? (
                                          <p className="diagnostic-empty" style={{ fontSize: '0.65rem', color: 'var(--text-muted)', paddingLeft: '0.5rem' }}>Nenhum tópico com domínio ≥70% nesta avaliação.</p>
                                        ) : (
                                          <ul className="diagnostic-list" style={{ listStyle: 'none', display: 'flex', flexDirection: 'column', gap: '0.25rem', paddingLeft: '0.5rem' }}>
                                            {pontosFortes.map(p => (
                                              <li key={p.assunto} className="diagnostic-item success" style={{ display: 'flex', justifyContent: 'space-between', fontSize: '0.7rem', borderBottom: '1px solid rgba(255,255,255,0.02)', paddingBottom: '0.2rem' }}>
                                                <span style={{ color: '#fff', textOverflow: 'ellipsis', whiteSpace: 'nowrap', overflow: 'hidden', maxWidth: '180px' }} title={p.assunto}>{p.assunto}</span>
                                                <strong style={{ color: 'var(--accent-emerald)', whiteSpace: 'nowrap' }}>{p.percent}% ({p.acertos}/{p.total})</strong>
                                              </li>
                                            ))}
                                          </ul>
                                        )}
                                      </div>
                                      <div className="diagnostic-section" style={{ marginTop: '0.75rem' }}>
                                        <span className="diagnostic-header-label danger" style={{ display: 'block', fontSize: '0.65rem', fontWeight: 'bold', textTransform: 'uppercase', color: 'var(--accent-rose)', marginBottom: '0.4rem' }}>
                                          🔴 Atenção (Domínio &lt;70%)
                                        </span>
                                        {pontosAtencao.length === 0 ? (
                                          <p className="diagnostic-empty" style={{ fontSize: '0.65rem', color: 'var(--text-muted)', paddingLeft: '0.5rem' }}>Excelente! Nenhum tópico com aproveitamento crítico.</p>
                                        ) : (
                                          <ul className="diagnostic-list" style={{ listStyle: 'none', display: 'flex', flexDirection: 'column', gap: '0.25rem', paddingLeft: '0.5rem' }}>
                                            {pontosAtencao.map(p => (
                                              <li key={p.assunto} className="diagnostic-item danger" style={{ display: 'flex', justifyContent: 'space-between', fontSize: '0.7rem', borderBottom: '1px solid rgba(255,255,255,0.02)', paddingBottom: '0.2rem' }}>
                                                <span style={{ color: '#fff', textOverflow: 'ellipsis', whiteSpace: 'nowrap', overflow: 'hidden', maxWidth: '180px' }} title={p.assunto}>{p.assunto}</span>
                                                <strong style={{ color: 'var(--accent-rose)', whiteSpace: 'nowrap' }}>{p.percent}% ({p.acertos}/{p.total})</strong>
                                              </li>
                                            ))}
                                          </ul>
                                        )}
                                      </div>
                                    </div>
                                  </div>
                                );
                              })()}
                            </div>

                            {/* Response Filter Buttons */}
                            <div className="responses-section-header" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginTop: '2rem', marginBottom: '1rem', borderBottom: '1px solid rgba(255, 255, 255, 0.05)', paddingBottom: '0.5rem' }}>
                              <h3 style={{ fontSize: '1rem', color: 'white', fontWeight: 'bold' }}>Caderno de Prova & Respostas</h3>
                              <div style={{ display: 'flex', background: 'rgba(0,0,0,0.2)', padding: '0.25rem', borderRadius: '8px', border: '1px solid rgba(255,255,255,0.04)' }}>
                                <button
                                  onClick={() => setAlunoQuestaoFilter('todas')}
                                  className={`btn-filter-sub ${alunoQuestaoFilter === 'todas' ? 'active' : ''}`}
                                  style={{ padding: '0.35rem 0.75rem', fontSize: '0.65rem', border: 'none', background: alunoQuestaoFilter === 'todas' ? 'var(--bg-secondary)' : 'transparent', color: alunoQuestaoFilter === 'todas' ? 'var(--accent-cyan)' : 'var(--text-secondary)', borderRadius: '6px', cursor: 'pointer', fontWeight: '700', transition: 'var(--transition-fast)' }}
                                >
                                  Ver Todas ({currentAluno.questoes.length})
                                </button>
                                <button
                                  onClick={() => setAlunoQuestaoFilter('acertos')}
                                  className={`btn-filter-sub ${alunoQuestaoFilter === 'acertos' ? 'active' : ''}`}
                                  style={{ padding: '0.35rem 0.75rem', fontSize: '0.65rem', border: 'none', background: alunoQuestaoFilter === 'acertos' ? 'var(--bg-secondary)' : 'transparent', color: alunoQuestaoFilter === 'acertos' ? 'var(--accent-emerald)' : 'var(--text-secondary)', borderRadius: '6px', cursor: 'pointer', fontWeight: '700', transition: 'var(--transition-fast)' }}
                                >
                                  Acertos ({currentAluno.acertos})
                                </button>
                                <button
                                  onClick={() => setAlunoQuestaoFilter('erros')}
                                  className={`btn-filter-sub ${alunoQuestaoFilter === 'erros' ? 'active' : ''}`}
                                  style={{ padding: '0.35rem 0.75rem', fontSize: '0.65rem', border: 'none', background: alunoQuestaoFilter === 'erros' ? 'var(--bg-secondary)' : 'transparent', color: alunoQuestaoFilter === 'erros' ? 'var(--accent-rose)' : 'var(--text-secondary)', borderRadius: '6px', cursor: 'pointer', fontWeight: '700', transition: 'var(--transition-fast)' }}
                                >
                                  Erros ({currentAluno.erros})
                                </button>
                              </div>
                            </div>

                            {/* Questions Cards Grid */}
                            {filteredQuestoesAluno.length === 0 ? (
                              <div className="glass-panel" style={{ padding: '3rem', textAlign: 'center', color: 'var(--text-secondary)' }}>
                                <CheckCircle2 size={32} style={{ color: 'var(--text-muted)', marginBottom: '0.5rem', margin: '0 auto' }} />
                                <p>Nenhuma questão atende aos filtros atuais para este aluno.</p>
                              </div>
                            ) : (
                              <div className="q-cards-grid-new">
                                {filteredQuestoesAluno.map((q, idx) => (
                                  <div key={q.identificador || idx} className={`q-card-new ${q.acertou ? 'success' : 'danger'}`}>
                                    <div className="q-card-badge-status">
                                      {q.acertou ? <CheckCircle2 size={16} /> : <XCircle size={16} />}
                                    </div>
                                    
                                    <div className="q-card-header-new">
                                      <span className="q-tag-new">Item #{q.identificador}</span>
                                      <span className={`difficulty-badge ${q.dificuldade?.toLowerCase() || 'medio'}`}>
                                        {q.dificuldade || 'Média'}
                                      </span>
                                    </div>

                                    <div className="q-card-body-new">
                                      {q.capacidade && (
                                        <p className="q-capacity"><strong>Capacidade:</strong> {q.capacidade}</p>
                                      )}
                                      {q.subfuncao && (
                                        <p className="q-skill"><strong>Subfunção:</strong> {q.subfuncao}</p>
                                      )}
                                      {q.padraoDesempenho && (
                                        <p className="q-pattern"><strong>Padrão de desempenho:</strong> {q.padraoDesempenho}</p>
                                      )}
                                      {q.conhecimento && (
                                        <p className="q-subject"><strong>Conhecimento:</strong> {q.conhecimento}</p>
                                      )}
                                    </div>

                                    <div className="q-card-footer-new">
                                      <div className={`answer-box ${q.acertou ? 'success' : 'warning'}`}>
                                        <span className="box-label">Resposta Aluno</span>
                                        <span className="box-letter">{q.marcacao || 'N/R'}</span>
                                      </div>
                                      {!q.acertou && (
                                        <div className="answer-box success">
                                          <span className="box-label">Gabarito</span>
                                          <span className="box-letter">{q.gabarito}</span>
                                        </div>
                                      )}
                                    </div>
                                  </div>
                                ))}
                              </div>
                            )}

                          </div>
                        )}
                      </>
                    )}
                  </div>
                )}

                {activeSubTab === 'registro' && (
                  <div style={{ display: 'flex', flexDirection: 'column', gap: '1.5rem' }}>
                    {!currentData.possuiAbaRegistro ? (
                      <div className="glass-panel animate-fade-in" style={{ padding: '3rem', textAlign: 'center', color: 'var(--text-secondary)' }}>
                        <FileSpreadsheet size={40} style={{ color: 'var(--text-muted)', marginBottom: '1rem', margin: '0 auto' }} />
                        <h3>Matriz de Itens Pedagógicos Indisponível</h3>
                        <p style={{ maxWidth: '400px', margin: '0.5rem auto' }}>
                          Esta planilha não contém dados de registros de respostas por questão. Importe uma planilha contendo a aba detalhada "por Registro" para habilitar a matriz de acertos e exportação de logs.
                        </p>
                      </div>
                    ) : (
                      <>
                        {/* Criticidade Metrics Grid */}
                        <div className="criticidade-grid" style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))', gap: '1rem' }}>
                          <div className="glass-panel kpi-card rose animate-fade-in">
                            <div>
                              <p className="kpi-label">Questões Críticas</p>
                              <h3 className="kpi-value">{criticidadeStats.criticas}</h3>
                              <p className="kpi-sub">Taxa de acertos abaixo de 40%</p>
                            </div>
                            <div className="kpi-icon-box" style={{ background: 'var(--accent-rose-glow)', color: 'var(--accent-rose)' }}>
                              <XCircle size={18} />
                            </div>
                          </div>
                          
                          <div className="glass-panel kpi-card amber animate-fade-in">
                            <div>
                              <p className="kpi-label">Questões em Atenção</p>
                              <h3 className="kpi-value">{criticidadeStats.atencao}</h3>
                              <p className="kpi-sub">Taxa de acertos de 40% a 70%</p>
                            </div>
                            <div className="kpi-icon-box" style={{ background: 'var(--accent-amber-glow)', color: 'var(--accent-amber)' }}>
                              <AlertCircle size={18} />
                            </div>
                          </div>

                          <div className="glass-panel kpi-card emerald animate-fade-in">
                            <div>
                              <p className="kpi-label">Questões Dominadas</p>
                              <h3 className="kpi-value">{criticidadeStats.dominadas}</h3>
                              <p className="kpi-sub">Taxa de acertos acima de 70%</p>
                            </div>
                            <div className="kpi-icon-box" style={{ background: 'var(--accent-emerald-glow)', color: 'var(--accent-emerald)' }}>
                              <CheckCircle2 size={18} />
                            </div>
                          </div>
                        </div>

                        {/* Matriz Pedagógica de Acertos por Item */}
                        <div className="glass-panel table-card animate-fade-in">
                          <div className="chart-header" style={{ padding: '1.5rem 1.5rem 0.5rem 1.5rem' }}>
                            <h3 className="chart-title">
                              <span style={{ backgroundColor: 'var(--accent-cyan)' }}></span>
                              Matriz Pedagógica de Desempenho por Item
                            </h3>
                            <p className="chart-sub">Mapeamento do nível de aproveitamento coletivo para intervenções pedagógicas direcionadas.</p>
                          </div>
                          
                          <div className="table-wrapper">
                            <table className="students-table">
                              <thead>
                                <tr>
                                  <th>Item / ID</th>
                                  <th>Capacidade</th>
                                  <th>Subfunção</th>
                                  <th>Padrão de desempenho</th>
                                  <th>Conhecimento</th>
                                  <th style={{ width: '80px', textAlign: 'center' }}>Dificuldade</th>
                                  <th style={{ width: '220px' }}>Taxa de Acertos da Turma</th>
                                  <th style={{ width: '120px', textAlign: 'center' }}>Classificação</th>
                                </tr>
                              </thead>
                              <tbody>
                                {matrizQuestoes.map((q) => (
                                  <tr key={q.identificador} className="student-row">
                                    <td style={{ fontWeight: 'bold', color: 'white' }}>Item #{q.identificador}</td>
                                    <td style={{ fontSize: '0.75rem', color: 'var(--text-primary)' }}>{q.capacidade}</td>
                                    <td style={{ fontSize: '0.75rem', color: 'var(--text-primary)' }}>{q.subfuncao}</td>
                                    <td style={{ fontSize: '0.75rem', color: 'var(--text-secondary)' }}>{q.padraoDesempenho}</td>
                                    <td style={{ fontSize: '0.75rem', color: 'var(--text-primary)' }}>{q.conhecimento}</td>
                                    <td style={{ textAlign: 'center' }}>
                                      <span className={`difficulty-badge ${q.dificuldade?.toLowerCase() || 'medio'}`}>
                                        {q.dificuldade || 'Média'}
                                      </span>
                                    </td>
                                    <td>
                                      <div style={{ display: 'flex', alignItems: 'center', gap: '0.5rem' }}>
                                        <div className="progress-bar-bg" style={{ flexGrow: 1, height: '6px', background: 'rgba(255,255,255,0.06)', borderRadius: '3px', position: 'relative' }}>
                                          <div 
                                            className="progress-bar-fill" 
                                            style={{ 
                                              height: '100%', 
                                              borderRadius: '3px', 
                                              width: `${q.taxa}%`, 
                                              background: q.status === 'critica' ? 'var(--accent-rose)' : q.status === 'atencao' ? 'var(--accent-amber)' : 'var(--accent-emerald)'
                                            }}
                                          ></div>
                                        </div>
                                        <span style={{ fontSize: '0.75rem', fontWeight: 'bold', color: q.status === 'critica' ? 'var(--accent-rose)' : q.status === 'atencao' ? 'var(--accent-amber)' : 'var(--accent-emerald)', width: '45px', textAlign: 'right' }}>
                                          {q.taxa.toFixed(1)}%
                                        </span>
                                      </div>
                                    </td>
                                    <td style={{ textAlign: 'center' }}>
                                      <span className={`badge-score ${q.status === 'critica' ? 'insuficiente' : q.status === 'atencao' ? 'regular' : 'excelente'}`} style={{ fontSize: '0.6rem', textTransform: 'uppercase', letterSpacing: '0.05em' }}>
                                        {q.status === 'critica' ? '🔴 Crítica' : q.status === 'atencao' ? '🟡 Atenção' : '🟢 Dominada'}
                                      </span>
                                    </td>
                                  </tr>
                                ))}
                              </tbody>
                            </table>
                          </div>
                        </div>

                        {/* Flat Log Table */}
                        <div className="glass-panel table-card animate-fade-in" style={{ marginTop: '0.5rem' }}>
                          <div className="chart-header" style={{ padding: '1.5rem 1.5rem 1rem 1.5rem', display: 'flex', flexWrap: 'wrap', gap: '1rem', alignItems: 'center', justifyContent: 'space-between', borderBottom: '1px solid rgba(255,255,255,0.03)' }}>
                            <div>
                              <h3 className="chart-title">
                                <span style={{ backgroundColor: 'var(--accent-violet)' }}></span>
                                Registro Plano de Respostas (Logs)
                              </h3>
                              <p className="chart-sub">Filtre, pesquise e exporte o log bruto de todas as interações da turma.</p>
                            </div>
                            
                            <div style={{ display: 'flex', gap: '0.75rem', flexWrap: 'wrap', alignItems: 'center' }}>
                              <div className="search-input-wrapper" style={{ width: '260px' }}>
                                <Search size={14} className="search-icon" />
                                <input
                                  type="text"
                                  value={registroSearch}
                                  onChange={(e) => {
                                    setRegistroSearch(e.target.value);
                                    setRegistroPage(1);
                                  }}
                                  placeholder="Pesquisar aluno, questão ou assunto..."
                                  className="search-input"
                                  style={{ padding: '0.5rem 0.75rem 0.5rem 2rem' }}
                                />
                              </div>
                              
                              <button
                                onClick={handleExportCSV}
                                className="btn-help"
                                style={{ display: 'flex', alignItems: 'center', gap: '0.35rem', borderColor: 'var(--accent-cyan-border)', color: 'var(--accent-cyan)', background: 'var(--accent-cyan-glow)', padding: '0.5rem 0.85rem' }}
                              >
                                <Download size={13} />
                                Exportar CSV ({filteredRegistros.length})
                              </button>
                            </div>
                          </div>

                          <div className="table-wrapper">
                            <table className="students-table">
                              <thead>
                                <tr>
                                  <th>Estudante</th>
                                  <th>Matrícula</th>
                                  <th style={{ width: '90px' }}>Item ID</th>
                                  <th>Capacidade</th>
                                  <th>Subfunção</th>
                                  <th>Padrão de desempenho</th>
                                  <th>Conhecimento</th>
                                  <th style={{ textAlign: 'center', width: '80px' }}>Alternativa</th>
                                  <th style={{ textAlign: 'center', width: '80px' }}>Gabarito</th>
                                  <th style={{ textAlign: 'center', width: '100px' }}>Resultado</th>
                                </tr>
                              </thead>
                              <tbody>
                                {paginatedRegistros.length === 0 ? (
                                  <tr>
                                    <td colSpan="10" style={{ textAlign: 'center', padding: '3rem', color: 'var(--text-muted)' }}>
                                      <Search size={32} style={{ marginBottom: '0.5rem', margin: '0 auto' }} />
                                      <p>Nenhum registro encontrado para a busca atual.</p>
                                    </td>
                                  </tr>
                                ) : (
                                  paginatedRegistros.map((r, idx) => (
                                    <tr key={`${r.alunoMatricula}-${r.identificador}-${idx}`} className="student-row">
                                      <td style={{ fontWeight: '600', color: 'white' }}>{r.alunoNome}</td>
                                      <td className="student-matricula">{r.alunoMatricula}</td>
                                      <td style={{ fontWeight: 'bold' }}>Item #{r.identificador}</td>
                                      <td style={{ fontSize: '0.75rem', color: 'var(--text-secondary)' }}>{r.capacidade}</td>
                                      <td style={{ fontSize: '0.75rem', color: 'var(--text-secondary)' }}>{r.subfuncao}</td>
                                      <td style={{ fontSize: '0.75rem', color: 'var(--text-secondary)' }}>{r.padraoDesempenho}</td>
                                      <td style={{ fontSize: '0.75rem', color: 'var(--text-secondary)' }}>{r.conhecimento}</td>
                                      <td style={{ textAlign: 'center', fontWeight: '800', color: r.acertou ? 'var(--accent-emerald)' : 'var(--accent-amber)' }}>
                                        {r.marcacao}
                                      </td>
                                      <td style={{ textAlign: 'center', fontWeight: '800', color: 'var(--accent-emerald)' }}>
                                        {r.gabarito}
                                      </td>
                                      <td style={{ textAlign: 'center' }}>
                                        <span className={`badge-score ${r.acertou ? 'excelente' : 'insuficiente'}`} style={{ padding: '0.15rem 0.5rem', fontSize: '0.6rem' }}>
                                          {r.acertou ? 'Acerto' : 'Erro'}
                                        </span>
                                      </td>
                                    </tr>
                                  ))
                                )}
                              </tbody>
                            </table>
                          </div>

                          {/* Pagination Controls Footer */}
                          {maxRegistroPages > 1 && (
                            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '1rem', borderTop: '1px solid rgba(255,255,255,0.03)', flexWrap: 'wrap', gap: '0.75rem' }}>
                              <span style={{ fontSize: '0.7rem', color: 'var(--text-secondary)' }}>
                                Exibindo <strong>{(registroPage - 1) * 25 + 1}</strong> - <strong>{Math.min(registroPage * 25, filteredRegistros.length)}</strong> de <strong>{filteredRegistros.length}</strong> registros
                              </span>
                              
                              <div style={{ display: 'flex', gap: '0.35rem', alignItems: 'center' }}>
                                <button
                                  onClick={() => setRegistroPage(prev => Math.max(1, prev - 1))}
                                  disabled={registroPage === 1}
                                  className="btn-clear"
                                  style={{ padding: '0.35rem 0.6rem', fontSize: '0.7rem' }}
                                >
                                  Anterior
                                </button>
                                <span style={{ fontSize: '0.7rem', color: 'white', padding: '0 0.5rem', fontFamily: 'monospace' }}>
                                  Página {registroPage} de {maxRegistroPages}
                                </span>
                                <button
                                  onClick={() => setRegistroPage(prev => Math.min(maxRegistroPages, prev + 1))}
                                  disabled={registroPage === maxRegistroPages}
                                  className="btn-clear"
                                  style={{ padding: '0.35rem 0.6rem', fontSize: '0.7rem' }}
                                >
                                  Próxima
                                </button>
                              </div>
                            </div>
                          )}
                        </div>
                      </>
                    )}
                  </div>
                )}

                {activeSubTab === 'relatorio' && (
                  <div style={{ display: 'flex', flexDirection: 'column', gap: '1.5rem' }} className="animate-fade-in">
                    {/* Header Card */}
                    <div className="glass-panel" style={{ padding: '2rem' }}>
                      <div style={{ display: 'flex', alignItems: 'center', gap: '1rem', marginBottom: '1rem' }}>
                        <div style={{ background: 'var(--accent-orange-glow)', color: 'var(--accent-orange)', padding: '0.75rem', borderRadius: '12px', border: '1px solid var(--accent-orange-border)', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
                          <Sparkles size={28} />
                        </div>
                        <div>
                          <h2 style={{ fontSize: '1.5rem', fontWeight: 800, color: 'white', margin: 0 }}>Relatório Pedagógico para IA</h2>
                          <p style={{ fontSize: '0.85rem', color: 'var(--text-secondary)', margin: '0.25rem 0 0 0' }}>
                            Gerador de prompts estruturados para você copiar e colar em IAs (Gemini, ChatGPT, Claude) criarem planos de revisão e roteiros pedagógicos personalizados.
                          </p>
                        </div>
                      </div>

                      {/* KPI Resumo do Relatório */}
                      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))', gap: '1rem', marginTop: '1.5rem' }}>
                        <div className="glass-panel kpi-card rose" style={{ padding: '1rem' }}>
                          <div>
                            <p className="kpi-label">Itens Críticos (&lt;50% acertos)</p>
                            <h3 className="kpi-value" style={{ fontSize: '1.25rem' }}>{matrizQuestoes.filter(q => q.taxa < 50).length} itens</h3>
                            <p className="kpi-sub">Inseridos para revisão prioritária</p>
                          </div>
                          <div className="kpi-icon-box" style={{ background: 'var(--accent-rose-glow)', color: 'var(--accent-rose)' }}>
                            <XCircle size={16} />
                          </div>
                        </div>

                        <div className="glass-panel kpi-card amber" style={{ padding: '1rem' }}>
                          <div>
                            <p className="kpi-label">Alunos sob Atenção (&lt;50% score)</p>
                            <h3 className="kpi-value" style={{ fontSize: '1.25rem' }}>{currentData.alunos.filter(a => a.desempenho < 50).length} alunos</h3>
                            <p className="kpi-sub">Foco de suporte individual</p>
                          </div>
                          <div className="kpi-icon-box" style={{ background: 'var(--accent-amber-glow)', color: 'var(--accent-amber)' }}>
                            <Users size={16} />
                          </div>
                        </div>

                        <div className="glass-panel kpi-card emerald" style={{ padding: '1rem' }}>
                          <div>
                            <p className="kpi-label">Assuntos Dominados (&gt;=70%)</p>
                            <h3 className="kpi-value" style={{ fontSize: '1.25rem' }}>{matrizQuestoes.filter(q => q.taxa >= 70).length} itens</h3>
                            <p className="kpi-sub">Consolidados pela turma</p>
                          </div>
                          <div className="kpi-icon-box" style={{ background: 'var(--accent-emerald-glow)', color: 'var(--accent-emerald)' }}>
                            <CheckCircle2 size={16} />
                          </div>
                        </div>
                      </div>
                      
                      {/* Prompt Action Area */}
                      <div style={{ marginTop: '2rem', display: 'flex', flexDirection: 'column', gap: '1rem' }}>
                        <div style={{ display: 'flex', flexWrap: 'wrap', gap: '1rem', alignItems: 'center', justifyContent: 'space-between' }}>
                          <span style={{ fontSize: '0.8rem', fontWeight: 'bold', color: 'var(--text-primary)', textTransform: 'uppercase', letterSpacing: '0.05em', display: 'flex', alignItems: 'center', gap: '0.4rem' }}>
                            <span style={{ width: '8px', height: '8px', borderRadius: '50%', backgroundColor: 'var(--accent-cyan)' }}></span>
                            Conteúdo do Prompt
                          </span>
                          <button
                            onClick={() => {
                              navigator.clipboard.writeText(iaPrompt);
                              setCopiedPrompt(true);
                              setTimeout(() => setCopiedPrompt(false), 3000);
                            }}
                            className="btn-select"
                            style={{ display: 'inline-flex', alignItems: 'center', gap: '0.5rem', padding: '0.75rem 1.5rem', fontSize: '0.8rem', fontWeight: 'bold', background: copiedPrompt ? 'var(--accent-emerald)' : 'linear-gradient(135deg, var(--accent-orange), #ff8533)', boxShadow: copiedPrompt ? '0 0 20px var(--accent-emerald-glow)' : 'var(--shadow-glow-orange)' }}
                          >
                            {copiedPrompt ? (
                              <>
                                <CheckCircle2 size={16} /> Copiado com sucesso!
                              </>
                            ) : (
                              <>
                                <Sparkles size={16} /> Copiar Prompt para IA
                              </>
                            )}
                          </button>
                        </div>

                        <div style={{ position: 'relative' }}>
                          <textarea
                            value={iaPrompt}
                            readOnly
                            style={{
                              width: '100%',
                              height: '350px',
                              background: 'rgba(0, 0, 0, 0.45)',
                              border: '1px solid var(--border-color)',
                              borderRadius: 'var(--radius-sm)',
                              color: 'var(--text-secondary)',
                              fontFamily: 'monospace',
                              fontSize: '0.75rem',
                              padding: '1.25rem',
                              resize: 'none',
                              outline: 'none',
                              lineHeight: '1.5'
                            }}
                          />
                          <div style={{ position: 'absolute', bottom: '1.5rem', right: '1.5rem', display: 'flex', gap: '0.5rem' }}>
                            <span style={{ fontSize: '0.65rem', background: 'rgba(255,255,255,0.05)', color: 'var(--text-muted)', padding: '0.3rem 0.6rem', borderRadius: '4px', border: '1px solid rgba(255,255,255,0.05)' }}>
                              {iaPrompt.length} caracteres
                            </span>
                          </div>
                        </div>
                      </div>
                    </div>

                    {/* How to use guidelines */}
                    <div className="glass-panel" style={{ padding: '1.5rem', display: 'flex', flexDirection: 'column', gap: '1rem' }}>
                      <h3 style={{ fontSize: '0.95rem', fontWeight: 700, color: 'white', display: 'flex', alignItems: 'center', gap: '0.5rem', margin: 0 }}>
                        <Target size={16} style={{ color: 'var(--accent-orange)' }} />
                        Como utilizar com a Inteligência Artificial?
                      </h3>
                      <ol style={{ fontSize: '0.8rem', color: 'var(--text-secondary)', paddingLeft: '1.25rem', margin: 0, lineHeight: '1.6', display: 'flex', flexDirection: 'column', gap: '0.75rem' }}>
                        <li>
                          Clique no botão <strong>"Copiar Prompt para IA"</strong> para salvar o texto estruturado com os dados pedagógicos reais da sua turma.
                        </li>
                        <li>
                          Acesse um chat de Inteligência Artificial de sua preferência (exemplos: <a href="https://gemini.google.com" target="_blank" rel="noopener noreferrer" style={{ color: 'var(--accent-cyan)', textDecoration: 'underline' }}>Google Gemini</a>, <a href="https://chatgpt.com" target="_blank" rel="noopener noreferrer" style={{ color: 'var(--accent-cyan)', textDecoration: 'underline' }}>OpenAI ChatGPT</a> ou Claude).
                        </li>
                        <li>
                          Cole o prompt copiado no campo de digitação e envie.
                        </li>
                        <li>
                          A IA agirá como consultor pedagógico e retornará um roteiro completo de intervenção focado estritamente nas capacidades e conhecimentos em defasagem da sua turma de <strong>{currentData.curso}</strong>.
                        </li>
                      </ol>
                    </div>
                  </div>
                )}
              </div>
            )}

          </div>
        )}
      </main>

      <footer className="footer animate-fade-in">
        <div className="container footer-content">
          <div className="footer-left">
            <p className="footer-copyright">
              © {new Date().getFullYear()} - Desenvolvido por <span className="footer-highlight">Peterson Ferreira</span>
            </p>
            <p className="footer-subtext">Dashboard SAEP - Portal de Acompanhamento e Análise Pedagógica Avançada</p>
          </div>
          <div className="footer-right">
            <span className="footer-info-badge">Versão 1.2.0</span>
            <span className="footer-info-badge">React & Vite</span>
            <span className="footer-info-badge">SENAI</span>
          </div>
        </div>
      </footer>
    </div>

      {/* Import Help Modal Dialog */}
      {showHelpModal && (
        <div className="help-modal-overlay">
          <div className="glass-panel help-modal">
            <button 
              onClick={() => setShowHelpModal(false)}
              className="help-close-btn"
            >
              <X size={16} />
            </button>
            
            <div className="help-modal-body">
              <h3 style={{ fontSize: '1rem', color: 'white', display: 'flex', alignItems: 'center', gap: '0.4rem', marginBottom: '0.75rem' }}>
                <FileSpreadsheet size={16} style={{ color: 'var(--accent-cyan)' }} />
                Como obter as planilhas do SAEP?
              </h3>
              
              <p style={{ marginBottom: '1rem' }}>
                O processador pedagógico interpreta a planilha no formato padrão gerado pelo sistema de relatórios do SAEP. Para habilitar todos os recursos, siga os passos abaixo:
              </p>
              
              <div className="help-steps-card">
                <h4>Instruções de Exportação:</h4>
                <ul className="help-steps-list">
                  <li>Acesse o portal oficial do SAEP com suas credenciais de gestor/docente.</li>
                  <li>Navegue até a aba de **Resultados** e selecione a avaliação desejada.</li>
                  <li>Clique em **Exportar Planilha de Desempenho** (certifique-se de exportar o arquivo contendo a aba detalhada por registros).</li>
                  <li>Arraste o arquivo `.xls` ou `.xlsx` diretamente na área de upload deste portal.</li>
                </ul>
              </div>

              <div style={{ marginTop: '1rem', padding: '0.75rem', background: 'rgba(6, 182, 212, 0.05)', borderRadius: '8px', border: '1px solid rgba(6, 182, 212, 0.15)', display: 'flex', gap: '0.5rem', alignItems: 'flex-start' }}>
                <Sparkles size={16} style={{ color: 'var(--accent-cyan)', shrink: 0, marginTop: '0.15rem' }} />
                <p style={{ margin: 0, fontSize: '0.7rem', color: 'rgba(255, 255, 255, 0.8)' }}>
                  <strong>Múltiplos Cursos:</strong> Você pode arrastar arquivos de diferentes turmas e cursos simultaneamente. O portal irá gerar abas superiores para que você alterne e compare os resultados de forma dinâmica.
                </p>
              </div>

              <button
                onClick={() => setShowHelpModal(false)}
                className="btn-help-modal-ok"
              >
                Entendido, prosseguir
              </button>
            </div>
          </div>
        </div>
      )}

    </div>
  );
}
