// Configuração inicial
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setTitle('Sistema de Armários Hospitalares');
}

function doPost(e) {
  return handlePost(e);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ID da pasta do Drive para salvar os PDFs - ATUALIZE COM SEU ID
const PASTA_DRIVE_ID = '1nYsGJJUIufxDYVvIanVXCbPx7YuBOYDP';

// Inicializar planilha com todas as abas e cabeçalhos
function inicializarPlanilha() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Criar abas se não existirem
    var abas = [
      { 
        nome: 'Histórico Visitantes', 
        cabecalhos: ['ID', 'Data', 'Número Armário', 'Nome Visitante', 'Nome Paciente', 'Leito', 'Volumes', 'Hora Início', 'Hora Fim', 'Status', 'Tipo', 'Unidade'] 
      },
      { 
        nome: 'Histórico Acompanhantes', 
        cabecalhos: ['ID', 'Data', 'Número Armário', 'Nome Acompanhante', 'Nome Paciente', 'Leito', 'Volumes', 'Hora Início', 'Hora Fim', 'Status', 'Tipo', 'Unidade'] 
      },
      { 
        nome: 'Visitantes', 
        cabecalhos: ['ID', 'Número', 'Status', 'Nome Visitante', 'Nome Paciente', 'Leito', 'Volumes', 'Hora Início', 'Hora Prevista', 'Data Registro', 'Unidade', 'Termo Aplicado'] 
      },
      { 
        nome: 'Acompanhantes', 
        cabecalhos: ['ID', 'Número', 'Status', 'Nome Acompanhante', 'Nome Paciente', 'Leito', 'Volumes', 'Hora Início', 'Data Registro', 'Unidade', 'Termo Aplicado'] 
      },
      { 
        nome: 'Cadastro Armários', 
        cabecalhos: ['ID', 'Número', 'Tipo', 'Unidade', 'Localização', 'Status', 'Data Cadastro'] 
      },
      { 
        nome: 'Unidades', 
        cabecalhos: ['ID', 'Nome', 'Status', 'Data Cadastro'] 
      },
      { 
        nome: 'Usuários', 
        cabecalhos: ['ID', 'Nome', 'Email', 'Perfil', 'Acesso Visitantes', 'Acesso Acompanhantes', 'Data Cadastro', 'Status'] 
      },
      { 
        nome: 'LOGS', 
        cabecalhos: ['Data/Hora', 'Usuário', 'Ação', 'Detalhes', 'IP'] 
      },
      { 
        nome: 'Termos de Responsabilidade', 
        cabecalhos: ['ID', 'ArmarioID', 'NumeroArmario', 'Paciente', 'Prontuario', 'Nascimento', 'Setor', 'Leito', 'Consciente', 'Acompanhante', 'Telefone', 'Documento', 'Parentesco', 'Orientacoes', 'Volumes', 'DescricaoVolumes', 'AplicadoEm', 'PDF_URL', 'AssinaturaBase64'] 
      },
      { 
        nome: 'Movimentações', 
        cabecalhos: ['ID', 'ArmarioID', 'NumeroArmario', 'Tipo', 'Descricao', 'Responsavel', 'Data', 'Hora', 'DataHoraRegistro'] 
      }
    ];
    
    abas.forEach(function(aba) {
      var sheet = ss.getSheetByName(aba.nome);
      if (!sheet) {
        sheet = ss.insertSheet(aba.nome);
        sheet.getRange(1, 1, 1, aba.cabecalhos.length).setValues([aba.cabecalhos]);
        sheet.setFrozenRows(1);
        
        // Formatar cabeçalhos
        var headerRange = sheet.getRange(1, 1, 1, aba.cabecalhos.length);
        headerRange.setBackground('#2c6e8f')
          .setFontColor('white')
          .setFontWeight('bold');
      }
    });
    
    // Adicionar alguns dados iniciais de exemplo
    adicionarDadosIniciais();
    
    registrarLog('SISTEMA', 'Planilha inicializada com sucesso');
    return { success: true, message: 'Planilha inicializada com sucesso' };
    
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

// Adicionar dados iniciais de exemplo
function adicionarDadosIniciais() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Cadastrar alguns armários físicos
  var cadastroSheet = ss.getSheetByName('Cadastro Armários');
  if (cadastroSheet.getLastRow() === 1) {
    var armariosIniciais = [
      ['V-01', 'visitante', 'NAC Eletiva', 'Bloco A - Térreo', 'ativo', new Date()],
      ['V-02', 'visitante', 'NAC Eletiva', 'Bloco A - Térreo', 'ativo', new Date()],
      ['V-03', 'visitante', 'UIB', 'Bloco A - Térreo', 'ativo', new Date()],
      ['V-04', 'visitante', 'UIB', 'Bloco A - Térreo', 'ativo', new Date()],
      ['A-01', 'acompanhante', 'NAC Eletiva', 'Bloco B - 1º Andar', 'ativo', new Date()],
      ['A-02', 'acompanhante', 'UIB', 'Bloco B - 1º Andar', 'ativo', new Date()],
      ['A-03', 'acompanhante', 'UIB', 'Bloco B - 1º Andar', 'ativo', new Date()]
    ];

    armariosIniciais.forEach(function(armario, index) {
      cadastroSheet.getRange(cadastroSheet.getLastRow() + 1, 1, 1, 7)
        .setValues([[index + 1, ...armario]]);
    });

    criarArmariosUso(armariosIniciais.map((armario, index) => [index + 1, ...armario]));
  }

  // Cadastrar usuário admin inicial
  var usuariosSheet = ss.getSheetByName('Usuários');
  if (usuariosSheet.getLastRow() === 1) {
    usuariosSheet.getRange(2, 1, 1, 8)
      .setValues([[1, 'Administrador', 'admin@hospital.com', 'admin', true, true, new Date(), 'ativo']]);
  }

  // Cadastrar unidades iniciais
  var unidadesSheet = ss.getSheetByName('Unidades');
  if (unidadesSheet && unidadesSheet.getLastRow() === 1) {
    var unidadesIniciais = [
      [1, 'NAC Eletiva', 'ativa', new Date()],
      [2, 'UIB', 'ativa', new Date()]
    ];
    unidadesSheet.getRange(2, 1, unidadesIniciais.length, 4).setValues(unidadesIniciais);
  }
}

// Função principal para lidar com requisições POST
function handlePost(e) {
  var action = e.parameter.action;
  
  try {
    switch(action) {
      case 'getArmarios':
        return ContentService.createTextOutput(JSON.stringify(getArmarios(e.parameter.tipo)))
          .setMimeType(ContentService.MimeType.JSON);
      
      case 'cadastrarArmario':
        return ContentService.createTextOutput(JSON.stringify(cadastrarArmario(e.parameter)))
          .setMimeType(ContentService.MimeType.JSON);
      
      case 'liberarArmario':
        return ContentService.createTextOutput(JSON.stringify(liberarArmario(e.parameter.id, e.parameter.tipo)))
          .setMimeType(ContentService.MimeType.JSON);
      
      case 'getUsuarios':
        return ContentService.createTextOutput(JSON.stringify(getUsuarios()))
          .setMimeType(ContentService.MimeType.JSON);
      
      case 'cadastrarUsuario':
        return ContentService.createTextOutput(JSON.stringify(cadastrarUsuario(e.parameter)))
          .setMimeType(ContentService.MimeType.JSON);
      
      case 'getLogs':
        return ContentService.createTextOutput(JSON.stringify(getLogs()))
          .setMimeType(ContentService.MimeType.JSON);
      
      case 'getNotificacoes':
        return ContentService.createTextOutput(JSON.stringify(getNotificacoes()))
          .setMimeType(ContentService.MimeType.JSON);
      
      case 'getEstatisticasDashboard':
        return ContentService.createTextOutput(JSON.stringify(getEstatisticasDashboard(e.parameter.tipoUsuario)))
          .setMimeType(ContentService.MimeType.JSON);
      
      case 'getHistorico':
        return ContentService.createTextOutput(JSON.stringify(getHistorico(e.parameter.tipo)))
          .setMimeType(ContentService.MimeType.JSON);
      
      case 'getCadastroArmarios':
        return ContentService.createTextOutput(JSON.stringify(getCadastroArmarios()))
          .setMimeType(ContentService.MimeType.JSON);
      
      case 'cadastrarArmarioFisico':
        return ContentService.createTextOutput(JSON.stringify(cadastrarArmarioFisico(e.parameter)))
          .setMimeType(ContentService.MimeType.JSON);
      
      case 'getUnidades':
        return ContentService.createTextOutput(JSON.stringify(getUnidades()))
          .setMimeType(ContentService.MimeType.JSON);
      
      case 'cadastrarUnidade':
        return ContentService.createTextOutput(JSON.stringify(cadastrarUnidade(e.parameter)))
          .setMimeType(ContentService.MimeType.JSON);
      
      case 'alternarStatusUnidade':
        return ContentService.createTextOutput(JSON.stringify(alternarStatusUnidade(e.parameter)))
          .setMimeType(ContentService.MimeType.JSON);
      
      case 'salvarTermoCompleto':
        return ContentService.createTextOutput(JSON.stringify(salvarTermoCompleto(e.parameter)))
          .setMimeType(ContentService.MimeType.JSON);
      
      case 'getTermo':
        return ContentService.createTextOutput(JSON.stringify(getTermo(e.parameter)))
          .setMimeType(ContentService.MimeType.JSON);
      
      case 'getMovimentacoes':
        return ContentService.createTextOutput(JSON.stringify(getMovimentacoes(e.parameter)))
          .setMimeType(ContentService.MimeType.JSON);
      
      case 'salvarMovimentacao':
        return ContentService.createTextOutput(JSON.stringify(salvarMovimentacao(e.parameter)))
          .setMimeType(ContentService.MimeType.JSON);
      
      case 'verificarInicializacao':
        return ContentService.createTextOutput(JSON.stringify(verificarInicializacao()))
          .setMimeType(ContentService.MimeType.JSON);
      
      case 'inicializarPlanilha':
        return ContentService.createTextOutput(JSON.stringify(inicializarPlanilha()))
          .setMimeType(ContentService.MimeType.JSON);
      
      default:
        return ContentService.createTextOutput(JSON.stringify({ success: false, error: 'Ação não reconhecida: ' + action }))
          .setMimeType(ContentService.MimeType.JSON);
    }
  } catch (error) {
    registrarLog('ERRO', `Erro em handlePost: ${error.toString()}`);
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Funções para Armários
function getArmarios(tipo) {
  try {
    if (tipo === 'admin' || tipo === 'ambos') {
      let visitantes = getArmariosFromSheet('Visitantes', 'visitante');
      let acompanhantes = getArmariosFromSheet('Acompanhantes', 'acompanhante');
      return { success: true, data: visitantes.concat(acompanhantes) };
    } else {
      let sheetName = tipo === 'acompanhante' ? 'Acompanhantes' : 'Visitantes';
      return { success: true, data: getArmariosFromSheet(sheetName, tipo) };
    }
  } catch (error) {
    registrarLog('ERRO', `Erro ao buscar armários: ${error.toString()}`);
    return { success: false, error: error.toString() };
  }
}

function getArmariosFromSheet(sheetName, tipo) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  
  if (!sheet || sheet.getLastRow() < 2) {
    return [];
  }
  
  var numColumns = sheetName === 'Visitantes' ? 12 : 11;
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, numColumns).getValues();
  var armarios = [];
  
  data.forEach(function(row) {
    if (row[0]) {
      var armario = {
        id: row[0],
        numero: row[1],
        status: row[2],
        nomeVisitante: row[3] || '',
        nomePaciente: row[4] || '',
        leito: row[5] || '',
        volumes: row[6] || 0,
        horaInicio: row[7] || '',
        tipo: tipo,
        unidade: row[sheetName === 'Visitantes' ? 10 : 9] || '',
        termoAplicado: row[sheetName === 'Visitantes' ? 11 : 10] || false
      };
      
      if (sheetName === 'Visitantes') {
        armario.horaPrevista = row[8] || '';
        armario.dataRegistro = row[9] || '';
      } else {
        armario.dataRegistro = row[8] || '';
      }
      
      armarios.push(armario);
    }
  });
  
  return armarios;
}

function cadastrarArmario(armarioData) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = armarioData.tipo === 'acompanhante' ? 'Acompanhantes' : 'Visitantes';
    var sheet = ss.getSheetByName(sheetName);
    var historicoSheet = ss.getSheetByName(
      armarioData.tipo === 'acompanhante' ? 'Histórico Acompanhantes' : 'Histórico Visitantes'
    );
    
    if (!sheet || !historicoSheet) {
      return { success: false, error: 'Abas não encontradas' };
    }
    
    // Buscar dados do armário físico
    var cadastroSheet = ss.getSheetByName('Cadastro Armários');
    var cadastroData = cadastroSheet.getDataRange().getValues();
    var armarioFisico = null;
    
    for (var i = 1; i < cadastroData.length; i++) {
      if (cadastroData[i][0] == armarioData.id && cadastroData[i][2] === armarioData.tipo) {
        armarioFisico = cadastroData[i];
        break;
      }
    }
    
    if (!armarioFisico) {
      return { success: false, error: 'Armário físico não encontrado' };
    }
    
    // Verificar se o armário já está em uso
    var data = sheet.getRange(2, 1, sheet.getLastRow()-1, 3).getValues();
    var armarioExistente = data.find(row => row[1] === armarioFisico[1] && row[2] !== 'livre');
    
    if (armarioExistente) {
      return { success: false, error: 'Armário já está em uso' };
    }
    
    // Gerar novo ID
    var lastRow = sheet.getLastRow();
    var novoId = lastRow > 1 ? Math.max(...sheet.getRange(2, 1, sheet.getLastRow()-1, 1).getValues().flat()) + 1 : 1;
    
    // Preparar dados para a aba atual
    var agora = new Date();
    var novaLinha = [
      novoId,
      armarioFisico[1], // número
      'em-uso',
      armarioData.nomeVisitante,
      armarioData.nomePaciente,
      armarioData.leito,
      parseInt(armarioData.volumes),
      agora.toLocaleTimeString('pt-BR'),
      agora,
      armarioFisico[3], // unidade
      false // termoAplicado
    ];
    
    if (armarioData.tipo === 'visitante') {
      novaLinha.splice(8, 0, armarioData.horaPrevista);
    }
    
    sheet.getRange(lastRow + 1, 1, 1, novaLinha.length).setValues([novaLinha]);
    
    // Registrar no histórico
    var historicoLastRow = historicoSheet.getLastRow();
    var historicoId = historicoLastRow > 1 ? Math.max(...historicoSheet.getRange(2, 1, historicoSheet.getLastRow()-1, 1).getValues().flat()) + 1 : 1;
    
    var historicoLinha = [
      historicoId,
      new Date(),
      armarioFisico[1],
      armarioData.nomeVisitante,
      armarioData.nomePaciente,
      armarioData.leito,
      parseInt(armarioData.volumes),
      agora.toLocaleTimeString('pt-BR'),
      '', // Hora fim vazia
      'EM USO',
      armarioData.tipo,
      armarioFisico[3] // unidade
    ];
    
    historicoSheet.getRange(historicoLastRow + 1, 1, 1, historicoLinha.length).setValues([historicoLinha]);
    
    registrarLog('CADASTRO', `Armário ${armarioFisico[1]} cadastrado para ${armarioData.nomeVisitante}`);
    
    return { 
      success: true, 
      message: 'Armário cadastrado com sucesso',
      id: novoId
    };
    
  } catch (error) {
    registrarLog('ERRO', `Erro ao cadastrar armário: ${error.toString()}`);
    return { success: false, error: error.toString() };
  }
}

function liberarArmario(id, tipo) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = tipo === 'acompanhante' ? 'Acompanhantes' : 'Visitantes';
    var sheet = ss.getSheetByName(sheetName);
    var historicoSheet = ss.getSheetByName(
      tipo === 'acompanhante' ? 'Histórico Acompanhantes' : 'Histórico Visitantes'
    );
    
    if (!sheet || !historicoSheet) {
      return { success: false, error: 'Abas não encontradas' };
    }
    
    // Encontrar o armário na aba atual
    var numColumns = sheetName === 'Visitantes' ? 12 : 11;
    var data = sheet.getRange(2, 1, sheet.getLastRow()-1, numColumns).getValues();
    var armarioIndex = -1;
    var armarioData = null;
    
    data.forEach(function(row, index) {
      if (row[0] == id) {
        armarioIndex = index;
        armarioData = row;
      }
    });
    
    if (armarioIndex === -1) {
      return { success: false, error: 'Armário não encontrado' };
    }
    
    var linha = armarioIndex + 2;
    
    // Limpar dados do armário (deixar apenas número e status livre)
    var novaLinha = [
      armarioData[0], // ID
      armarioData[1], // Número
      'livre', // Status
      '', // Nome
      '', // Paciente
      '', // Leito
      '', // Volumes
      '', // Hora Início
      new Date() // Data Registro
    ];
    
    if (tipo === 'visitante') {
      novaLinha.splice(8, 0, ''); // Hora Prevista
    }
    
    novaLinha.push(armarioData[ sheetName === 'Visitantes' ? 10 : 9 ] || ''); // Unidade
    novaLinha.push(false); // TermoAplicado
    
    sheet.getRange(linha, 1, 1, novaLinha.length).setValues([novaLinha]);
    
    // Atualizar histórico - encontrar a entrada mais recente deste armário
    var historicoData = historicoSheet.getRange(2, 1, historicoSheet.getLastRow()-1, 12).getValues();
    var historicoIndex = -1;
    
    for (var i = historicoData.length - 1; i >= 0; i--) {
      if (historicoData[i][2] === armarioData[1] && historicoData[i][9] === 'EM USO') {
        historicoIndex = i;
        break;
      }
    }
    
    if (historicoIndex !== -1) {
      var historicoLinha = historicoIndex + 2;
      var agora = new Date();
      historicoSheet.getRange(historicoLinha, 9).setValue(agora.toLocaleTimeString('pt-BR')); // Hora fim
      historicoSheet.getRange(historicoLinha, 10).setValue('FINALIZADO'); // Status
    }
    
    // Remover termo se existir
    if (tipo === 'acompanhante') {
      var termosSheet = ss.getSheetByName('Termos de Responsabilidade');
      if (termosSheet) {
        var termosData = termosSheet.getDataRange().getValues();
        for (var j = 1; j < termosData.length; j++) {
          if (termosData[j][1] == id) {
            termosSheet.deleteRow(j + 1);
            break;
          }
        }
      }
    }
    
    registrarLog('LIBERAÇÃO', `Armário ${armarioData[1]} liberado`);
    
    return { success: true, message: 'Armário liberado com sucesso' };
    
  } catch (error) {
    registrarLog('ERRO', `Erro ao liberar armário: ${error.toString()}`);
    return { success: false, error: error.toString() };
  }
}

// Funções para Histórico
function getHistorico(tipo) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = tipo === 'acompanhante' ? 'Histórico Acompanhantes' : 'Histórico Visitantes';
    var sheet = ss.getSheetByName(sheetName);
    
    if (!sheet || sheet.getLastRow() < 2) {
      return { success: true, data: [] };
    }
    
    var data = sheet.getRange(2, 1, sheet.getLastRow()-1, 12).getValues();
    var historico = [];
    
    data.forEach(function(row) {
      if (row[0]) {
        historico.push({
          id: row[0],
          data: row[1],
          armario: row[2],
          nome: row[3],
          paciente: row[4],
          leito: row[5],
          volumes: row[6],
          horaInicio: row[7],
          horaFim: row[8],
          status: row[9],
          tipo: row[10],
          unidade: row[11]
        });
      }
    });
    
    return { success: true, data: historico.reverse() }; // Mais recentes primeiro
    
  } catch (error) {
    registrarLog('ERRO', `Erro ao buscar histórico: ${error.toString()}`);
    return { success: false, error: error.toString() };
  }
}

// Funções para Cadastro de Armários Físicos
function getCadastroArmarios() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Cadastro Armários');
    
    if (!sheet || sheet.getLastRow() < 2) {
      return { success: true, data: [] };
    }
    
    var data = sheet.getRange(2, 1, sheet.getLastRow()-1, 7).getValues();
    var armarios = [];
    
    data.forEach(function(row) {
      if (row[0]) {
        armarios.push({
          id: row[0],
          numero: row[1],
          tipo: row[2],
          unidade: row[3],
          localizacao: row[4],
          status: row[5],
          dataCadastro: row[6]
        });
      }
    });
    
    return { success: true, data: armarios };
    
  } catch (error) {
    registrarLog('ERRO', `Erro ao buscar cadastro de armários: ${error.toString()}`);
    return { success: false, error: error.toString() };
  }
}

function cadastrarArmarioFisico(armarioData) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Cadastro Armários');
    
    if (!sheet) {
      return { success: false, error: 'Aba de cadastro não encontrada' };
    }
    
    var totalLinhas = Math.max(sheet.getLastRow()-1, 0);
    var todosNumeros = totalLinhas > 0 ? sheet.getRange(2, 2, totalLinhas, 1).getValues().flat().filter(String) : [];

    var quantidade = parseInt(armarioData.quantidade || 1, 10);
    if (isNaN(quantidade) || quantidade < 1) {
      quantidade = 1;
    }

    var prefixo = armarioData.prefixo || '';
    var numeroInicial = parseInt(armarioData.numeroInicial || 1, 10);
    if (isNaN(numeroInicial) || numeroInicial < 1) {
      numeroInicial = 1;
    }

    var novosArmarios = [];

    if (quantidade === 1 && armarioData.numero) {
      if (todosNumeros.indexOf(armarioData.numero) !== -1) {
        return { success: false, error: 'Número de armário já existe' };
      }
      novosArmarios.push(armarioData.numero);
    } else {
      for (var i = 0; i < quantidade; i++) {
        var numeroGerado = prefixo ? prefixo + '-' + String(numeroInicial + i).padStart(3, '0') : String(numeroInicial + i);
        if (todosNumeros.indexOf(numeroGerado) !== -1 || novosArmarios.indexOf(numeroGerado) !== -1) {
          return { success: false, error: 'Não foi possível gerar numeração sem conflitos. Ajuste o prefixo ou número inicial.' };
        }
        novosArmarios.push(numeroGerado);
      }
    }

    var lastRow = sheet.getLastRow();
    var ultimoId = lastRow > 1 ? Math.max.apply(null, sheet.getRange(2, 1, sheet.getLastRow()-1, 1).getValues().flat()) : 0;

    var linhas = novosArmarios.map(function(numero, index) {
      return [
        ultimoId + index + 1,
        numero,
        armarioData.tipo,
        armarioData.unidade,
        armarioData.localizacao,
        'ativo',
        new Date()
      ];
    });

    if (linhas.length > 0) {
      sheet.getRange(lastRow + 1, 1, linhas.length, 7).setValues(linhas);
      
      // Também criar nas abas de uso
      criarArmariosUso(linhas);
    }

    registrarLog('CADASTRO', `Armários físicos cadastrados: ${novosArmarios.join(', ')}`);

    return {
      success: true,
      message: 'Armários físicos cadastrados com sucesso',
      ids: linhas.map(function(linha) { return linha[0]; }),
      numeros: novosArmarios
    };
    
  } catch (error) {
    registrarLog('ERRO', `Erro ao cadastrar armário físico: ${error.toString()}`);
    return { success: false, error: error.toString() };
  }
}

function criarArmariosUso(armarios) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    armarios.forEach(function(armario) {
      var sheetName = armario[2] === 'visitante' ? 'Visitantes' : 'Acompanhantes';
      var sheet = ss.getSheetByName(sheetName);
      
      if (sheet) {
        var lastRow = sheet.getLastRow();
        var novoId = lastRow > 1 ? Math.max(...sheet.getRange(2, 1, sheet.getLastRow()-1, 1).getValues().flat()) + 1 : 1;
        
        var novaLinha = [
          novoId,
          armario[1], // número
          'livre', // status
          '', // nome
          '', // paciente
          '', // leito
          0, // volumes
          '', // hora início
          new Date(), // data registro
          armario[3], // unidade
          false // termo aplicado
        ];
        
        if (armario[2] === 'visitante') {
          novaLinha.splice(8, 0, ''); // hora prevista
        }
        
        sheet.getRange(lastRow + 1, 1, 1, novaLinha.length).setValues([novaLinha]);
      }
    });
    
  } catch (error) {
    console.error('Erro ao criar armários de uso:', error);
  }
}

// Funções para Unidades
function getUnidades() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Unidades');
    
    if (!sheet || sheet.getLastRow() < 2) {
      return { success: true, data: [] };
    }
    
    var data = sheet.getRange(2, 1, sheet.getLastRow()-1, 4).getValues();
    var unidades = [];
    
    data.forEach(function(row) {
      if (row[0]) {
        unidades.push({
          id: row[0],
          nome: row[1],
          status: row[2],
          dataCadastro: row[3]
        });
      }
    });
    
    return { success: true, data: unidades };
    
  } catch (error) {
    registrarLog('ERRO', `Erro ao buscar unidades: ${error.toString()}`);
    return { success: false, error: error.toString() };
  }
}

function cadastrarUnidade(dados) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Unidades');
    
    if (!sheet) {
      return { success: false, error: 'Aba de unidades não encontrada' };
    }
    
    // Verificar se unidade já existe
    var data = sheet.getRange(2, 1, sheet.getLastRow()-1, 2).getValues();
    var unidadeExistente = data.find(row => row[1].toLowerCase() === dados.nome.toLowerCase());
    
    if (unidadeExistente) {
      return { success: false, error: 'Unidade já cadastrada' };
    }
    
    var lastRow = sheet.getLastRow();
    var novoId = lastRow > 1 ? Math.max(...sheet.getRange(2, 1, sheet.getLastRow()-1, 1).getValues().flat()) + 1 : 1;
    
    var novaLinha = [
      novoId,
      dados.nome,
      'ativa',
      new Date()
    ];
    
    sheet.getRange(lastRow + 1, 1, 1, 4).setValues([novaLinha]);
    
    registrarLog('CADASTRO UNIDADE', `Unidade ${dados.nome} cadastrada`);
    
    return { success: true, message: 'Unidade cadastrada com sucesso', id: novoId };
    
  } catch (error) {
    registrarLog('ERRO', `Erro ao cadastrar unidade: ${error.toString()}`);
    return { success: false, error: error.toString() };
  }
}

function alternarStatusUnidade(dados) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Unidades');
    
    if (!sheet) {
      return { success: false, error: 'Aba de unidades não encontrada' };
    }
    
    var data = sheet.getDataRange().getValues();
    var unidadeIndex = -1;
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] === dados.nome) {
        unidadeIndex = i;
        break;
      }
    }
    
    if (unidadeIndex === -1) {
      return { success: false, error: 'Unidade não encontrada' };
    }
    
    var novoStatus = data[unidadeIndex][2] === 'ativa' ? 'inativa' : 'ativa';
    sheet.getRange(unidadeIndex + 1, 3).setValue(novoStatus);
    
    registrarLog('ALTERAÇÃO UNIDADE', `Status da unidade ${dados.nome} alterado para ${novoStatus}`);
    
    return { success: true, message: `Unidade ${novoStatus === 'ativa' ? 'ativada' : 'desativada'} com sucesso` };
    
  } catch (error) {
    registrarLog('ERRO', `Erro ao alternar status da unidade: ${error.toString()}`);
    return { success: false, error: error.toString() };
  }
}

// Funções para Termos de Responsabilidade
function salvarTermoCompleto(dadosTermo) {
  try {
    // 1. Gerar e salvar PDF no Drive
    var resultadoPDF = gerarESalvarTermoPDF(dadosTermo);
    
    if (!resultadoPDF.success) {
      throw new Error('Erro ao gerar PDF: ' + resultadoPDF.error);
    }
    
    // 2. Salvar na aba "Termos de Responsabilidade"
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Termos de Responsabilidade');
    
    if (!sheet) {
      throw new Error('Aba "Termos de Responsabilidade" não encontrada');
    }
    
    var lastRow = sheet.getLastRow();
    var novoId = lastRow > 1 ? Math.max(...sheet.getRange(2, 1, sheet.getLastRow()-1, 1).getValues().flat()) + 1 : 1;
    
    // Preparar dados para a planilha
    var novaLinha = [
      novoId,
      dadosTermo.armarioId,
      dadosTermo.numeroArmario,
      dadosTermo.paciente,
      dadosTermo.prontuario,
      dadosTermo.nascimento,
      dadosTermo.setor,
      dadosTermo.leito,
      dadosTermo.consciente,
      dadosTermo.acompanhante,
      dadosTermo.telefone || '',
      dadosTermo.documento || '',
      dadosTermo.parentesco || '',
      dadosTermo.orientacoes.join(','),
      JSON.stringify(dadosTermo.volumes),
      dadosTermo.descricaoVolumes,
      new Date(),
      resultadoPDF.pdfUrl,
      dadosTermo.assinaturaBase64 || ''
    ];
    
    sheet.getRange(lastRow + 1, 1, 1, novaLinha.length).setValues([novaLinha]);
    
    // 3. Atualizar status do armário na aba "Acompanhantes"
    var sheetAcompanhantes = ss.getSheetByName('Acompanhantes');
    var dataAcompanhantes = sheetAcompanhantes.getDataRange().getValues();
    
    for (var i = 1; i < dataAcompanhantes.length; i++) {
      if (dataAcompanhantes[i][0] == dadosTermo.armarioId) {
        // Atualizar volumes e marcar termo como aplicado
        sheetAcompanhantes.getRange(i + 1, 7).setValue(dadosTermo.volumes.reduce((total, volume) => total + (Number(volume.quantidade) || 0), 0));
        sheetAcompanhantes.getRange(i + 1, 11).setValue(true); // Termo aplicado
        break;
      }
    }
    
    registrarLog('TERMO_APLICADO', `Termo aplicado para armário ${dadosTermo.numeroArmario} - PDF: ${resultadoPDF.pdfUrl}`);
    
    return {
      success: true,
      message: 'Termo salvo com sucesso e PDF gerado',
      pdfUrl: resultadoPDF.pdfUrl,
      termoId: novoId
    };
    
  } catch (error) {
    registrarLog('ERRO_TERMO', `Erro ao salvar termo: ${error.toString()}`);
    return { success: false, error: error.toString() };
  }
}

function getTermo(dados) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Termos de Responsabilidade');
    
    if (!sheet || sheet.getLastRow() < 2) {
      return { success: false, error: 'Termo não encontrado' };
    }
    
    var data = sheet.getDataRange().getValues();
    var termo = null;
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] == dados.armarioId) {
        termo = {
          id: data[i][0],
          armarioId: data[i][1],
          numeroArmario: data[i][2],
          paciente: data[i][3],
          prontuario: data[i][4],
          nascimento: data[i][5],
          setor: data[i][6],
          leito: data[i][7],
          consciente: data[i][8],
          acompanhante: data[i][9],
          telefone: data[i][10],
          documento: data[i][11],
          parentesco: data[i][12],
          orientacoes: data[i][13] ? data[i][13].split(',') : [],
          volumes: data[i][14] ? JSON.parse(data[i][14]) : [],
          descricaoVolumes: data[i][15],
          aplicadoEm: data[i][16],
          pdfUrl: data[i][17],
          assinaturaBase64: data[i][18]
        };
        break;
      }
    }
    
    if (!termo) {
      return { success: false, error: 'Termo não encontrado' };
    }
    
    return { success: true, data: termo };
    
  } catch (error) {
    registrarLog('ERRO', `Erro ao buscar termo: ${error.toString()}`);
    return { success: false, error: error.toString() };
  }
}

// Função para gerar e salvar PDF
function gerarESalvarTermoPDF(dadosTermo) {
  try {
    // Acessar a pasta do Drive
    var pastaDestino;
    try {
      pastaDestino = DriveApp.getFolderById(PASTA_DRIVE_ID);
    } catch (error) {
      // Se a pasta não for encontrada, criar na raiz
      pastaDestino = DriveApp.getRootFolder();
    }
    
    // Criar HTML do termo
    var htmlContent = criarHTMLTermo(dadosTermo);
    
    // Criar arquivo temporário como PDF
    var blob = Utilities.newBlob(htmlContent, 'text/html', 'temp.html')
      .getAs('application/pdf');
    
    // Nome do arquivo
    var nomeArquivo = 'Termo_Responsabilidade_' + dadosTermo.numeroArmario + '_' + 
                     Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'ddMMyyyy_HHmmss') + '.pdf';
    
    // Salvar na pasta
    var arquivoPDF = pastaDestino.createFile(blob).setName(nomeArquivo);
    
    // Tornar acessível via link
    arquivoPDF.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    return {
      success: true,
      pdfUrl: arquivoPDF.getUrl(),
      fileId: arquivoPDF.getId()
    };
    
  } catch (error) {
    console.error('Erro ao gerar PDF:', error);
    return { success: false, error: error.toString() };
  }
}

function criarHTMLTermo(dadosTermo) {
  var html = `
    <!DOCTYPE html>
    <html>
    <head>
        <base target="_top">
        <style>
            body { 
                font-family: Arial, sans-serif; 
                margin: 20px; 
                line-height: 1.6;
            }
            .header { 
                text-align: center; 
                border-bottom: 2px solid #333;
                padding-bottom: 10px;
                margin-bottom: 20px;
            }
            .section { 
                margin-bottom: 15px; 
                padding: 10px;
                border: 1px solid #ddd;
                border-radius: 5px;
            }
            .section-title { 
                font-weight: bold; 
                color: #2c6e8f;
                margin-bottom: 8px;
            }
            .assinatura-area {
                margin-top: 30px;
                text-align: center;
            }
            .assinatura-img {
                max-width: 300px;
                border: 1px solid #ccc;
                margin: 10px 0;
            }
            .volumes-list {
                margin-left: 20px;
            }
            .footer {
                margin-top: 40px;
                font-size: 0.9em;
                color: #666;
                text-align: center;
            }
        </style>
    </head>
    <body>
        <div class="header">
            <h2>TERMO DE RESPONSABILIDADE</h2>
            <h3>Controle de Armários Hospitalares</h3>
            <p>Armário: ${dadosTermo.numeroArmario}</p>
        </div>

        <div class="section">
            <div class="section-title">DADOS DO PACIENTE</div>
            <p><strong>Nome:</strong> ${dadosTermo.paciente}</p>
            <p><strong>Prontuário:</strong> ${dadosTermo.prontuario}</p>
            <p><strong>Data de Nascimento:</strong> ${formatarDataParaHTML(dadosTermo.nascimento)}</p>
            <p><strong>Setor/Leito:</strong> ${dadosTermo.setor} - ${dadosTermo.leito}</p>
            <p><strong>Paciente consciente/orientado:</strong> ${dadosTermo.consciente}</p>
        </div>

        <div class="section">
            <div class="section-title">RESPONSÁVEL PELO ARMÁRIO</div>
            <p><strong>Nome:</strong> ${dadosTermo.acompanhante}</p>
            <p><strong>Telefone:</strong> ${dadosTermo.telefone || 'Não informado'}</p>
            <p><strong>Documento:</strong> ${dadosTermo.documento || 'Não informado'}</p>
            <p><strong>Parentesco:</strong> ${dadosTermo.parentesco || 'Não informado'}</p>
        </div>

        <div class="section">
            <div class="section-title">VOLUMES ARMAZENADOS</div>
            <div class="volumes-list">
  `;
  
  // Adicionar volumes
  if (dadosTermo.volumes && Array.isArray(dadosTermo.volumes)) {
    dadosTermo.volumes.forEach(function(volume) {
      html += `<p>${volume.quantidade || 0}x - ${volume.descricao || ''}</p>`;
    });
  }
  
  html += `
            </div>
        </div>

        <div class="section">
            <div class="section-title">DECLARAÇÕES E ORIENTAÇÕES</div>
            <p>Declaro estar ciente e de acordo com as seguintes orientações:</p>
            <ul>
                <li>Seus pertences estão sob sua guarda e responsabilidade</li>
                <li>Em piora clínica, os pertences serão recolhidos e protocolados no NAC</li>
                <li>Após 15 dias da alta/transferência, itens não retirados poderão ser descartados conforme normas</li>
            </ul>
        </div>

        <div class="assinatura-area">
            <div class="section-title">ASSINATURA DO RESPONSÁVEL</div>
  `;
  
  // Adicionar assinatura se existir
  if (dadosTermo.assinaturaBase64) {
    html += `<img src="data:image/png;base64,${dadosTermo.assinaturaBase64}" class="assinatura-img" />`;
  } else {
    html += `<p>Assinatura digital registrada no sistema</p>`;
  }
  
  html += `
            <p><strong>Nome:</strong> ${dadosTermo.acompanhante}</p>
            <p><strong>Data/Hora:</strong> ${Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm')}</p>
        </div>

        <div class="footer">
            <p>Documento gerado automaticamente em ${Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm')}</p>
            <p>Hospital Central - Sistema de Controle de Armários</p>
        </div>
    </body>
    </html>
  `;
  
  return html;
}

function formatarDataParaHTML(data) {
  if (!data) return 'Não informada';
  try {
    var date = new Date(data);
    return Utilities.formatDate(date, 'America/Sao_Paulo', 'dd/MM/yyyy');
  } catch (error) {
    return data;
  }
}

// Funções para Movimentações
function getMovimentacoes(dados) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Movimentações');
    
    if (!sheet || sheet.getLastRow() < 2) {
      return { success: true, data: [] };
    }
    
    var data = sheet.getDataRange().getValues();
    var movimentacoes = [];
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] == dados.armarioId) {
        movimentacoes.push({
          id: data[i][0],
          armarioId: data[i][1],
          numeroArmario: data[i][2],
          tipo: data[i][3],
          descricao: data[i][4],
          responsavel: data[i][5],
          data: data[i][6],
          hora: data[i][7],
          dataHoraRegistro: data[i][8]
        });
      }
    }
    
    return { success: true, data: movimentacoes };
    
  } catch (error) {
    registrarLog('ERRO', `Erro ao buscar movimentações: ${error.toString()}`);
    return { success: false, error: error.toString() };
  }
}

function salvarMovimentacao(dados) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Movimentações');
    
    if (!sheet) {
      return { success: false, error: 'Aba de movimentações não encontrada' };
    }
    
    // Buscar número do armário
    var armarioSheet = ss.getSheetByName('Acompanhantes');
    var armarioData = armarioSheet.getDataRange().getValues();
    var numeroArmario = '';
    
    for (var i = 1; i < armarioData.length; i++) {
      if (armarioData[i][0] == dados.armarioId) {
        numeroArmario = armarioData[i][1];
        break;
      }
    }
    
    var lastRow = sheet.getLastRow();
    var novoId = lastRow > 1 ? Math.max(...sheet.getRange(2, 1, sheet.getLastRow()-1, 1).getValues().flat()) + 1 : 1;
    
    var novaLinha = [
      novoId,
      dados.armarioId,
      numeroArmario,
      dados.tipo,
      dados.descricao,
      dados.responsavel,
      dados.data,
      dados.hora,
      new Date()
    ];
    
    sheet.getRange(lastRow + 1, 1, 1, 9).setValues([novaLinha]);
    
    registrarLog('MOVIMENTAÇÃO', `Movimentação registrada para armário ${numeroArmario}`);
    
    return { success: true, message: 'Movimentação registrada com sucesso', id: novoId };
    
  } catch (error) {
    registrarLog('ERRO', `Erro ao salvar movimentação: ${error.toString()}`);
    return { success: false, error: error.toString() };
  }
}

// Funções para LOGS
function registrarLog(acao, detalhes) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('LOGS');
    
    if (!sheet) {
      return;
    }
    
    var lastRow = sheet.getLastRow();
    
    var novaLinha = [
      new Date(),
      Session.getEffectiveUser().getEmail(),
      acao,
      detalhes,
      '' // IP (não disponível no Apps Script)
    ];
    
    sheet.getRange(lastRow + 1, 1, 1, 5).setValues([novaLinha]);
    
  } catch (error) {
    // Não faz nada em caso de erro nos logs
  }
}

function getLogs() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('LOGS');
    
    if (!sheet || sheet.getLastRow() < 2) {
      return { success: true, data: [] };
    }
    
    var data = sheet.getRange(2, 1, sheet.getLastRow()-1, 5).getValues();
    var logs = [];
    
    data.forEach(function(row) {
      if (row[0]) {
        logs.push({
          dataHora: row[0],
          usuario: row[1],
          acao: row[2],
          detalhes: row[3],
          ip: row[4]
        });
      }
    });
    
    return { success: true, data: logs.reverse() }; // Mais recentes primeiro
    
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

// Funções para notificações
function getNotificacoes() {
  try {
    var agora = new Date();
    var notificacoes = [];
    
    // Verificar armários vencidos e próximos do vencimento
    var tipos = ['visitante', 'acompanhante'];
    
    tipos.forEach(function(tipo) {
      var armarios = getArmarios(tipo);
      if (armarios.success) {
        armarios.data.forEach(function(armario) {
          if (armario.status === 'em-uso' && armario.horaPrevista) {
            try {
              // Converter hora prevista para objeto Date
              var hoje = new Date().toISOString().split('T')[0];
              var horaPrevista = new Date(hoje + 'T' + armario.horaPrevista + ':00');
              var diferencaMinutos = (horaPrevista - agora) / (1000 * 60);
              
              if (diferencaMinutos < 0) {
                // Vencido
                notificacoes.push({
                  tipo: 'danger',
                  titulo: `Armário ${armario.numero} vencido`,
                  tempo: `Há ${Math.abs(Math.round(diferencaMinutos))} minutos`
                });
              } else if (diferencaMinutos <= 10) {
                // Próximo do vencimento (10 minutos ou menos)
                notificacoes.push({
                  tipo: 'warning', 
                  titulo: `Armário ${armario.numero} próximo do horário`,
                  tempo: `Há ${Math.round(diferencaMinutos)} minutos`
                });
              }
            } catch (e) {
              // Ignora erro de parsing de data
            }
          }
        });
      }
    });
    
    return { success: true, data: notificacoes };
    
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

// Função para obter estatísticas do dashboard
function getEstatisticasDashboard(tipoUsuario) {
  try {
    var estatisticas = {
      livres: 0,
      emUso: 0,
      proximo: 0,
      vencidos: 0
    };
    
    var tipos = [];
    
    // Definir quais tipos de armário o usuário pode ver
    if (tipoUsuario === 'admin' || tipoUsuario === 'ambos') {
      tipos = ['visitante', 'acompanhante'];
    } else if (tipoUsuario === 'visitante') {
      tipos = ['visitante'];
    } else if (tipoUsuario === 'acompanhante') {
      tipos = ['acompanhante'];
    }
    
    var agora = new Date();
    
    tipos.forEach(function(tipo) {
      var armarios = getArmarios(tipo);
      if (armarios.success) {
        armarios.data.forEach(function(armario) {
          if (armario.status === 'livre') {
            estatisticas.livres++;
          } else if (armario.status === 'em-uso') {
            if (armario.horaPrevista) {
              try {
                var hoje = new Date().toISOString().split('T')[0];
                var horaPrevista = new Date(hoje + 'T' + armario.horaPrevista + ':00');
                var diferencaMinutos = (horaPrevista - agora) / (1000 * 60);
                
                if (diferencaMinutos < 0) {
                  estatisticas.vencidos++;
                } else if (diferencaMinutos <= 10) {
                  estatisticas.proximo++;
                } else {
                  estatisticas.emUso++;
                }
              } catch (e) {
                estatisticas.emUso++;
              }
            } else {
              estatisticas.emUso++;
            }
          }
        });
      }
    });
    
    return { success: true, data: estatisticas };
    
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

// Função para verificar se o sistema está inicializado
function verificarInicializacao() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var abas = [
    'Histórico Visitantes', 
    'Histórico Acompanhantes', 
    'Visitantes', 
    'Acompanhantes', 
    'Cadastro Armários', 
    'Unidades', 
    'Usuários', 
    'LOGS',
    'Termos de Responsabilidade',
    'Movimentações'
  ];
  
  for (var i = 0; i < abas.length; i++) {
    if (!ss.getSheetByName(abas[i])) {
      return { success: true, inicializado: false };
    }
  }
  
  return { success: true, inicializado: true };
}
