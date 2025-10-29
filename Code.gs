// Configuração inicial
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setTitle('LockNAC - Gerenciamento de Armários');
}

function doPost(e) {
  return handlePost(e);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function normalizarTextoBasico(valor) {
  if (valor === null || valor === undefined) {
    return '';
  }
  var texto = valor.toString().trim().toLowerCase();
  if (typeof texto.normalize === 'function') {
    texto = texto.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  }
  return texto;
}

function obterEstruturaPlanilha(sheet) {
  var ultimaColuna = sheet.getLastColumn();
  var cabecalhos = ultimaColuna > 0 ? sheet.getRange(1, 1, 1, ultimaColuna).getValues()[0] : [];
  var mapaIndices = {};

  cabecalhos.forEach(function(cabecalho, indice) {
    var chave = normalizarTextoBasico(cabecalho);
    if (chave && mapaIndices[chave] === undefined) {
      mapaIndices[chave] = indice;
    }
  });

  return {
    ultimaColuna: ultimaColuna,
    mapaIndices: mapaIndices
  };
}

function obterIndiceColuna(estrutura, chave, padrao) {
  if (estrutura.mapaIndices.hasOwnProperty(chave)) {
    return estrutura.mapaIndices[chave];
  }
  return padrao;
}

function obterValorLinha(linha, estrutura, chave, padrao) {
  var indice = obterIndiceColuna(estrutura, chave, null);
  if (indice === null || indice === undefined) {
    return padrao;
  }
  if (indice >= linha.length) {
    return padrao;
  }
  return linha[indice];
}

function definirValorLinha(linha, estrutura, chave, valor) {
  var indice = obterIndiceColuna(estrutura, chave, null);
  if (indice === null || indice === undefined) {
    return;
  }
  while (linha.length < estrutura.ultimaColuna) {
    linha.push('');
  }
  linha[indice] = valor;
}

function obterValorLinhaFlexivel(linha, estrutura, chaves, padrao) {
  if (!Array.isArray(chaves)) {
    chaves = [chaves];
  }

  for (var i = 0; i < chaves.length; i++) {
    var indice = obterIndiceColuna(estrutura, chaves[i], null);
    if (indice !== null && indice !== undefined && indice < linha.length) {
      return linha[indice];
    }
  }

  return padrao;
}

function definirValorLinhaFlexivel(linha, estrutura, chaves, valor) {
  if (!Array.isArray(chaves)) {
    chaves = [chaves];
  }

  for (var i = 0; i < chaves.length; i++) {
    var indice = obterIndiceColuna(estrutura, chaves[i], null);
    if (indice === null || indice === undefined) {
      continue;
    }

    var tamanhoMinimo = Math.max(estrutura.ultimaColuna, indice + 1);
    while (linha.length < tamanhoMinimo) {
      linha.push('');
    }

    linha[indice] = valor;
    return true;
  }

  return false;
}

function converterParaBoolean(valor) {
  if (valor === true || valor === false) {
    return valor;
  }
  if (typeof valor === 'number') {
    return valor !== 0;
  }
  if (typeof valor === 'string') {
    var texto = valor.trim().toLowerCase();
    return texto === 'true' || texto === '1' || texto === 'sim';
  }
  return false;
}

function normalizarListaUnidadesParametro(valor) {
  try {
    if (valor === null || valor === undefined) {
      return [];
    }

    if (Array.isArray(valor)) {
      return valor.map(function(item) {
        return item !== null && item !== undefined ? item.toString().trim() : '';
      }).filter(function(item) {
        return item;
      });
    }

    if (typeof valor === 'string') {
      var texto = valor.trim();
      if (!texto) {
        return [];
      }

      if ((texto.charAt(0) === '[' && texto.charAt(texto.length - 1) === ']') ||
          (texto.charAt(0) === '{' && texto.charAt(texto.length - 1) === '}')) {
        try {
          var convertido = JSON.parse(texto);
          if (Array.isArray(convertido)) {
            return normalizarListaUnidadesParametro(convertido);
          }
        } catch (erroJSON) {
          console.error('Falha ao interpretar unidades como JSON:', erroJSON);
        }
      }

      if (texto.indexOf(';') !== -1 || texto.indexOf(',') !== -1) {
        return texto.split(/[;,]/).map(function(item) {
          return item.trim();
        }).filter(function(item) {
          return item;
        });
      }

      return [texto];
    }

    if (typeof valor === 'number' || typeof valor === 'boolean') {
      return [valor.toString()];
    }

    if (typeof valor === 'object') {
      var itens = [];
      for (var chave in valor) {
        if (!valor.hasOwnProperty(chave)) {
          continue;
        }
        var item = valor[chave];
        if (Array.isArray(item)) {
          itens = itens.concat(normalizarListaUnidadesParametro(item));
        } else if (item !== null && item !== undefined) {
          itens.push(item.toString().trim());
        }
      }
      return itens.filter(function(item) {
        return item;
      });
    }
  } catch (erro) {
    console.error('Erro ao normalizar unidades informadas:', erro);
  }

  return [];
}

function obterMapasUnidades() {
  var mapas = {
    porId: {},
    porNome: {}
  };

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Unidades');

    if (!sheet) {
      return mapas;
    }

    var ultimaLinha = sheet.getLastRow();
    if (ultimaLinha < 2) {
      return mapas;
    }

    var dados = sheet.getRange(2, 1, ultimaLinha - 1, 2).getValues();
    dados.forEach(function(row) {
      var id = row[0];
      var nome = row[1];
      if (id === null || id === undefined) {
        return;
      }
      var idTexto = id.toString().trim();
      if (!idTexto) {
        return;
      }
      var nomeTexto = nome !== null && nome !== undefined ? nome.toString().trim() : '';
      mapas.porId[idTexto] = nomeTexto;
      if (nomeTexto) {
        var chaveNome = normalizarTextoBasico(nomeTexto);
        if (!mapas.porNome[chaveNome]) {
          mapas.porNome[chaveNome] = idTexto;
        }
      }
    });
  } catch (erroMapas) {
    console.error('Erro ao obter mapas de unidades:', erroMapas);
  }

  return mapas;
}

function formatarUnidadesParaRegistro(unidadesIds, mapas) {
  if (!Array.isArray(unidadesIds)) {
    return [];
  }

  var resultado = [];
  unidadesIds.forEach(function(unidadeId) {
    if (unidadeId === null || unidadeId === undefined) {
      return;
    }

    var chave = unidadeId.toString().trim();
    if (!chave) {
      return;
    }

    if (normalizarTextoBasico(chave) === 'all') {
      if (resultado.indexOf('Todas as unidades') === -1) {
        resultado.push('Todas as unidades');
      }
      return;
    }

    var nome = mapas && mapas.porId ? mapas.porId[chave] : '';
    if (nome) {
      if (resultado.indexOf(nome) === -1) {
        resultado.push(nome);
      }
    } else {
      if (resultado.indexOf(chave) === -1) {
        resultado.push(chave);
      }
    }
  });

  return resultado;
}

function resolverIdsUnidadesArmazenadas(unidadesValor, mapas) {
  var brutas = normalizarListaUnidadesParametro(unidadesValor);
  if (!brutas.length) {
    return [];
  }

  var ids = [];
  brutas.forEach(function(item) {
    if (item === null || item === undefined) {
      return;
    }

    var textoOriginal = item.toString().trim();
    if (!textoOriginal) {
      return;
    }

    var textoNormalizado = normalizarTextoBasico(textoOriginal);
    if (textoNormalizado === 'all' || textoNormalizado === 'todas as unidades') {
      if (ids.indexOf('all') === -1) {
        ids.push('all');
      }
      return;
    }

    if (mapas && mapas.porId && mapas.porId.hasOwnProperty(textoOriginal)) {
      if (ids.indexOf(textoOriginal) === -1) {
        ids.push(textoOriginal);
      }
      return;
    }

    var separadores = [' - ', '|', ':', '–', ' — '];
    for (var i = 0; i < separadores.length; i++) {
      var sep = separadores[i];
      if (textoOriginal.indexOf(sep) !== -1) {
        var candidato = textoOriginal.split(sep)[0].trim();
        if (candidato) {
          if (mapas && mapas.porId && mapas.porId.hasOwnProperty(candidato)) {
            if (ids.indexOf(candidato) === -1) {
              ids.push(candidato);
            }
            return;
          }
          if (mapas && mapas.porNome) {
            var candidatoNormalizado = normalizarTextoBasico(candidato);
            var idPorNome = mapas.porNome[candidatoNormalizado];
            if (idPorNome && ids.indexOf(idPorNome) === -1) {
              ids.push(idPorNome);
              return;
            }
          }
        }
      }
    }

    if (mapas && mapas.porNome) {
      var idPorNomeDireto = mapas.porNome[textoNormalizado];
      if (idPorNomeDireto && ids.indexOf(idPorNomeDireto) === -1) {
        ids.push(idPorNomeDireto);
        return;
      }
    }

    var matchNumero = textoOriginal.match(/\d+/);
    if (matchNumero && mapas && mapas.porId && mapas.porId.hasOwnProperty(matchNumero[0])) {
      var idNumero = matchNumero[0];
      if (ids.indexOf(idNumero) === -1) {
        ids.push(idNumero);
      }
      return;
    }

    if (ids.indexOf(textoOriginal) === -1) {
      ids.push(textoOriginal);
    }
  });

  return ids;
}

// ID da pasta do Drive para salvar os PDFs - ATUALIZE COM SEU ID
const PASTA_DRIVE_ID = '1nYsGJJUIufxDYVvIanVXCbPx7YuBOYDP';

// Configuração de cache para leitura dos termos
const TERMOS_CACHE_KEY = 'termos_registrados_cache_v1';
const TERMOS_CACHE_TTL = 120; // segundos

// Inicializar planilha com todas as abas e cabeçalhos
function inicializarPlanilha() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Criar abas se não existirem
    var abas = [
      {
        nome: 'Histórico Visitantes',
        cabecalhos: ['ID', 'Data', 'Número Armário', 'Nome Visitante', 'Nome Paciente', 'Leito', 'Volumes', 'Hora Início', 'Hora Fim', 'Status', 'Tipo', 'Unidade', 'WhatsApp']
      },
      {
        nome: 'Histórico Acompanhantes',
        cabecalhos: ['ID', 'Data', 'Número Armário', 'Nome Acompanhante', 'Nome Paciente', 'Leito', 'Volumes', 'Hora Início', 'Hora Fim', 'Status', 'Tipo', 'Unidade', 'WhatsApp']
      },
      {
        nome: 'Visitantes',
        cabecalhos: ['ID', 'Número', 'Status', 'Nome Visitante', 'Nome Paciente', 'Leito', 'Volumes', 'Hora Início', 'Hora Prevista', 'Data Registro', 'Unidade', 'Termo Aplicado', 'WhatsApp']
      },
      {
        nome: 'Acompanhantes',
        cabecalhos: ['ID', 'Número', 'Status', 'Nome Acompanhante', 'Nome Paciente', 'Leito', 'Volumes', 'Hora Início', 'Data Registro', 'WhatsApp', 'Unidade', 'Termo Aplicado']
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
        cabecalhos: ['ID', 'Nome', 'Email', 'Perfil', 'Acesso Visitantes', 'Acesso Acompanhantes', 'Data Cadastro', 'Status', 'Senha', 'Unidades']
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
    usuariosSheet.getRange(2, 1, 1, 10)
      .setValues([[1, 'Administrador', 'admin', 'admin', true, true, new Date(), 'ativo', 'admin123', 'all']]);
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

      case 'atualizarUsuario':
        return ContentService.createTextOutput(JSON.stringify(atualizarUsuario(e.parameter)))
          .setMimeType(ContentService.MimeType.JSON);

      case 'excluirUsuario':
        return ContentService.createTextOutput(JSON.stringify(excluirUsuario(e.parameter)))
          .setMimeType(ContentService.MimeType.JSON);

      case 'autenticarUsuario':
        return ContentService.createTextOutput(JSON.stringify(autenticarUsuario(e.parameter)))
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

      case 'finalizarTermo':
        return ContentService.createTextOutput(JSON.stringify(finalizarTermo(e.parameter)))
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
    var tipoNormalizado = normalizarTextoBasico(tipo);
    if (!tipoNormalizado) {
      tipoNormalizado = 'geral';
    }

    var incluirTermos = tipoNormalizado === 'acompanhante' || tipoNormalizado === 'admin' ||
      tipoNormalizado === 'ambos' || tipoNormalizado === 'todos' || tipoNormalizado === 'geral';
    var termosMap = {};

    if (incluirTermos) {
      var termosInfo = obterTermosRegistrados();
      termosInfo.termos.forEach(function(termo) {
        var chave = termo.armarioId;
        if (!chave && chave !== 0) {
          return;
        }

        var termoAtual = termosMap[chave];
        var termoFinalizado = Boolean(termo.pdfUrl || (termo.assinaturas && termo.assinaturas.finalizadoEm));
        if (!termoAtual) {
          termosMap[chave] = termo;
          return;
        }

        var atualFinalizado = Boolean(termoAtual.pdfUrl || (termoAtual.assinaturas && termoAtual.assinaturas.finalizadoEm));

        if (!termoFinalizado && atualFinalizado) {
          termosMap[chave] = termo;
        } else if (termoFinalizado === atualFinalizado && termoAtual.id < termo.id) {
          termosMap[chave] = termo;
        }
      });
    }

    if (tipoNormalizado === 'admin' || tipoNormalizado === 'ambos' || tipoNormalizado === 'todos' || tipoNormalizado === 'geral') {
      var visitantes = getArmariosFromSheet('Visitantes', 'visitante', null);
      var acompanhantes = getArmariosFromSheet('Acompanhantes', 'acompanhante', termosMap);
      return { success: true, data: visitantes.concat(acompanhantes) };
    } else {
      var sheetName = tipoNormalizado === 'acompanhante' ? 'Acompanhantes' : 'Visitantes';
      var mapa = tipoNormalizado === 'acompanhante' ? termosMap : null;
      return { success: true, data: getArmariosFromSheet(sheetName, tipoNormalizado, mapa) };
    }
  } catch (error) {
    registrarLog('ERRO', `Erro ao buscar armários: ${error.toString()}`);
    return { success: false, error: error.toString() };
  }
}

function getArmariosFromSheet(sheetName, tipo, termosMap) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);

  if (!sheet || sheet.getLastRow() < 2) {
    return [];
  }

  var isVisitante = sheetName === 'Visitantes';
  var estrutura = obterEstruturaPlanilha(sheet);
  var totalLinhas = sheet.getLastRow() - 1;
  var totalColunas = estrutura.ultimaColuna || (isVisitante ? 13 : 12);
  var dados = sheet.getRange(2, 1, totalLinhas, totalColunas).getValues();
  var armarios = [];

  var idIndex = obterIndiceColuna(estrutura, 'id', 0);
  var numeroIndex = obterIndiceColuna(estrutura, 'numero', 1);
  var statusIndex = obterIndiceColuna(estrutura, 'status', 2);
  var nomeIndex = obterIndiceColuna(estrutura, isVisitante ? 'nome visitante' : 'nome acompanhante', 3);
  var pacienteIndex = obterIndiceColuna(estrutura, 'nome paciente', 4);
  var leitoIndex = obterIndiceColuna(estrutura, 'leito', 5);
  var volumesIndex = obterIndiceColuna(estrutura, 'volumes', 6);
  var horaInicioIndex = obterIndiceColuna(estrutura, 'hora inicio', 7);
  var horaPrevistaIndex = isVisitante ? obterIndiceColuna(estrutura, 'hora prevista', 8) : -1;
  var dataRegistroIndex = obterIndiceColuna(estrutura, 'data registro', isVisitante ? 9 : 8);
  var unidadeIndex = obterIndiceColuna(estrutura, 'unidade', null);
  if (unidadeIndex === null || unidadeIndex === undefined) {
    unidadeIndex = isVisitante ? 10 : 10;
  }
  var termoIndex = obterIndiceColuna(estrutura, 'termo aplicado', null);
  if (termoIndex === null || termoIndex === undefined) {
    termoIndex = isVisitante ? 11 : 11;
  }
  var whatsappIndex = obterIndiceColuna(estrutura, 'whatsapp', null);
  if (whatsappIndex === null || whatsappIndex === undefined) {
    whatsappIndex = isVisitante ? 12 : 9;
  }

  dados.forEach(function(row) {
    var id = row[idIndex];
    if (!id && id !== 0) {
      return;
    }

    var statusValor = row[statusIndex];
    var statusNormalizado = normalizarTextoBasico(statusValor);
    var status;
    switch (statusNormalizado) {
      case 'em-uso':
      case 'em uso':
        status = 'em-uso';
        break;
      case 'proximo':
        status = 'proximo';
        break;
      case 'vencido':
        status = 'vencido';
        break;
      case 'livre':
        status = 'livre';
        break;
      default:
        status = statusNormalizado || 'livre';
        break;
    }

    var armario = {
      id: id,
      numero: row[numeroIndex] || '',
      status: status,
      nomeVisitante: row[nomeIndex] || '',
      nomePaciente: row[pacienteIndex] || '',
      leito: row[leitoIndex] || '',
      volumes: row[volumesIndex] || 0,
      horaInicio: row[horaInicioIndex] || '',
      tipo: tipo,
      unidade: unidadeIndex !== null && unidadeIndex !== undefined ? (row[unidadeIndex] || '') : '',
      termoAplicado: termoIndex !== null && termoIndex !== undefined ? converterParaBoolean(row[termoIndex]) : false,
      whatsapp: whatsappIndex !== null && whatsappIndex !== undefined ? (row[whatsappIndex] || '') : ''
    };

    if (isVisitante) {
      armario.horaPrevista = horaPrevistaIndex > -1 ? (row[horaPrevistaIndex] || '') : '';
      armario.dataRegistro = dataRegistroIndex > -1 ? (row[dataRegistroIndex] || '') : '';
    } else {
      armario.dataRegistro = dataRegistroIndex > -1 ? (row[dataRegistroIndex] || '') : '';
    }

    var volumesNumero = parseInt(armario.volumes, 10);
    armario.volumes = isNaN(volumesNumero) ? 0 : volumesNumero;

    if (tipo === 'acompanhante') {
      var termoRelacionado = termosMap ? termosMap[armario.id] : null;
      if (termoRelacionado) {
        var termoFinalizado = Boolean(termoRelacionado.pdfUrl || (termoRelacionado.assinaturas && termoRelacionado.assinaturas.finalizadoEm));
        armario.termoAplicado = !termoFinalizado;
        armario.termoFinalizado = termoFinalizado;
        armario.termoInfo = {
          id: termoRelacionado.id,
          aplicadoEm: termoRelacionado.aplicadoEm,
          finalizadoEm: termoRelacionado.assinaturas ? termoRelacionado.assinaturas.finalizadoEm : '',
          pdfUrl: termoRelacionado.pdfUrl || '',
          responsavel: termoRelacionado.acompanhante,
          metodoFinal: termoRelacionado.assinaturas ? termoRelacionado.assinaturas.metodoFinal : '',
          cpfFinal: termoRelacionado.assinaturas ? termoRelacionado.assinaturas.cpfFinal : ''
        };
      } else {
        armario.termoAplicado = false;
        armario.termoFinalizado = false;
        armario.termoInfo = null;
      }
    } else {
      armario.termoFinalizado = false;
      armario.termoInfo = null;
    }

    armarios.push(armario);
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

    var totalLinhas = sheet.getLastRow();
    if (totalLinhas < 2) {
      return { success: false, error: 'Nenhum armário cadastrado' };
    }

    var estrutura = obterEstruturaPlanilha(sheet);
    var totalColunas = estrutura.ultimaColuna || (sheetName === 'Visitantes' ? 13 : 12);
    var data = sheet.getRange(2, 1, totalLinhas - 1, totalColunas).getValues();
    var linhaPlanilha = -1;
    var linhaAtual = null;
    var idNumerico = Number(armarioData.id);
    var idIndex = obterIndiceColuna(estrutura, 'id', 0);
    var numeroIndex = obterIndiceColuna(estrutura, 'numero', 1);
    var statusIndex = obterIndiceColuna(estrutura, 'status', 2);

    for (var i = 0; i < data.length; i++) {
      if (data[i][idIndex] == idNumerico) {
        linhaPlanilha = i + 2;
        linhaAtual = data[i].slice();
        break;
      }
    }

    if ((linhaPlanilha === -1 || !linhaAtual) && armarioData.numero) {
      var numeroInformado = String(armarioData.numero);
      for (var j = 0; j < data.length; j++) {
        var statusLinha = normalizarTextoBasico(data[j][statusIndex]);
        if (data[j][numeroIndex] === numeroInformado && statusLinha === 'livre') {
          linhaPlanilha = j + 2;
          linhaAtual = data[j].slice();
          break;
        }
      }
    }

    if (linhaPlanilha === -1 || !linhaAtual) {
      return { success: false, error: 'Armário não encontrado' };
    }

    var statusAtual = normalizarTextoBasico(linhaAtual[statusIndex]);
    if (statusAtual !== 'livre') {
      return { success: false, error: 'Armário já está em uso' };
    }

    var agora = new Date();
    var horaInicio = agora.toLocaleTimeString('pt-BR');
    var volumes = parseInt(armarioData.volumes, 10);
    if (isNaN(volumes) || volumes < 0) {
      volumes = 0;
    }
    var whatsapp = armarioData.whatsapp || '';
    var numeroArmario = linhaAtual[numeroIndex];
    var unidadeAtual = obterValorLinha(linhaAtual, estrutura, 'unidade', '');
    var novaLinha = linhaAtual.slice();
    while (novaLinha.length < totalColunas) {
      novaLinha.push('');
    }
    var nomeColuna = sheetName === 'Visitantes' ? 'nome visitante' : 'nome acompanhante';

    definirValorLinha(novaLinha, estrutura, 'status', 'em-uso');
    definirValorLinha(novaLinha, estrutura, nomeColuna, armarioData.nomeVisitante);
    definirValorLinha(novaLinha, estrutura, 'nome paciente', armarioData.nomePaciente);
    definirValorLinha(novaLinha, estrutura, 'leito', armarioData.leito);
    definirValorLinha(novaLinha, estrutura, 'volumes', volumes);
    definirValorLinha(novaLinha, estrutura, 'hora inicio', horaInicio);
    if (sheetName === 'Visitantes') {
      definirValorLinha(novaLinha, estrutura, 'hora prevista', armarioData.horaPrevista || '');
    } else {
      definirValorLinha(novaLinha, estrutura, 'hora prevista', '');
    }
    definirValorLinha(novaLinha, estrutura, 'data registro', agora);
    definirValorLinha(novaLinha, estrutura, 'unidade', unidadeAtual);
    definirValorLinha(novaLinha, estrutura, 'whatsapp', whatsapp);
    definirValorLinha(novaLinha, estrutura, 'termo aplicado', false);

    sheet.getRange(linhaPlanilha, 1, 1, totalColunas).setValues([novaLinha]);

    var historicoLastRow = historicoSheet.getLastRow();
    var historicoId = historicoLastRow > 1 ? Math.max(...historicoSheet.getRange(2, 1, historicoLastRow - 1, 1).getValues().flat()) + 1 : 1;

    var historicoLinha = [
      historicoId,
      new Date(),
      numeroArmario,
      armarioData.nomeVisitante,
      armarioData.nomePaciente,
      armarioData.leito,
      volumes,
      horaInicio,
      '',
      'EM USO',
      armarioData.tipo,
      unidadeAtual,
      whatsapp
    ];

    historicoSheet.getRange(historicoLastRow + 1, 1, 1, historicoLinha.length).setValues([historicoLinha]);

    registrarLog('CADASTRO', `Armário ${numeroArmario} cadastrado para ${armarioData.nomeVisitante}`);

    return {
      success: true,
      message: 'Armário cadastrado com sucesso',
      id: linhaAtual[idIndex]
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
    var estrutura = obterEstruturaPlanilha(sheet);
    var totalColunas = estrutura.ultimaColuna || (sheetName === 'Visitantes' ? 13 : 12);
    var data = sheet.getRange(2, 1, sheet.getLastRow()-1, totalColunas).getValues();
    var armarioIndex = -1;
    var armarioData = null;
    var idIndex = obterIndiceColuna(estrutura, 'id', 0);

    data.forEach(function(row, index) {
      if (row[idIndex] == id) {
        armarioIndex = index;
        armarioData = row.slice();
      }
    });
    
    if (armarioIndex === -1) {
      return { success: false, error: 'Armário não encontrado' };
    }
    
    var linha = armarioIndex + 2;
    
    // Limpar dados do armário (deixar apenas número e status livre)
    var unidadeAtual = obterValorLinha(armarioData, estrutura, 'unidade', '');
    var novaLinha = armarioData.slice();
    while (novaLinha.length < totalColunas) {
      novaLinha.push('');
    }

    var nomeColuna = sheetName === 'Visitantes' ? 'nome visitante' : 'nome acompanhante';
    definirValorLinha(novaLinha, estrutura, 'status', 'livre');
    definirValorLinha(novaLinha, estrutura, nomeColuna, '');
    definirValorLinha(novaLinha, estrutura, 'nome paciente', '');
    definirValorLinha(novaLinha, estrutura, 'leito', '');
    definirValorLinha(novaLinha, estrutura, 'volumes', '');
    definirValorLinha(novaLinha, estrutura, 'hora inicio', '');
    if (sheetName === 'Visitantes') {
      definirValorLinha(novaLinha, estrutura, 'hora prevista', '');
    }
    definirValorLinha(novaLinha, estrutura, 'data registro', new Date());
    definirValorLinha(novaLinha, estrutura, 'whatsapp', '');
    definirValorLinha(novaLinha, estrutura, 'unidade', unidadeAtual);
    definirValorLinha(novaLinha, estrutura, 'termo aplicado', false);

    sheet.getRange(linha, 1, 1, totalColunas).setValues([novaLinha]);

    // Atualizar histórico - encontrar a entrada mais recente deste armário
    var historicoData = historicoSheet.getRange(2, 1, historicoSheet.getLastRow()-1, 13).getValues();
    var historicoIndex = -1;
    var numeroIndex = obterIndiceColuna(estrutura, 'numero', 1);
    var numeroArmario = obterValorLinha(armarioData, estrutura, 'numero', armarioData[numeroIndex]);

    for (var i = historicoData.length - 1; i >= 0; i--) {
      if (historicoData[i][2] === numeroArmario && historicoData[i][9] === 'EM USO') {
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
    
    registrarLog('LIBERAÇÃO', `Armário ${numeroArmario} liberado`);

    return { success: true, message: 'Armário liberado com sucesso' };
    
  } catch (error) {
    registrarLog('ERRO', `Erro ao liberar armário: ${error.toString()}`);
    return { success: false, error: error.toString() };
  }
}

// Funções para Usuários
function getUsuarios() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Usuários');

    if (!sheet || sheet.getLastRow() < 2) {
      return { success: true, data: [] };
    }

    var estrutura = obterEstruturaPlanilha(sheet);
    var totalColunas = estrutura.ultimaColuna || 10;
    var mapasUnidades = obterMapasUnidades();
    var dados = sheet.getRange(2, 1, sheet.getLastRow() - 1, totalColunas).getValues();
    var usuarios = [];

    dados.forEach(function(linha) {
      var id = obterValorLinha(linha, estrutura, 'id', linha[0]);
      if (!id && id !== 0) {
        return;
      }

      var perfilValor = obterValorLinha(linha, estrutura, 'perfil', 'usuario');
      var perfil = perfilValor ? perfilValor.toString().trim().toLowerCase() : 'usuario';
      var unidadesBrutas = obterValorLinhaFlexivel(linha, estrutura, ['unidades', 'unidade', 'acesso unidades'], '');
      var unidades = resolverIdsUnidadesArmazenadas(unidadesBrutas, mapasUnidades);
      var unidadesUnicas = [];
      unidades.forEach(function(unidade) {
        if (unidadesUnicas.indexOf(unidade) === -1) {
          unidadesUnicas.push(unidade);
        }
      });

      usuarios.push({
        id: id,
        nome: obterValorLinha(linha, estrutura, 'nome', ''),
        email: obterValorLinha(linha, estrutura, 'email', ''),
        perfil: perfil,
        acessoVisitantes: converterParaBoolean(obterValorLinha(linha, estrutura, 'acesso visitantes', false)),
        acessoAcompanhantes: converterParaBoolean(obterValorLinha(linha, estrutura, 'acesso acompanhantes', false)),
        dataCadastro: obterValorLinha(linha, estrutura, 'data cadastro', ''),
        status: obterValorLinha(linha, estrutura, 'status', ''),
        senha: obterValorLinha(linha, estrutura, 'senha', ''),
        unidades: unidadesUnicas
      });
    });

    return { success: true, data: usuarios };

  } catch (error) {
    registrarLog('ERRO', `Erro ao buscar usuários: ${error.toString()}`);
    return { success: false, error: error.toString() };
  }
}

function cadastrarUsuario(dados) {
  try {
    var nome = (dados.nome || '').toString().trim();
    var email = (dados.email || '').toString().trim();
    var perfil = (dados.perfil || '').toString().trim().toLowerCase();
    var senha = (dados.senha || '').toString().trim();
    var unidadesLista = normalizarListaUnidadesParametro(dados.unidades);
    var unidadesUnicas = [];
    var incluiTodas = false;

    unidadesLista.forEach(function(unidade) {
      var chave = unidade.toString().trim();
      if (!chave) {
        return;
      }
      if (normalizarTextoBasico(chave) === 'all') {
        incluiTodas = true;
        return;
      }
      if (unidadesUnicas.indexOf(chave) === -1) {
        unidadesUnicas.push(chave);
      }
    });

    if (incluiTodas || (perfil === 'admin' && unidadesUnicas.length === 0)) {
      unidadesUnicas = ['all'];
    }

    if (!nome || !email || !perfil) {
      return { success: false, error: 'Nome, matrícula e perfil são obrigatórios' };
    }

    if (!senha) {
      return { success: false, error: 'Informe uma senha para o usuário' };
    }

    if (unidadesUnicas.length === 0 && perfil !== 'admin') {
      return { success: false, error: 'Informe ao menos uma unidade de acesso' };
    }

    var mapasUnidades = obterMapasUnidades();
    var unidadesFormatadas = formatarUnidadesParaRegistro(unidadesUnicas, mapasUnidades);
    var unidadesTexto = unidadesFormatadas.join('; ');

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Usuários');

    if (!sheet) {
      return { success: false, error: 'Aba de usuários não encontrada' };
    }

    var estrutura = obterEstruturaPlanilha(sheet);
    var totalColunas = estrutura.ultimaColuna || 10;
    var ultimaLinha = sheet.getLastRow();
    var idIndex = obterIndiceColuna(estrutura, 'id', 0);
    var proximoId = 1;

    if (ultimaLinha >= 2) {
      var idRange = sheet.getRange(2, idIndex + 1, ultimaLinha - 1, 1).getValues().flat();
      var idsNumericos = idRange.map(function(valor) {
        var numero = parseInt(valor, 10);
        return isNaN(numero) ? null : numero;
      }).filter(function(valor) {
        return valor !== null;
      });
      var ultimoId = idsNumericos.length > 0 ? Math.max.apply(null, idsNumericos) : 0;
      proximoId = ultimoId + 1;
    }

    var acessoVisitantes = converterParaBoolean(dados.acessoVisitantes);
    var acessoAcompanhantes = converterParaBoolean(dados.acessoAcompanhantes);
    var dataCadastro = new Date();

    var novaLinha = new Array(totalColunas);
    for (var i = 0; i < totalColunas; i++) {
      novaLinha[i] = '';
    }

    definirValorLinha(novaLinha, estrutura, 'id', proximoId);
    definirValorLinha(novaLinha, estrutura, 'nome', nome);
    definirValorLinha(novaLinha, estrutura, 'email', email);
    definirValorLinha(novaLinha, estrutura, 'perfil', perfil);
    definirValorLinha(novaLinha, estrutura, 'acesso visitantes', acessoVisitantes);
    definirValorLinha(novaLinha, estrutura, 'acesso acompanhantes', acessoAcompanhantes);
    definirValorLinha(novaLinha, estrutura, 'data cadastro', dataCadastro);
    definirValorLinha(novaLinha, estrutura, 'status', 'ativo');
    definirValorLinha(novaLinha, estrutura, 'senha', senha);
    if (!definirValorLinhaFlexivel(novaLinha, estrutura, ['unidades', 'unidade', 'acesso unidades'], unidadesTexto)) {
      definirValorLinha(novaLinha, estrutura, 'unidades', unidadesTexto);
    }

    sheet.getRange(ultimaLinha + 1, 1, 1, totalColunas).setValues([novaLinha]);

    registrarLog('CADASTRO USUARIO', `Usuário ${nome} cadastrado`);

    return {
      success: true,
      message: 'Usuário cadastrado com sucesso',
      id: proximoId,
      usuario: {
        id: proximoId,
        nome: nome,
        email: email,
        perfil: perfil,
        acessoVisitantes: acessoVisitantes,
        acessoAcompanhantes: acessoAcompanhantes,
        dataCadastro: dataCadastro,
        status: 'ativo',
        senha: senha,
        unidades: unidadesUnicas.slice()
      }
    };

  } catch (error) {
    registrarLog('ERRO', `Erro ao cadastrar usuário: ${error.toString()}`);
    return { success: false, error: error.toString() };
  }
}

function atualizarUsuario(dados) {
  try {
    var id = parseInt(dados.id, 10);
    if (!id) {
      return { success: false, error: 'ID do usuário inválido' };
    }

    var nome = (dados.nome || '').toString().trim();
    var email = (dados.email || '').toString().trim();
    var perfil = (dados.perfil || '').toString().trim().toLowerCase();
    var senha = (dados.senha || '').toString().trim();
    var status = (dados.status || '').toString().trim().toLowerCase();
    var acessoVisitantes = converterParaBoolean(dados.acessoVisitantes);
    var acessoAcompanhantes = converterParaBoolean(dados.acessoAcompanhantes);
    var unidadesLista = normalizarListaUnidadesParametro(dados.unidades);
    var unidadesUnicas = [];
    var incluiTodas = false;

    unidadesLista.forEach(function(unidade) {
      var chave = unidade.toString().trim();
      if (!chave) {
        return;
      }
      if (normalizarTextoBasico(chave) === 'all') {
        incluiTodas = true;
        return;
      }
      if (unidadesUnicas.indexOf(chave) === -1) {
        unidadesUnicas.push(chave);
      }
    });

    if (incluiTodas || (perfil === 'admin' && unidadesUnicas.length === 0)) {
      unidadesUnicas = ['all'];
    }

    if (!nome || !email || !perfil) {
      return { success: false, error: 'Nome, matrícula e perfil são obrigatórios' };
    }

    if (!senha) {
      return { success: false, error: 'Informe a senha do usuário' };
    }

    if (unidadesUnicas.length === 0 && perfil !== 'admin') {
      return { success: false, error: 'Informe ao menos uma unidade de acesso' };
    }

    var mapasUnidades = obterMapasUnidades();
    var unidadesFormatadas = formatarUnidadesParaRegistro(unidadesUnicas, mapasUnidades);
    var unidadesTexto = unidadesFormatadas.join('; ');

    if (!status || ['ativo', 'inativo'].indexOf(status) === -1) {
      status = 'ativo';
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Usuários');
    if (!sheet || sheet.getLastRow() < 2) {
      return { success: false, error: 'Usuário não encontrado' };
    }

    var estrutura = obterEstruturaPlanilha(sheet);
    var totalColunas = estrutura.ultimaColuna || 10;
    var idIndex = obterIndiceColuna(estrutura, 'id', 0);
    var ultimaLinha = sheet.getLastRow();
    var faixa = sheet.getRange(2, 1, ultimaLinha - 1, totalColunas);
    var valores = faixa.getValues();
    var encontrado = false;

    for (var i = 0; i < valores.length; i++) {
      var valorId = valores[i][idIndex];
      if (parseInt(valorId, 10) === id) {
        definirValorLinha(valores[i], estrutura, 'nome', nome);
        definirValorLinha(valores[i], estrutura, 'email', email);
        definirValorLinha(valores[i], estrutura, 'perfil', perfil);
        definirValorLinha(valores[i], estrutura, 'acesso visitantes', acessoVisitantes);
        definirValorLinha(valores[i], estrutura, 'acesso acompanhantes', acessoAcompanhantes);
        definirValorLinha(valores[i], estrutura, 'status', status);
        definirValorLinha(valores[i], estrutura, 'senha', senha);
        if (!definirValorLinhaFlexivel(valores[i], estrutura, ['unidades', 'unidade', 'acesso unidades'], unidadesTexto)) {
          definirValorLinha(valores[i], estrutura, 'unidades', unidadesTexto);
        }
        encontrado = true;
        break;
      }
    }

    if (!encontrado) {
      return { success: false, error: 'Usuário não encontrado' };
    }

    faixa.setValues(valores);

    registrarLog('ATUALIZAR USUARIO', 'Usuário ' + nome + ' atualizado');

    return {
      success: true,
      message: 'Usuário atualizado com sucesso',
      usuario: {
        id: id,
        nome: nome,
        email: email,
        perfil: perfil,
        acessoVisitantes: acessoVisitantes,
        acessoAcompanhantes: acessoAcompanhantes,
        status: status,
        senha: senha,
        unidades: unidadesUnicas.slice()
      }
    };

  } catch (error) {
    registrarLog('ERRO', 'Erro ao atualizar usuário: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

function excluirUsuario(dados) {
  try {
    var id = parseInt(dados.id, 10);
    if (!id) {
      return { success: false, error: 'ID inválido' };
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Usuários');
    if (!sheet || sheet.getLastRow() < 2) {
      return { success: false, error: 'Usuário não encontrado' };
    }

    var estrutura = obterEstruturaPlanilha(sheet);
    var totalColunas = estrutura.ultimaColuna || 10;
    var idIndex = obterIndiceColuna(estrutura, 'id', 0);
    var ultimaLinha = sheet.getLastRow();
    var valores = sheet.getRange(2, 1, ultimaLinha - 1, totalColunas).getValues();
    var linhaExcluir = -1;
    var nomeUsuario = '';

    for (var i = 0; i < valores.length; i++) {
      var valorId = valores[i][idIndex];
      if (parseInt(valorId, 10) === id) {
        linhaExcluir = i + 2;
        nomeUsuario = obterValorLinha(valores[i], estrutura, 'nome', '');
        break;
      }
    }

    if (linhaExcluir === -1) {
      return { success: false, error: 'Usuário não encontrado' };
    }

    sheet.deleteRow(linhaExcluir);
    registrarLog('EXCLUIR USUARIO', 'Usuário ' + nomeUsuario + ' removido');

    return { success: true };

  } catch (error) {
    registrarLog('ERRO', 'Erro ao excluir usuário: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

function autenticarUsuario(dados) {
  try {
    var login = (dados.usuario || dados.matricula || dados.email || dados.login || '').toString().trim();
    var senhaInformada = (dados.senha || '').toString().trim();

    if (!login || !senhaInformada) {
      return { success: false, error: 'Informe usuário e senha' };
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Usuários');
    if (!sheet || sheet.getLastRow() < 2) {
      return { success: false, error: 'Nenhum usuário cadastrado' };
    }

    var estrutura = obterEstruturaPlanilha(sheet);
    var totalColunas = estrutura.ultimaColuna || 10;
    var mapasUnidades = obterMapasUnidades();
    var dadosUsuarios = sheet.getRange(2, 1, sheet.getLastRow() - 1, totalColunas).getValues();
    var alvoNormalizado = normalizarTextoBasico(login);
    var linhaUsuario = null;
    var indiceLinhaUsuario = -1;

    for (var i = 0; i < dadosUsuarios.length; i++) {
      var linha = dadosUsuarios[i];
      var identificadores = [];

      ['usuario', 'nome', 'matricula', 'email'].forEach(function(chave) {
        var valor = obterValorLinha(linha, estrutura, chave, '');
        if (valor !== null && valor !== undefined) {
          var texto = valor.toString().trim();
          if (texto) {
            identificadores.push(texto);
          }
        }
      });

      var encontrou = identificadores.some(function(valor) {
        return normalizarTextoBasico(valor) === alvoNormalizado;
      });

      if (encontrou) {
        linhaUsuario = linha;
        indiceLinhaUsuario = i;
        break;
      }
    }

    if (!linhaUsuario) {
      return { success: false, error: 'Usuário não encontrado' };
    }

    var status = obterValorLinha(linhaUsuario, estrutura, 'status', '');
    if (normalizarTextoBasico(status) !== 'ativo') {
      return { success: false, error: 'Usuário inativo' };
    }

    var senhaArmazenada = obterValorLinha(linhaUsuario, estrutura, 'senha', '');
    if (senhaInformada !== senhaArmazenada) {
      return { success: false, error: 'Senha incorreta' };
    }

    var unidadesTexto = obterValorLinhaFlexivel(linhaUsuario, estrutura, ['unidades', 'unidade', 'acesso unidades'], '');
    var unidadesLista = resolverIdsUnidadesArmazenadas(unidadesTexto, mapasUnidades);

    var usuarioEncontrado = {
      id: obterValorLinha(linhaUsuario, estrutura, 'id', ''),
      nome: obterValorLinha(linhaUsuario, estrutura, 'nome', ''),
      email: obterValorLinha(linhaUsuario, estrutura, 'email', ''),
      usuario: obterValorLinha(linhaUsuario, estrutura, 'usuario', login) || login,
      perfil: obterValorLinha(linhaUsuario, estrutura, 'perfil', ''),
      acessoVisitantes: converterParaBoolean(obterValorLinha(linhaUsuario, estrutura, 'acesso visitantes', false)),
      acessoAcompanhantes: converterParaBoolean(obterValorLinha(linhaUsuario, estrutura, 'acesso acompanhantes', false)),
      unidades: unidadesLista,
      status: status
    };

    registrarLog('LOGIN', 'Usuário ' + usuarioEncontrado.nome + ' autenticado');

    return {
      success: true,
      usuario: usuarioEncontrado,
      linha: indiceLinhaUsuario + 2
    };

  } catch (error) {
    registrarLog('ERRO', 'Erro ao autenticar usuário: ' + error.toString());
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
    
    var data = sheet.getRange(2, 1, sheet.getLastRow()-1, 13).getValues();
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
          unidade: row[11],
          whatsapp: row[12] || ''
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
        
        var novaLinha;

        if (armario[2] === 'visitante') {
          novaLinha = [
            novoId,
            armario[1], // número
            'livre', // status
            '', // nome
            '', // paciente
            '', // leito
            0, // volumes
            '', // hora início
            '', // hora prevista
            new Date(), // data registro
            armario[3], // unidade
            false, // termo aplicado
            '' // WhatsApp
          ];
        } else {
          novaLinha = [
            novoId,
            armario[1], // número
            'livre', // status
            '', // nome
            '', // paciente
            '', // leito
            0, // volumes
            '', // hora início
            new Date(), // data registro
            '', // WhatsApp
            armario[3], // unidade
            false // termo aplicado
          ];
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
function limparCacheTermos() {
  try {
    CacheService.getScriptCache().remove(TERMOS_CACHE_KEY);
  } catch (erroCache) {
    registrarLog('AVISO_CACHE', 'Falha ao limpar cache de termos: ' + erroCache);
  }
}

function salvarTermoCompleto(dadosTermo) {
  try {
    var orientacoes = dadosTermo.orientacoes;
    if (typeof orientacoes === 'string' && orientacoes !== '') {
      try {
        orientacoes = JSON.parse(orientacoes);
      } catch (erroOrientacoes) {
        orientacoes = orientacoes.split(',').map(function(item) { return item.trim(); }).filter(String);
      }
    }
    if (!Array.isArray(orientacoes)) {
      orientacoes = [];
    }

    var volumes = dadosTermo.volumes;
    if (typeof volumes === 'string' && volumes !== '') {
      try {
        volumes = JSON.parse(volumes);
      } catch (erroVolumes) {
        volumes = [];
      }
    }
    if (!Array.isArray(volumes)) {
      volumes = [];
    }
    volumes = volumes.map(function(item) {
      if (typeof item === 'string') {
        return { quantidade: 0, descricao: item };
      }
      var quantidadeNumero = Number(item.quantidade);
      return {
        quantidade: isNaN(quantidadeNumero) ? 0 : quantidadeNumero,
        descricao: item && item.descricao ? String(item.descricao) : ''
      };
    }).filter(function(item) {
      return item.quantidade > 0 && item.descricao;
    });

    var descricaoVolumes = dadosTermo.descricaoVolumes;
    if (!descricaoVolumes) {
      descricaoVolumes = volumes.map(function(item) {
        return item.quantidade + 'x ' + item.descricao;
      }).join('; ');
    }

    var totalVolumes = volumes.reduce(function(total, volume) {
      return total + (Number(volume.quantidade) || 0);
    }, 0);

    dadosTermo.orientacoes = orientacoes;
    dadosTermo.volumes = volumes;
    dadosTermo.descricaoVolumes = descricaoVolumes;

    // 1. Salvar na aba "Termos de Responsabilidade"
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Termos de Responsabilidade');

    if (!sheet) {
      throw new Error('Aba "Termos de Responsabilidade" não encontrada');
    }

    var dadosExistentes = sheet.getDataRange().getValues();
    var linhaExistente = -1;
    var termoId = null;
    var aplicadoEm = new Date();

    for (var i = dadosExistentes.length - 1; i >= 1; i--) {
      if (dadosExistentes[i][1] == dadosTermo.armarioId) {
        if (!dadosExistentes[i][17]) { // Termo ainda não finalizado
          linhaExistente = i + 1;
          termoId = dadosExistentes[i][0];
          aplicadoEm = dadosExistentes[i][16] || new Date();
          break;
        }
      }
    }

    if (linhaExistente === -1) {
      var lastRow = sheet.getLastRow();
      termoId = lastRow > 1 ? Math.max.apply(null, sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat()) + 1 : 1;
      linhaExistente = lastRow + 1;
    }

    var valorAtualAssinatura = '';
    if (linhaExistente <= dadosExistentes.length && linhaExistente - 1 >= 0) {
      var linhaAtual = dadosExistentes[linhaExistente - 1];
      if (linhaAtual && linhaAtual.length > 18) {
        valorAtualAssinatura = linhaAtual[18];
      }
    }

    var assinaturasInfo = normalizarAssinaturas(valorAtualAssinatura);
    assinaturasInfo.inicial = dadosTermo.assinaturaBase64 || assinaturasInfo.inicial || '';

    var linhaDados = [
      termoId,
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
      orientacoes.join(','),
      JSON.stringify(volumes),
      descricaoVolumes,
      aplicadoEm,
      '',
      JSON.stringify(assinaturasInfo)
    ];

    sheet.getRange(linhaExistente, 1, 1, linhaDados.length).setValues([linhaDados]);

    // 2. Atualizar status do armário na aba "Acompanhantes"
    var sheetAcompanhantes = ss.getSheetByName('Acompanhantes');
    var dataAcompanhantes = sheetAcompanhantes.getDataRange().getValues();

    for (var i = 1; i < dataAcompanhantes.length; i++) {
      if (dataAcompanhantes[i][0] == dadosTermo.armarioId) {
        // Atualizar volumes e marcar termo como aplicado
        sheetAcompanhantes.getRange(i + 1, 7).setValue(totalVolumes);
        sheetAcompanhantes.getRange(i + 1, 12).setValue(true); // Termo iniciado
        break;
      }
    }

    limparCacheTermos();

    registrarLog('TERMO_APLICADO', `Termo inicial registrado para armário ${dadosTermo.numeroArmario}`);

    return {
      success: true,
      message: 'Termo registrado. Finalize na liberação para gerar o PDF.',
      termoId: termoId
    };

  } catch (error) {
    registrarLog('ERRO_TERMO', `Erro ao salvar termo: ${error.toString()}`);
    return { success: false, error: error.toString() };
  }
}

function normalizarAssinaturas(valor) {
  var info = {
    inicial: '',
    final: '',
    metodoFinal: '',
    cpfFinal: '',
    finalizadoEm: ''
  };

  if (!valor) {
    return info;
  }

  if (typeof valor === 'string') {
    try {
      var json = JSON.parse(valor);
      info.inicial = json.inicial || '';
      info.final = json.final || '';
      info.metodoFinal = json.metodoFinal || '';
      info.cpfFinal = json.cpfFinal || '';
      info.finalizadoEm = json.finalizadoEm || '';
      return info;
    } catch (erro) {
      info.inicial = valor;
      return info;
    }
  }

  if (typeof valor === 'object') {
    info.inicial = valor.inicial || '';
    info.final = valor.final || '';
    info.metodoFinal = valor.metodoFinal || '';
    info.cpfFinal = valor.cpfFinal || '';
    info.finalizadoEm = valor.finalizadoEm || '';
  }

  return info;
}

function obterTermosRegistrados() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Termos de Responsabilidade');

  if (!sheet || sheet.getLastRow() < 2) {
    return { sheet: sheet, termos: [] };
  }

  var cache = CacheService.getScriptCache();
  var dadosCache = null;
  try {
    dadosCache = cache.get(TERMOS_CACHE_KEY);
    if (dadosCache) {
      var termosCache = JSON.parse(dadosCache);
      return { sheet: sheet, termos: termosCache };
    }
  } catch (erroCache) {
    cache.remove(TERMOS_CACHE_KEY);
    registrarLog('AVISO_CACHE', 'Falha ao ler cache de termos: ' + erroCache);
  }

  var data = sheet.getDataRange().getValues();
  var termos = [];

  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) {
      continue;
    }

    var assinaturas = normalizarAssinaturas(data[i][18]);
    var orientacoes = data[i][13] ? data[i][13].split(',').filter(String) : [];
    var volumes = [];

    if (data[i][14]) {
      try {
        volumes = JSON.parse(data[i][14]);
      } catch (erroVolume) {
        volumes = [];
      }
    }

    termos.push({
      linha: i + 1,
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
      orientacoes: orientacoes,
      volumes: Array.isArray(volumes) ? volumes : [],
      descricaoVolumes: data[i][15],
      aplicadoEm: data[i][16],
      pdfUrl: data[i][17],
      assinaturas: assinaturas
    });
  }

  try {
    cache.put(TERMOS_CACHE_KEY, JSON.stringify(termos), TERMOS_CACHE_TTL);
  } catch (erroArmazenamento) {
    registrarLog('AVISO_CACHE', 'Falha ao armazenar cache de termos: ' + erroArmazenamento);
  }

  return { sheet: sheet, termos: termos };
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
    
    for (var i = data.length - 1; i >= 1; i--) {
      if (data[i][1] == dados.armarioId) {
        var assinaturas = normalizarAssinaturas(data[i][18]);
        var orientacoesSalvas = data[i][13] ? data[i][13].split(',').filter(String) : [];
        var volumesSalvos = [];
        if (data[i][14]) {
          try {
            volumesSalvos = JSON.parse(data[i][14]);
          } catch (erroVolumes) {
            volumesSalvos = [];
          }
        }

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
          orientacoes: orientacoesSalvas,
          volumes: Array.isArray(volumesSalvos) ? volumesSalvos : [],
          descricaoVolumes: data[i][15],
          aplicadoEm: data[i][16],
          pdfUrl: data[i][17],
          assinaturaBase64: assinaturas.inicial,
          assinaturas: assinaturas,
          finalizadoEm: assinaturas.finalizadoEm,
          metodoFinal: assinaturas.metodoFinal,
          cpfConfirmacao: assinaturas.cpfFinal
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

function finalizarTermo(dados) {
  try {
    var armarioId = parseInt(dados.armarioId, 10);
    if (!armarioId) {
      return { success: false, error: 'ID do armário inválido' };
    }

    var metodo = (dados.metodo || 'assinatura').toString();
    var confirmacao = dados.confirmacao || '';
    var assinaturaFinal = dados.assinaturaFinal || '';

    var termosInfo = obterTermosRegistrados();
    if (!termosInfo.sheet) {
      return { success: false, error: 'Aba "Termos de Responsabilidade" não encontrada' };
    }

    var termoEncontrado = null;
    for (var i = termosInfo.termos.length - 1; i >= 0; i--) {
      if (termosInfo.termos[i].armarioId == armarioId && !termosInfo.termos[i].pdfUrl) {
        termoEncontrado = termosInfo.termos[i];
        break;
      }
    }

    if (!termoEncontrado) {
      for (var j = termosInfo.termos.length - 1; j >= 0; j--) {
        if (termosInfo.termos[j].armarioId == armarioId) {
          termoEncontrado = termosInfo.termos[j];
          break;
        }
      }
    }

    if (!termoEncontrado) {
      return { success: false, error: 'Termo não localizado para este armário' };
    }

    var assinaturas = termoEncontrado.assinaturas || normalizarAssinaturas('');
    var agora = new Date();
    assinaturas.metodoFinal = metodo;
    assinaturas.cpfFinal = metodo === 'cpf' ? confirmacao : '';
    assinaturas.finalizadoEm = agora;
    assinaturas.final = metodo === 'assinatura' ? assinaturaFinal : '';

    var movimentacoesResultado = getMovimentacoes({ armarioId: armarioId });
    var movimentacoes = [];
    if (movimentacoesResultado && movimentacoesResultado.success && Array.isArray(movimentacoesResultado.data)) {
      movimentacoes = movimentacoesResultado.data;
    } else if (movimentacoesResultado && movimentacoesResultado.success) {
      registrarLog('AVISO_TERMO', 'Dados de movimentações inválidos ao finalizar termo do armário ' + termoEncontrado.numeroArmario);
    } else if (movimentacoesResultado && !movimentacoesResultado.success) {
      registrarLog('AVISO_TERMO', 'Movimentações indisponíveis ao finalizar termo do armário ' + termoEncontrado.numeroArmario + ': ' + (movimentacoesResultado.error || 'dados inválidos'));
    }

    var dadosPDF = {
      numeroArmario: termoEncontrado.numeroArmario,
      paciente: termoEncontrado.paciente,
      prontuario: termoEncontrado.prontuario,
      nascimento: termoEncontrado.nascimento,
      setor: termoEncontrado.setor,
      leito: termoEncontrado.leito,
      consciente: termoEncontrado.consciente,
      acompanhante: termoEncontrado.acompanhante,
      telefone: termoEncontrado.telefone,
      documento: termoEncontrado.documento,
      parentesco: termoEncontrado.parentesco,
      orientacoes: termoEncontrado.orientacoes,
      volumes: termoEncontrado.volumes,
      descricaoVolumes: termoEncontrado.descricaoVolumes,
      aplicadoEm: termoEncontrado.aplicadoEm,
      finalizadoEm: agora,
      assinaturaInicial: assinaturas.inicial,
      assinaturaFinal: assinaturas.final,
      metodoFinal: assinaturas.metodoFinal,
      cpfFinal: assinaturas.cpfFinal,
      movimentacoes: movimentacoes
    };

    var resultadoPDF = gerarESalvarTermoPDF(dadosPDF);
    if (!resultadoPDF.success) {
      throw new Error(resultadoPDF.error || 'Falha ao gerar PDF');
    }

    termosInfo.sheet.getRange(termoEncontrado.linha, 18).setValue(resultadoPDF.pdfUrl);
    termosInfo.sheet.getRange(termoEncontrado.linha, 19).setValue(JSON.stringify(assinaturas));

    limparCacheTermos();

    registrarLog('TERMO_FINALIZADO', 'Termo finalizado para armário ' + termoEncontrado.numeroArmario);

    return {
      success: true,
      pdfUrl: resultadoPDF.pdfUrl,
      finalizadoEm: agora
    };

  } catch (error) {
    registrarLog('ERRO_TERMO', 'Erro ao finalizar termo: ' + error.toString());
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
  var hospitalNome = 'Hospital Universitário do Ceará';
  var orientacoesPredefinidas = [
    'Seus pertences estão sob sua guarda e responsabilidade.',
    'Em piora clínica, os pertences serão recolhidos e protocolados no NAC.',
    'Após 15 dias da alta/transferência, itens não retirados poderão ser descartados conforme normas.'
  ];
  var orientacoes = [];
  if (Array.isArray(dadosTermo.orientacoes) && dadosTermo.orientacoes.length) {
    orientacoes = dadosTermo.orientacoes.map(function(item) {
      switch (item) {
        case 'ori1':
          return orientacoesPredefinidas[0];
        case 'ori2':
          return orientacoesPredefinidas[1];
        case 'ori3':
          return orientacoesPredefinidas[2];
        default:
          return item;
      }
    }).filter(function(texto) { return texto && texto.trim(); });
  }
  if (!orientacoes.length) {
    orientacoes = orientacoesPredefinidas;
  }

  var volumesLista = Array.isArray(dadosTermo.volumes) ? dadosTermo.volumes : [];
  var movimentacoesLista = Array.isArray(dadosTermo.movimentacoes) ? dadosTermo.movimentacoes : [];

  function formatarMovimentacao(mov) {
    var data = mov.data ? new Date(mov.data) : null;
    var hora = mov.hora ? new Date('1970-01-01T' + mov.hora + 'Z') : null;
    var dataFormatada = data && !isNaN(data.getTime())
      ? Utilities.formatDate(data, 'America/Sao_Paulo', 'dd/MM/yyyy')
      : (mov.data || '');
    var horaFormatada = mov.hora
      ? mov.hora
      : (hora && !isNaN(hora.getTime())
          ? Utilities.formatDate(hora, 'America/Sao_Paulo', 'HH:mm')
          : '');
    return {
      data: dataFormatada,
      hora: horaFormatada,
      tipo: (mov.tipo || '').toString().toUpperCase(),
      descricao: mov.descricao || '',
      responsavel: mov.responsavel || ''
    };
  }

  var movimentosNormalizados = movimentacoesLista.map(formatarMovimentacao);
  while (movimentosNormalizados.length < 8) {
    movimentosNormalizados.push({ data: '', hora: '', tipo: '', descricao: '', responsavel: '' });
  }

  var assinaturaInicialHtml = dadosTermo.assinaturaInicial
    ? '<img src="data:image/png;base64,' + dadosTermo.assinaturaInicial + '" class="assinatura-img" alt="Assinatura inicial" />'
    : '<div class="assinatura-linha">Assinatura não registrada digitalmente.</div>';

  var assinaturaFinalHtml = '';
  if (dadosTermo.metodoFinal === 'cpf' && dadosTermo.cpfFinal) {
    assinaturaFinalHtml = '<div class="assinatura-linha">Confirmação por CPF: ' + dadosTermo.cpfFinal + '</div>';
  } else if (dadosTermo.assinaturaFinal) {
    assinaturaFinalHtml = '<img src="data:image/png;base64,' + dadosTermo.assinaturaFinal + '" class="assinatura-img" alt="Assinatura final" />';
  } else {
    assinaturaFinalHtml = '<div class="assinatura-linha">Assinatura final não registrada.</div>';
  }

  var partes = [];
  partes.push('<!DOCTYPE html>');
  partes.push('<html>');
  partes.push('<head>');
  partes.push('<base target="_top">');
  partes.push('<style>');
  partes.push('  body { font-family: Arial, sans-serif; margin: 24px; color: #0b1324; }');
  partes.push('  h1, h2, h3 { margin: 0; }');
  partes.push('  .header { display: flex; justify-content: space-between; align-items: center; border-bottom: 2px solid #0b1324; padding-bottom: 12px; margin-bottom: 16px; }');
  partes.push('  .header h1 { font-size: 20px; text-transform: uppercase; }');
  partes.push('  .info-grid { display: grid; grid-template-columns: repeat(2, 1fr); gap: 6px 24px; margin-bottom: 16px; font-size: 13px; }');
  partes.push('  .section-title { font-weight: bold; text-transform: uppercase; font-size: 13px; margin: 18px 0 8px; }');
  partes.push('  .orientacoes { font-size: 12px; margin-left: 18px; }');
  partes.push('  .volumes-table { width: 100%; border-collapse: collapse; margin-top: 6px; font-size: 12px; }');
  partes.push('  .volumes-table th, .volumes-table td { border: 1px solid #0b1324; padding: 6px; text-align: left; }');
  partes.push('  .assinatura-box { margin-top: 20px; text-align: center; }');
  partes.push('  .assinatura-img { max-width: 260px; max-height: 120px; border: 1px solid #d0d7e2; padding: 6px; }');
  partes.push('  .assinatura-linha { border-bottom: 1px solid #0b1324; display: inline-block; padding: 4px 16px; min-width: 240px; font-size: 12px; }');
  partes.push('  .footer { margin-top: 18px; font-size: 10px; text-align: center; color: #3d4a63; }');
  partes.push('  .page-break { page-break-before: always; }');
  partes.push('  .mov-table { width: 100%; border-collapse: collapse; font-size: 11px; margin-top: 6px; }');
  partes.push('  .mov-table th, .mov-table td { border: 1px solid #0b1324; padding: 5px; vertical-align: top; }');
  partes.push('  .mov-table th { background: #eef3fb; }');
  partes.push('  .devolucao-box, .descarte-box { border: 1px solid #0b1324; padding: 10px; margin-top: 10px; min-height: 70px; }');
  partes.push('  .observacoes { border: 1px solid #0b1324; min-height: 90px; margin-top: 12px; padding: 8px; font-size: 12px; }');
  partes.push('  .label { font-weight: bold; }');
  partes.push('</style>');
  partes.push('</head>');
  partes.push('<body>');
  partes.push('<div class="header">');
  partes.push('  <div>');
  partes.push('    <h1>Termo de Responsabilidade</h1>');
  partes.push('    <h3>' + hospitalNome + '</h3>');
  partes.push('  </div>');
  partes.push('  <div style="text-align:right;font-size:12px;">');
  partes.push('    <div><strong>Nº do Armário:</strong> ' + (dadosTermo.numeroArmario || '') + '</div>');
  partes.push('    <div><strong>Data de início:</strong> ' + formatarDataParaHTML(dadosTermo.aplicadoEm) + '</div>');
  partes.push('  </div>');
  partes.push('</div>');
  partes.push('<div class="section-title">Dados do Paciente</div>');
  partes.push('<div class="info-grid">');
  partes.push('  <div><span class="label">Nome:</span> ' + (dadosTermo.paciente || '') + '</div>');
  partes.push('  <div><span class="label">Prontuário:</span> ' + (dadosTermo.prontuario || '') + '</div>');
  partes.push('  <div><span class="label">Data de nascimento:</span> ' + formatarDataParaHTML(dadosTermo.nascimento) + '</div>');
  partes.push('  <div><span class="label">Setor/Leito:</span> ' + (dadosTermo.setor || '') + ' ' + (dadosTermo.leito || '') + '</div>');
  partes.push('  <div><span class="label">Paciente consciente/orientado:</span> ' + (dadosTermo.consciente || '') + '</div>');
  partes.push('</div>');
  partes.push('<div class="section-title">Responsável pelo Armário</div>');
  partes.push('<div class="info-grid">');
  partes.push('  <div><span class="label">Nome:</span> ' + (dadosTermo.acompanhante || '') + '</div>');
  partes.push('  <div><span class="label">Documento:</span> ' + (dadosTermo.documento || 'Não informado') + '</div>');
  partes.push('  <div><span class="label">Telefone:</span> ' + (dadosTermo.telefone || 'Não informado') + '</div>');
  partes.push('  <div><span class="label">Parentesco:</span> ' + (dadosTermo.parentesco || 'Não informado') + '</div>');
  partes.push('</div>');
  partes.push('<div class="section-title">Orientações repassadas</div>');
  partes.push('<ul class="orientacoes">');
  orientacoes.forEach(function(item) {
    partes.push('<li>' + item + '</li>');
  });
  partes.push('</ul>');
  partes.push('<div class="section-title">Volumes armazenados</div>');
  partes.push('<table class="volumes-table">');
  partes.push('  <thead><tr><th style="width:20%">Quantidade</th><th>Descrição</th></tr></thead>');
  partes.push('  <tbody>');
  if (volumesLista.length) {
    volumesLista.forEach(function(volume) {
      partes.push('<tr><td>' + (volume.quantidade || '') + '</td><td>' + (volume.descricao || '') + '</td></tr>');
    });
  } else {
    partes.push('<tr><td colspan="2">Sem volumes informados.</td></tr>');
  }
  partes.push('  </tbody>');
  partes.push('</table>');
  partes.push('<div class="assinatura-box">');
  partes.push('  <div class="section-title">Assinatura do responsável - Etapa inicial</div>');
  partes.push(assinaturaInicialHtml);
  partes.push('  <div style="margin-top:6px; font-size:12px;">' + (dadosTermo.acompanhante || '') + '</div>');
  partes.push('</div>');
  partes.push('<div class="footer">Primeira etapa concluída em ' + Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm') + '.</div>');
  partes.push('<div class="page-break"></div>');
  partes.push('<div class="section-title">Movimentações registradas</div>');
  partes.push('<table class="mov-table">');
  partes.push('  <thead><tr><th style="width:16%">Data</th><th style="width:12%">Hora</th><th style="width:18%">Tipo</th><th>Descrição</th><th style="width:20%">Responsável</th></tr></thead>');
  partes.push('  <tbody>');
  movimentosNormalizados.forEach(function(mov) {
    partes.push('<tr><td>' + (mov.data || '') + '</td><td>' + (mov.hora || '') + '</td><td>' + (mov.tipo || '') + '</td><td>' + (mov.descricao || '') + '</td><td>' + (mov.responsavel || '') + '</td></tr>');
  });
  partes.push('  </tbody>');
  partes.push('</table>');
  partes.push('<div class="section-title">Devolução de pertences</div>');
  partes.push('<div class="devolucao-box">Data: _____________ &nbsp;&nbsp; Conferente: __________________________</div>');
  partes.push('<div class="section-title">Descarte de pertences</div>');
  partes.push('<div class="descarte-box">');
  partes.push('  <p>Declaramos que o descarte dos devidos esclarecimentos por parte da equipe do ' + hospitalNome + ' sobre os pertences deixados pelo paciente.</p>');
  partes.push('  <p>Informamos que o Núcleo Ético Clínica está ciente dos procedimentos adotados. O ato da entrega está acompanhado por dois colaboradores designados para este fim.</p>');
  partes.push('</div>');
  partes.push('<div class="section-title">Observações complementares</div>');
  partes.push('<div class="observacoes"></div>');
  partes.push('<div class="assinatura-box">');
  partes.push('  <div class="section-title">Assinatura de encerramento</div>');
  partes.push(assinaturaFinalHtml);
  partes.push('  <div style="margin-top:6px; font-size:12px;">' + (dadosTermo.acompanhante || '') + '</div>');
  partes.push('  <div style="margin-top:4px; font-size:11px;">Encerrado em: ' + formatarDataParaHTML(dadosTermo.finalizadoEm) + '</div>');
  partes.push('</div>');
  partes.push('<div class="footer">Documento gerado automaticamente em ' + Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm') + ' - ' + hospitalNome + '.</div>');
  partes.push('</body>');
  partes.push('</html>');

  return partes.join('');
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
    var perfil = normalizarTextoBasico(tipoUsuario);
    if (!perfil) {
      perfil = 'geral';
    }

    if (perfil === 'admin' || perfil === 'ambos' || perfil === 'geral' || perfil === 'todos') {
      tipos = ['visitante', 'acompanhante'];
    } else if (perfil === 'visitante') {
      tipos = ['visitante'];
    } else if (perfil === 'acompanhante') {
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
