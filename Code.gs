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

function normalizarNumeroArmario(valor) {
  if (valor === null || valor === undefined) {
    return '';
  }
  return valor.toString().trim();
}

function obterChaveNumeroArmario(numero) {
  var numeroNormalizado = normalizarNumeroArmario(numero);
  return numeroNormalizado ? numeroNormalizado : '__sem_numero__';
}

function montarChaveArmarioInterface(tipo, numero, idPlanilha) {
  var tipoNormalizado = normalizarTextoBasico(tipo) || 'geral';
  var numeroNormalizado = normalizarNumeroArmario(numero);
  if (numeroNormalizado) {
    return tipoNormalizado + ':' + numeroNormalizado;
  }
  var idTexto = '';
  if (idPlanilha !== null && idPlanilha !== undefined) {
    idTexto = idPlanilha.toString().trim();
  }
  return tipoNormalizado + ':id-' + (idTexto || Utilities.getUuid());
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
  if (Array.isArray(chave)) {
    for (var i = 0; i < chave.length; i++) {
      var indiceFlexivel = obterIndiceColuna(estrutura, chave[i], null);
      if (indiceFlexivel !== null && indiceFlexivel !== undefined) {
        return indiceFlexivel;
      }
    }
    return padrao;
  }

  if (chave === null || chave === undefined) {
    return padrao;
  }

  var chaveNormalizada = normalizarTextoBasico(chave);
  if (estrutura.mapaIndices.hasOwnProperty(chaveNormalizada)) {
    return estrutura.mapaIndices[chaveNormalizada];
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

var CABECALHOS_WHATSAPP = ['whatsapp', 'wpp', 'whats app', 'whatsap', 'zap'];
var CABECALHOS_NOME_VISITANTE = ['nome visitante', 'visitante', 'nome do visitante'];
var CABECALHOS_NOME_ACOMPANHANTE = ['nome acompanhante', 'acompanhante', 'nome do acompanhante', 'responsavel', 'responsável'];

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

// Configurações gerais de cache para otimizar leituras
const CACHE_PREFIXO = 'locknac_cache_v1';
const CACHE_TTL_PADRAO = 60; // segundos
const CACHE_TTL_ARMARIOS = 45;
const CACHE_TTL_HISTORICO = 90;
const CACHE_TTL_MOVIMENTACOES = 45;

function montarChaveCache() {
  var partes = Array.prototype.slice.call(arguments).filter(function(parte) {
    return parte !== null && parte !== undefined && parte !== '';
  }).map(function(parte) {
    if (typeof parte === 'object') {
      try {
        return JSON.stringify(parte);
      } catch (erro) {
        return '';
      }
    }
    return parte.toString().trim().toLowerCase().replace(/\s+/g, '-');
  });

  if (!partes.length) {
    return CACHE_PREFIXO;
  }

  return CACHE_PREFIXO + ':' + partes.join(':');
}

function executarComCache(chave, ttl, fornecedor) {
  if (!chave) {
    return fornecedor();
  }

  var cache = CacheService.getScriptCache();

  try {
    var armazenado = cache.get(chave);
    if (armazenado) {
      return JSON.parse(armazenado);
    }
  } catch (erroLeitura) {
    try {
      cache.remove(chave);
    } catch (erroRemocao) {
      // Ignorado propositalmente
    }
  }

  var resultado = fornecedor();

  if (resultado && resultado.success) {
    try {
      cache.put(chave, JSON.stringify(resultado), ttl || CACHE_TTL_PADRAO);
    } catch (erroGravacao) {
      // Falhas de cache não devem interromper o fluxo principal
    }
  }

  return resultado;
}

function limparCaches(chaves) {
  if (!chaves) {
    return;
  }

  var lista = Array.isArray(chaves) ? chaves : [chaves];
  if (!lista.length) {
    return;
  }

  var cache = CacheService.getScriptCache();
  lista.forEach(function(chave) {
    if (!chave) {
      return;
    }
    try {
      cache.remove(chave);
    } catch (erroRemocao) {
      // Ignorado propositalmente
    }
  });
}

function limparCacheArmarios() {
  limparCaches([
    montarChaveCache('armarios', 'visitante'),
    montarChaveCache('armarios', 'acompanhante'),
    montarChaveCache('armarios', 'geral')
  ]);
}

function limparCacheUsuarios() {
  limparCaches(montarChaveCache('usuarios'));
}

function limparCacheHistorico() {
  limparCaches([
    montarChaveCache('historico', 'visitante'),
    montarChaveCache('historico', 'acompanhante')
  ]);
}

function limparCacheCadastroArmarios() {
  limparCaches(montarChaveCache('cadastro-armarios'));
}

function limparCacheUnidades() {
  limparCaches(montarChaveCache('unidades'));
}

function limparCacheMovimentacoes(armarioId, numeroArmario, tipo) {
  var idTexto = armarioId !== undefined && armarioId !== null ? armarioId.toString().trim() : '';
  var numeroTexto = normalizarNumeroArmario(numeroArmario);
  var tipoTexto = tipo ? normalizarTextoBasico(tipo) : '';
  var chaveEspecifica = idTexto ? montarChaveCache('movimentacoes', [idTexto, numeroTexto, tipoTexto].join('|')) : null;
  var chaveLegado = idTexto ? montarChaveCache('movimentacoes', idTexto) : null;
  limparCaches([
    chaveEspecifica,
    chaveLegado,
    montarChaveCache('movimentacoes', 'todos')
  ]);
}

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
        cabecalhos: ['ID', 'ArmarioID', 'NumeroArmario', 'Tipo', 'Descricao', 'Responsavel', 'Data', 'Hora', 'DataHoraRegistro', 'Status']
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

    limparCacheArmarios();
    limparCacheUsuarios();
    limparCacheHistorico();
    limparCacheCadastroArmarios();
    limparCacheUnidades();
    limparCacheMovimentacoes();

    return { success: true, message: 'Planilha inicializada com sucesso' };
    
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

function obterFusoHorarioPadrao() {
  var timezone = '';
  try {
    timezone = Session.getScriptTimeZone();
  } catch (erro) {
    timezone = '';
  }
  return timezone || 'America/Sao_Paulo';
}

function formatarDataPlanilha(valor) {
  if (!valor) {
    return '';
  }
  var timezone = obterFusoHorarioPadrao();
  if (Object.prototype.toString.call(valor) === '[object Date]' && !isNaN(valor.getTime())) {
    return Utilities.formatDate(valor, timezone, 'dd/MM/yyyy');
  }
  if (typeof valor === 'string') {
    var texto = valor.trim();
    if (!texto) {
      return '';
    }
    if (/^\d{4}-\d{2}-\d{2}$/.test(texto)) {
      var partesData = texto.split('-');
      var ano = parseInt(partesData[0], 10);
      var mes = parseInt(partesData[1], 10) - 1;
      var dia = parseInt(partesData[2], 10);
      if (!isNaN(ano) && !isNaN(mes) && !isNaN(dia)) {
        var dataLocal = new Date(ano, mes, dia);
        return Utilities.formatDate(dataLocal, timezone, 'dd/MM/yyyy');
      }
    }
    if (/^\d{2}\/\d{2}\/\d{4}$/.test(texto)) {
      return texto;
    }
    var textoISO = texto.replace(' ', 'T');
    var data = new Date(textoISO);
    if (!isNaN(data.getTime())) {
      return Utilities.formatDate(data, timezone, 'dd/MM/yyyy');
    }
  }
  return valor;
}

function formatarHorarioPlanilha(valor) {
  if (!valor) {
    return '';
  }
  var timezone = obterFusoHorarioPadrao();
  if (Object.prototype.toString.call(valor) === '[object Date]' && !isNaN(valor.getTime())) {
    var formato = valor.getFullYear() <= 1900 ? 'HH:mm' : 'dd/MM/yyyy HH:mm';
    return Utilities.formatDate(valor, timezone, formato);
  }
  if (typeof valor === 'string') {
    var texto = valor.trim();
    if (!texto) {
      return '';
    }
    if (/^\d{1,2}:\d{2}(:\d{2})?$/.test(texto)) {
      return texto.slice(0, 5);
    }
    var isoSemFuso = texto.match(/^(\d{4}-\d{2}-\d{2})[ T](\d{2}:\d{2})(?::\d{2})?$/);
    if (isoSemFuso) {
      return isoSemFuso[2];
    }
    var textoISO = texto.replace(' ', 'T');
    var data = new Date(textoISO);
    if (!isNaN(data.getTime())) {
      return Utilities.formatDate(data, timezone, 'HH:mm');
    }
  }
  return valor;
}

function determinarResponsavelRegistro(valorPreferencial) {
  if (valorPreferencial !== undefined && valorPreferencial !== null) {
    var texto = valorPreferencial.toString().trim();
    if (texto) {
      return texto;
    }
  }
  try {
    var usuarioAtivo = Session.getActiveUser();
    if (usuarioAtivo && typeof usuarioAtivo.getEmail === 'function') {
      var emailAtivo = usuarioAtivo.getEmail();
      if (emailAtivo) {
        return emailAtivo;
      }
    }
  } catch (erroUsuarioAtivo) {
    // Ignora erro ao obter usuário ativo
  }
  try {
    var usuarioEfetivo = Session.getEffectiveUser();
    if (usuarioEfetivo && typeof usuarioEfetivo.getEmail === 'function') {
      var emailEfetivo = usuarioEfetivo.getEmail();
      if (emailEfetivo) {
        return emailEfetivo;
      }
    }
  } catch (erroUsuarioEfetivo) {
    // Ignora erro ao obter usuário efetivo
  }
  return '';
}

function obterDataHoraAtualFormatada() {
  var agora = new Date();
  var timezone = obterFusoHorarioPadrao();
  return {
    data: agora,
    horaCurta: Utilities.formatDate(agora, timezone, 'HH:mm'),
    dataHoraIso: Utilities.formatDate(agora, timezone, "yyyy-MM-dd'T'HH:mm:ss")
  };
}

function converterParaDataHoraIso(valor, padrao) {
  if (Object.prototype.toString.call(valor) === '[object Date]' && !isNaN(valor.getTime())) {
    return Utilities.formatDate(valor, obterFusoHorarioPadrao(), "yyyy-MM-dd'T'HH:mm:ss");
  }
  if (valor && typeof valor === 'string') {
    return valor;
  }
  return padrao !== undefined ? padrao : '';
}

// Adicionar dados iniciais de exemplo
function adicionarDadosIniciais() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Cadastrar alguns armários físicos
  var cadastroSheet = ss.getSheetByName('Cadastro Armários');
  if (cadastroSheet.getLastRow() === 1) {
    var dataCadastroArmarios = obterDataHoraAtualFormatada().dataHoraIso;
    var armariosIniciais = [
      ['V-01', 'visitante', 'NAC Eletiva', 'Bloco A - Térreo', 'ativo', dataCadastroArmarios],
      ['V-02', 'visitante', 'NAC Eletiva', 'Bloco A - Térreo', 'ativo', dataCadastroArmarios],
      ['V-03', 'visitante', 'UIB', 'Bloco A - Térreo', 'ativo', dataCadastroArmarios],
      ['V-04', 'visitante', 'UIB', 'Bloco A - Térreo', 'ativo', dataCadastroArmarios],
      ['A-01', 'acompanhante', 'NAC Eletiva', 'Bloco B - 1º Andar', 'ativo', dataCadastroArmarios],
      ['A-02', 'acompanhante', 'UIB', 'Bloco B - 1º Andar', 'ativo', dataCadastroArmarios],
      ['A-03', 'acompanhante', 'UIB', 'Bloco B - 1º Andar', 'ativo', dataCadastroArmarios]
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
    var dataCadastroUsuario = obterDataHoraAtualFormatada().dataHoraIso;
    usuariosSheet.getRange(2, 1, 1, 10)
      .setValues([[1, 'Administrador', 'admin', 'admin', true, true, dataCadastroUsuario, 'ativo', 'admin123', 'all']]);
  }

  // Cadastrar unidades iniciais
  var unidadesSheet = ss.getSheetByName('Unidades');
  if (unidadesSheet && unidadesSheet.getLastRow() === 1) {
    var dataCadastroUnidades = obterDataHoraAtualFormatada().dataHoraIso;
    var unidadesIniciais = [
      [1, 'NAC Eletiva', 'ativa', dataCadastroUnidades],
      [2, 'UIB', 'ativa', dataCadastroUnidades]
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
        return ContentService.createTextOutput(JSON.stringify(liberarArmario(
          e.parameter.id,
          e.parameter.tipo,
          e.parameter.numero,
          e.parameter.usuarioResponsavel
        )))
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

      case 'getSetores':
        return ContentService.createTextOutput(JSON.stringify(getSetores()))
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
  var tipoNormalizadoOriginal = normalizarTextoBasico(tipo);
  if (!tipoNormalizadoOriginal) {
    tipoNormalizadoOriginal = 'geral';
  }

  var chaveCacheTipo = tipoNormalizadoOriginal;
  if (chaveCacheTipo === 'admin' || chaveCacheTipo === 'ambos' || chaveCacheTipo === 'todos') {
    chaveCacheTipo = 'geral';
  }

  var chaveCache = montarChaveCache('armarios', chaveCacheTipo);

  return executarComCache(chaveCache, CACHE_TTL_ARMARIOS, function() {
    try {
      var tipoNormalizado = tipoNormalizadoOriginal;
      var incluirTermos = tipoNormalizado === 'acompanhante' || tipoNormalizado === 'admin' ||
        tipoNormalizado === 'ambos' || tipoNormalizado === 'todos' || tipoNormalizado === 'geral';
      var termosMap = {};

      if (incluirTermos) {
        var termosInfo = obterTermosRegistrados();
        termosInfo.termos.forEach(function(termo) {
          if (!termo) {
            return;
          }

          var chaveId = '';
          if (termo.armarioId !== null && termo.armarioId !== undefined) {
            chaveId = termo.armarioId.toString().trim();
          }

          if (!chaveId) {
            return;
          }

          if (!termosMap[chaveId]) {
            termosMap[chaveId] = {};
          }

          var numeroChave = obterChaveNumeroArmario(termo.numeroArmario);
          var termoAtual = termosMap[chaveId][numeroChave];
          var termoFinalizado = termo.finalizado;
          if (termoFinalizado === undefined) {
            var statusTermo = normalizarTextoBasico(termo.status);
            termoFinalizado = Boolean(termo.pdfUrl || (termo.assinaturas && termo.assinaturas.finalizadoEm) || statusTermo === 'finalizado');
          }

          if (!termoAtual) {
            termosMap[chaveId][numeroChave] = termo;
            return;
          }

          var atualFinalizado = termoAtual.finalizado;
          if (atualFinalizado === undefined) {
            var statusAtual = normalizarTextoBasico(termoAtual.status);
            atualFinalizado = Boolean(termoAtual.pdfUrl || (termoAtual.assinaturas && termoAtual.assinaturas.finalizadoEm) || statusAtual === 'finalizado');
          }

          if (!termoFinalizado && atualFinalizado) {
            termosMap[chaveId][numeroChave] = termo;
          } else if (termoFinalizado === atualFinalizado) {
            var idAtualNumero = Number(termoAtual.id) || 0;
            var idNovoNumero = Number(termo.id) || 0;
            if (idNovoNumero > idAtualNumero) {
              termosMap[chaveId][numeroChave] = termo;
            }
          }
        });
      }

      if (tipoNormalizado === 'admin' || tipoNormalizado === 'ambos' || tipoNormalizado === 'todos' || tipoNormalizado === 'geral') {
        var visitantes = getArmariosFromSheet('Visitantes', 'visitante', null);
        var acompanhantes = getArmariosFromSheet('Acompanhantes', 'acompanhante', termosMap);
        return { success: true, data: visitantes.concat(acompanhantes) };
      }

      var sheetName = tipoNormalizado === 'acompanhante' ? 'Acompanhantes' : 'Visitantes';
      var mapa = tipoNormalizado === 'acompanhante' ? termosMap : null;
      return { success: true, data: getArmariosFromSheet(sheetName, tipoNormalizado, mapa) };
    } catch (error) {
      registrarLog('ERRO', `Erro ao buscar armários: ${error.toString()}`);
      return { success: false, error: error.toString() };
    }
  });
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
  var nomeChaves = isVisitante ? CABECALHOS_NOME_VISITANTE : CABECALHOS_NOME_ACOMPANHANTE;
  var nomeIndex = obterIndiceColuna(estrutura, nomeChaves, 3);
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
  var whatsappIndex = obterIndiceColuna(estrutura, CABECALHOS_WHATSAPP, null);
  if (whatsappIndex === null || whatsappIndex === undefined) {
    whatsappIndex = isVisitante ? 12 : 9;
  }

  dados.forEach(function(row) {
    var idPlanilha = row[idIndex];
    if (!idPlanilha && idPlanilha !== 0) {
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

    var numeroBruto = row[numeroIndex] || '';
    var numeroNormalizado = normalizarNumeroArmario(numeroBruto);
    var idInterface = montarChaveArmarioInterface(tipo, numeroNormalizado, idPlanilha);

    var armario = {
      id: idInterface,
      idPlanilha: idPlanilha,
      numero: numeroNormalizado,
      status: status,
      nomeVisitante: obterValorLinha(row, estrutura, nomeChaves, row[nomeIndex] || ''),
      nomePaciente: row[pacienteIndex] || '',
      leito: row[leitoIndex] || '',
      volumes: row[volumesIndex] || 0,
      horaInicio: formatarHorarioPlanilha(row[horaInicioIndex]),
      tipo: tipo,
      unidade: unidadeIndex !== null && unidadeIndex !== undefined ? (row[unidadeIndex] || '') : '',
      termoAplicado: termoIndex !== null && termoIndex !== undefined ? converterParaBoolean(row[termoIndex]) : false,
      whatsapp: whatsappIndex !== null && whatsappIndex !== undefined ? (row[whatsappIndex] || '') : ''
    };

    if (isVisitante) {
      armario.horaPrevista = horaPrevistaIndex > -1 ? formatarHorarioPlanilha(row[horaPrevistaIndex]) : '';
      armario.dataRegistro = dataRegistroIndex > -1 ? formatarDataPlanilha(row[dataRegistroIndex]) : '';
    } else {
      armario.dataRegistro = dataRegistroIndex > -1 ? formatarDataPlanilha(row[dataRegistroIndex]) : '';
    }

    var volumesNumero = parseInt(armario.volumes, 10);
    armario.volumes = isNaN(volumesNumero) ? 0 : volumesNumero;

    if (tipo === 'acompanhante') {
      var termosPorId = null;
      if (termosMap) {
        var chaveId = idPlanilha !== null && idPlanilha !== undefined ? idPlanilha.toString().trim() : '';
        termosPorId = chaveId ? termosMap[chaveId] : null;
      }

      var termoRelacionado = null;
      if (termosPorId) {
        var chaveNumero = obterChaveNumeroArmario(numeroNormalizado);
        termoRelacionado = termosPorId[chaveNumero] || null;
        if (!termoRelacionado && chaveNumero !== '__sem_numero__') {
          termoRelacionado = termosPorId['__sem_numero__'] || null;
        }
      }
      if (termoRelacionado) {
        var statusTermoNormalizado = normalizarTextoBasico(termoRelacionado.status);
        var termoFinalizado = statusTermoNormalizado === 'finalizado';

        if (!termoFinalizado && (termoRelacionado.pdfUrl || (termoRelacionado.assinaturas && termoRelacionado.assinaturas.finalizadoEm))) {
          termoFinalizado = true;
          statusTermoNormalizado = 'finalizado';
        }

        var possuiTermo = Boolean(termoRelacionado);
        var termoEmAndamento = possuiTermo && !termoFinalizado;
        var statusDescricao = termoRelacionado.status || '';

        if (!statusDescricao) {
          statusDescricao = termoFinalizado ? 'Finalizado' : (possuiTermo ? 'Em andamento' : '');
        } else if (statusTermoNormalizado === 'finalizado') {
          statusDescricao = 'Finalizado';
        } else if (statusTermoNormalizado === 'em andamento') {
          statusDescricao = 'Em andamento';
        }

        var termoStatus = termoFinalizado ? 'finalizado' : (termoEmAndamento ? 'em andamento' : 'pendente');

        armario.termoAplicado = termoEmAndamento;
        armario.termoFinalizado = termoFinalizado;
        armario.termoStatus = termoStatus;
        armario.termoInfo = {
          id: termoRelacionado.id,
          aplicadoEm: termoRelacionado.aplicadoEm,
          finalizadoEm: termoRelacionado.assinaturas ? termoRelacionado.assinaturas.finalizadoEm : '',
          pdfUrl: termoRelacionado.pdfUrl || '',
          responsavel: termoRelacionado.acompanhante,
          metodoFinal: termoRelacionado.assinaturas ? termoRelacionado.assinaturas.metodoFinal : '',
          cpfFinal: termoRelacionado.assinaturas ? termoRelacionado.assinaturas.cpfFinal : '',
          status: statusDescricao
        };
      } else {
        armario.termoAplicado = false;
        armario.termoFinalizado = false;
        armario.termoStatus = 'pendente';
        armario.termoInfo = null;
      }
    } else {
      armario.termoFinalizado = false;
      armario.termoStatus = 'pendente';
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

    garantirEstruturaHistorico(historicoSheet);

    var totalLinhas = sheet.getLastRow();
    if (totalLinhas < 2) {
      return { success: false, error: 'Nenhum armário cadastrado' };
    }

    var estrutura = obterEstruturaPlanilha(sheet);
    var totalColunas = estrutura.ultimaColuna || (sheetName === 'Visitantes' ? 13 : 12);
    var linhaPlanilha = -1;
    var linhaAtual = null;
    var idParametroBruto = armarioData.idPlanilha !== undefined && armarioData.idPlanilha !== ''
      ? armarioData.idPlanilha
      : armarioData.id;
    var idIndex = obterIndiceColuna(estrutura, 'id', 0);
    var numeroIndex = obterIndiceColuna(estrutura, 'numero', 1);
    var statusIndex = obterIndiceColuna(estrutura, 'status', 2);

    if (totalLinhas > 1 && idParametroBruto !== undefined && idParametroBruto !== null && idParametroBruto !== '') {
      var idTexto = idParametroBruto.toString().trim();
      if (idTexto) {
        var intervaloId = sheet.getRange(2, idIndex + 1, totalLinhas - 1, 1);
        var idFinder = intervaloId.createTextFinder(idTexto).matchEntireCell(true);
        var idEncontrado = idFinder ? idFinder.findNext() : null;
        if (idEncontrado) {
          linhaPlanilha = idEncontrado.getRow();
          linhaAtual = sheet.getRange(linhaPlanilha, 1, 1, totalColunas).getValues()[0];
        }
      }
    }

    if ((linhaPlanilha === -1 || !linhaAtual) && armarioData.numero && totalLinhas > 1) {
      var numeroInformado = armarioData.numero.toString().trim();
      if (numeroInformado) {
        var intervaloNumero = sheet.getRange(2, numeroIndex + 1, totalLinhas - 1, 1);
        var numeroFinder = intervaloNumero.createTextFinder(numeroInformado).matchEntireCell(true);
        var correspondencias = numeroFinder ? numeroFinder.findAll() : [];
        for (var j = 0; j < correspondencias.length; j++) {
          var linhaCandidata = correspondencias[j].getRow();
          var valoresLinha = sheet.getRange(linhaCandidata, 1, 1, totalColunas).getValues()[0];
          var statusLinha = normalizarTextoBasico(
            obterValorLinha(valoresLinha, estrutura, 'status', valoresLinha[statusIndex])
          );
          if (statusLinha === 'livre') {
            linhaPlanilha = linhaCandidata;
            linhaAtual = valoresLinha;
            break;
          }
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

    var dataHoraAtual = obterDataHoraAtualFormatada();
    var responsavelRegistro = determinarResponsavelRegistro(armarioData.usuarioResponsavel);
    var horaInicio = dataHoraAtual.horaCurta;
    var dataRegistro = dataHoraAtual.dataHoraIso;
    var volumes = parseInt(armarioData.volumes, 10);
    if (isNaN(volumes) || volumes < 0) {
      volumes = 0;
    }
    var whatsapp = armarioData.whatsapp !== null && armarioData.whatsapp !== undefined
      ? armarioData.whatsapp.toString().trim()
      : '';
    var numeroArmario = linhaAtual[numeroIndex];
    var unidadeAtual = obterValorLinha(linhaAtual, estrutura, 'unidade', '');
    var novaLinha = linhaAtual.slice();
    while (novaLinha.length < totalColunas) {
      novaLinha.push('');
    }
    var nomeChavesCadastro = sheetName === 'Visitantes' ? CABECALHOS_NOME_VISITANTE : CABECALHOS_NOME_ACOMPANHANTE;

    definirValorLinha(novaLinha, estrutura, 'status', 'em-uso');
    definirValorLinha(novaLinha, estrutura, nomeChavesCadastro, armarioData.nomeVisitante);
    definirValorLinha(novaLinha, estrutura, 'nome paciente', armarioData.nomePaciente);
    definirValorLinha(novaLinha, estrutura, 'leito', armarioData.leito);
    definirValorLinha(novaLinha, estrutura, 'volumes', volumes);
    definirValorLinha(novaLinha, estrutura, 'hora inicio', horaInicio);
    if (sheetName === 'Visitantes') {
      definirValorLinha(novaLinha, estrutura, 'hora prevista', armarioData.horaPrevista || '');
    } else {
      definirValorLinha(novaLinha, estrutura, 'hora prevista', '');
    }
    definirValorLinha(novaLinha, estrutura, 'data registro', dataRegistro);
    definirValorLinha(novaLinha, estrutura, 'unidade', unidadeAtual);
    definirValorLinha(novaLinha, estrutura, CABECALHOS_WHATSAPP, whatsapp);
    definirValorLinha(novaLinha, estrutura, 'termo aplicado', false);

    sheet.getRange(linhaPlanilha, 1, 1, totalColunas).setValues([novaLinha]);

    var historicoLastRow = historicoSheet.getLastRow();
    var ultimoHistoricoId = historicoLastRow > 1
      ? Number(historicoSheet.getRange(historicoLastRow, 1).getValue()) || 0
      : 0;
    var historicoId = ultimoHistoricoId + 1;
    var proximaLinhaHistorico = historicoLastRow + 1;

    var dataHistorico = dataHoraAtual.data;

    var historicoLinha = [
      historicoId,
      dataHistorico,
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
      whatsapp,
      responsavelRegistro
    ];

    historicoSheet.getRange(proximaLinhaHistorico, 1, 1, historicoLinha.length).setValues([historicoLinha]);

    registrarLog('CADASTRO', `Armário ${numeroArmario} cadastrado para ${armarioData.nomeVisitante}`);

    limparCacheArmarios();
    limparCacheHistorico();

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

function liberarArmario(id, tipo, numero, usuarioResponsavel) {
  try {
    var tipoNormalizado = normalizarTextoBasico(tipo);
    var ehAcompanhante = tipoNormalizado === 'acompanhante';
    var idComparacao = id !== null && id !== undefined ? id.toString().trim() : '';
    var numeroInformado = normalizarNumeroArmario(numero);

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = ehAcompanhante ? 'Acompanhantes' : 'Visitantes';
    var sheet = ss.getSheetByName(sheetName);
    var historicoSheet = ss.getSheetByName(
      ehAcompanhante ? 'Histórico Acompanhantes' : 'Histórico Visitantes'
    );

    if (!sheet || !historicoSheet) {
      return { success: false, error: 'Abas não encontradas' };
    }

    garantirEstruturaHistorico(historicoSheet);

    var responsavelRegistro = determinarResponsavelRegistro(usuarioResponsavel);

    // Encontrar o armário na aba atual
    var estrutura = obterEstruturaPlanilha(sheet);
    var totalColunas = estrutura.ultimaColuna || (sheetName === 'Visitantes' ? 13 : 12);
    var totalLinhas = sheet.getLastRow();
    if (totalLinhas <= 1) {
      return { success: false, error: 'Nenhum armário cadastrado' };
    }

    var idIndex = obterIndiceColuna(estrutura, 'id', 0);
    var statusIndex = obterIndiceColuna(estrutura, 'status', 2);
    var numeroIndex = obterIndiceColuna(estrutura, 'numero', 1);
    var linhaPlanilha = -1;
    var armarioData = null;

    if (totalLinhas > 1 && idComparacao) {
      var intervaloId = sheet.getRange(2, idIndex + 1, totalLinhas - 1, 1);
      var idFinder = intervaloId.createTextFinder(idComparacao).matchEntireCell(true);
      var idEncontrado = idFinder ? idFinder.findNext() : null;
      if (idEncontrado) {
        linhaPlanilha = idEncontrado.getRow();
        armarioData = sheet.getRange(linhaPlanilha, 1, 1, totalColunas).getValues()[0];
      }
    }

    if ((linhaPlanilha === -1 || !armarioData) && numeroInformado && totalLinhas > 1) {
      var intervaloNumero = sheet.getRange(2, numeroIndex + 1, totalLinhas - 1, 1);
      var numeroFinder = intervaloNumero.createTextFinder(numeroInformado).matchEntireCell(true);
      var correspondencias = numeroFinder ? numeroFinder.findAll() : [];
      for (var indice = 0; indice < correspondencias.length; indice++) {
        var linhaCandidata = correspondencias[indice].getRow();
        var valoresLinha = sheet.getRange(linhaCandidata, 1, 1, totalColunas).getValues()[0];
        linhaPlanilha = linhaCandidata;
        armarioData = valoresLinha;
        break;
      }
    }

    if (linhaPlanilha === -1 || !armarioData) {
      return { success: false, error: 'Armário não encontrado' };
    }

    var statusPadrao = (statusIndex !== null && statusIndex !== undefined && statusIndex < armarioData.length)
      ? armarioData[statusIndex]
      : '';
    var statusAtual = normalizarTextoBasico(
      obterValorLinha(armarioData, estrutura, 'status', statusPadrao)
    );
    if (statusAtual === 'livre') {
      return { success: false, error: 'Armário já está livre' };
    }

    // Limpar dados do armário (deixar apenas número e status livre)
    var unidadeAtual = obterValorLinha(armarioData, estrutura, 'unidade', '');
    var novaLinha = armarioData.slice();
    while (novaLinha.length < totalColunas) {
      novaLinha.push('');
    }

    var nomeColuna = sheetName === 'Visitantes' ? CABECALHOS_NOME_VISITANTE : CABECALHOS_NOME_ACOMPANHANTE;
    definirValorLinha(novaLinha, estrutura, 'status', 'livre');
    definirValorLinha(novaLinha, estrutura, nomeColuna, '');
    definirValorLinha(novaLinha, estrutura, 'nome paciente', '');
    definirValorLinha(novaLinha, estrutura, 'leito', '');
    definirValorLinha(novaLinha, estrutura, 'volumes', '');
    definirValorLinha(novaLinha, estrutura, 'hora inicio', '');
    if (sheetName === 'Visitantes') {
      definirValorLinha(novaLinha, estrutura, 'hora prevista', '');
    }
    var dataHoraAtual = obterDataHoraAtualFormatada();
    definirValorLinha(novaLinha, estrutura, 'data registro', dataHoraAtual.dataHoraIso);
    definirValorLinha(novaLinha, estrutura, CABECALHOS_WHATSAPP, '');
    definirValorLinha(novaLinha, estrutura, 'unidade', unidadeAtual);
    definirValorLinha(novaLinha, estrutura, 'termo aplicado', false);

    sheet.getRange(linhaPlanilha, 1, 1, totalColunas).setValues([novaLinha]);

    // Atualizar histórico - encontrar a entrada mais recente deste armário
    var historicoLastRow = historicoSheet.getLastRow();
    var numeroArmario = obterValorLinha(armarioData, estrutura, 'numero', armarioData[numeroIndex]);
    numeroArmario = numeroArmario ? numeroArmario.toString().trim() : '';
    if (!numeroArmario) {
      numeroArmario = numeroInformado;
    }

    if (historicoLastRow > 1 && numeroArmario) {
      var intervaloHistoricoNumeros = historicoSheet.getRange(2, 3, historicoLastRow - 1, 1);
      var historicoFinder = intervaloHistoricoNumeros.createTextFinder(numeroArmario).matchEntireCell(true);
      var ocorrencias = historicoFinder ? historicoFinder.findAll() : [];
      for (var i = ocorrencias.length - 1; i >= 0; i--) {
        var linhaHistorico = ocorrencias[i].getRow();
        var statusRegistro = historicoSheet.getRange(linhaHistorico, 10).getValue();
        if (statusRegistro === 'EM USO') {
          var horaFim = dataHoraAtual.horaCurta;
          historicoSheet.getRange(linhaHistorico, 9).setValue(horaFim);
          historicoSheet.getRange(linhaHistorico, 10).setValue('FINALIZADO');
          if (responsavelRegistro) {
            historicoSheet.getRange(linhaHistorico, 14).setValue(responsavelRegistro);
          }
          break;
        }
      }
    }

    registrarLog('LIBERAÇÃO', `Armário ${numeroArmario} liberado`);

    limparCacheArmarios();
    limparCacheHistorico();

    return { success: true, message: 'Armário liberado com sucesso' };

  } catch (error) {
    registrarLog('ERRO', `Erro ao liberar armário: ${error.toString()}`);
    return { success: false, error: error.toString() };
  }
}

// Funções para Usuários
function getUsuarios() {
  return executarComCache(montarChaveCache('usuarios'), CACHE_TTL_PADRAO, function() {
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
  });
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
    var dataCadastro = obterDataHoraAtualFormatada().dataHoraIso;

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

    limparCacheUsuarios();

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

    limparCacheUsuarios();

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

    limparCacheUsuarios();

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
    if (senhaArmazenada !== null && senhaArmazenada !== undefined) {
      senhaArmazenada = senhaArmazenada.toString().trim();
    } else {
      senhaArmazenada = '';
    }

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
      matricula: obterValorLinha(linhaUsuario, estrutura, 'matricula', ''),
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
  var tipoNormalizado = normalizarTextoBasico(tipo) === 'acompanhante' ? 'acompanhante' : 'visitante';
  var chaveCache = montarChaveCache('historico', tipoNormalizado);

  return executarComCache(chaveCache, CACHE_TTL_HISTORICO, function() {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheetName = tipoNormalizado === 'acompanhante' ? 'Histórico Acompanhantes' : 'Histórico Visitantes';
      var sheet = ss.getSheetByName(sheetName);

      if (!sheet || sheet.getLastRow() < 2) {
        return { success: true, data: [] };
      }

      garantirEstruturaHistorico(sheet);

      var totalLinhasDados = sheet.getLastRow() - 1;
      var totalColunasDados = Math.max(sheet.getLastColumn(), 14);
      var data = sheet.getRange(2, 1, totalLinhasDados, totalColunasDados).getValues();
      var historico = [];

      data.forEach(function(row) {
        if (row[0]) {
          historico.push({
            id: row[0],
            data: formatarDataPlanilha(row[1]),
            armario: row[2],
            nome: row[3],
            paciente: row[4],
            leito: row[5],
            volumes: row[6],
            horaInicio: formatarHorarioPlanilha(row[7]),
            horaFim: formatarHorarioPlanilha(row[8]),
            status: row[9],
            tipo: row[10],
            unidade: row[11],
            whatsapp: row[12] || '',
            usuario: row[13] || ''
          });
        }
      });

      return { success: true, data: historico.reverse() }; // Mais recentes primeiro

    } catch (error) {
      registrarLog('ERRO', `Erro ao buscar histórico: ${error.toString()}`);
      return { success: false, error: error.toString() };
    }
  });
}

// Funções para Cadastro de Armários Físicos
function getCadastroArmarios() {
  return executarComCache(montarChaveCache('cadastro-armarios'), CACHE_TTL_PADRAO, function() {
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
  });
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

    var dataCadastro = obterDataHoraAtualFormatada().dataHoraIso;
    var linhas = novosArmarios.map(function(numero, index) {
      return [
        ultimoId + index + 1,
        numero,
        armarioData.tipo,
        armarioData.unidade,
        armarioData.localizacao,
        'ativo',
        dataCadastro
      ];
    });

    if (linhas.length > 0) {
      sheet.getRange(lastRow + 1, 1, linhas.length, 7).setValues(linhas);

      // Também criar nas abas de uso
      criarArmariosUso(linhas);
    }

    registrarLog('CADASTRO', `Armários físicos cadastrados: ${novosArmarios.join(', ')}`);

    limparCacheCadastroArmarios();
    limparCacheArmarios();

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
        var dataRegistro = obterDataHoraAtualFormatada().dataHoraIso;

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
            dataRegistro, // data registro
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
            dataRegistro, // data registro
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
  return executarComCache(montarChaveCache('unidades'), CACHE_TTL_PADRAO, function() {
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
  });
}

function getSetores() {
  return executarComCache(montarChaveCache('setores'), CACHE_TTL_PADRAO, function() {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName('Cadastro');

      if (!sheet) {
        return { success: true, data: [] };
      }

      var ultimaLinha = sheet.getLastRow();
      if (ultimaLinha < 2) {
        return { success: true, data: [] };
      }

      var valores = sheet.getRange(2, 1, ultimaLinha - 1, 1).getValues();
      var setoresMapeados = {};
      var setores = [];

      for (var i = 0; i < valores.length; i++) {
        var bruto = valores[i][0];
        if (bruto === null || bruto === undefined) {
          continue;
        }
        var texto = bruto.toString().trim();
        if (!texto) {
          continue;
        }
        var chave = normalizarTextoBasico(texto);
        if (!chave) {
          continue;
        }
        if (setoresMapeados[chave]) {
          continue;
        }
        setoresMapeados[chave] = true;
        setores.push(texto);
      }

      setores.sort(function(a, b) {
        return a.localeCompare(b, 'pt-BR', { sensitivity: 'base' });
      });

      return { success: true, data: setores };

    } catch (error) {
      registrarLog('ERRO', 'Erro ao buscar setores: ' + error.toString());
      return { success: false, error: error.toString() };
    }
  });
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
    
    var dataCadastro = obterDataHoraAtualFormatada().dataHoraIso;

    var novaLinha = [
      novoId,
      dados.nome,
      'ativa',
      dataCadastro
    ];
    
    sheet.getRange(lastRow + 1, 1, 1, 4).setValues([novaLinha]);

    registrarLog('CADASTRO UNIDADE', `Unidade ${dados.nome} cadastrada`);

    limparCacheUnidades();

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

    limparCacheUnidades();

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
    var numeroInformado = normalizarNumeroArmario(dadosTermo.numeroArmario);

    // 1. Salvar na aba "Termos de Responsabilidade"
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Termos de Responsabilidade');

    if (!sheet) {
      throw new Error('Aba "Termos de Responsabilidade" não encontrada');
    }

    var dadosExistentes = sheet.getDataRange().getValues();
    var linhaExistente = -1;
    var termoId = null;
    var aplicadoEmAtual = obterDataHoraAtualFormatada().dataHoraIso;
    var aplicadoEm = aplicadoEmAtual;

    for (var i = dadosExistentes.length - 1; i >= 1; i--) {
      var idLinha = dadosExistentes[i][1];
      if (String(idLinha) !== String(dadosTermo.armarioId)) {
        continue;
      }

      var numeroLinha = dadosExistentes[i][2] ? dadosExistentes[i][2].toString().trim() : '';
      if (numeroInformado && normalizarNumeroArmario(numeroLinha) !== numeroInformado) {
        continue;
      }

      var assinaturasExistentes = normalizarAssinaturas(dadosExistentes[i][18]);
      var statusLinha = normalizarTextoBasico(dadosExistentes[i][19]);
      var finalizado = Boolean(dadosExistentes[i][17] || statusLinha === 'finalizado' || (assinaturasExistentes && assinaturasExistentes.finalizadoEm));

      if (!finalizado) {
        linhaExistente = i + 1;
        termoId = dadosExistentes[i][0];
        aplicadoEm = converterParaDataHoraIso(dadosExistentes[i][16], aplicadoEmAtual);
        break;
      }
    }

    if (linhaExistente === -1) {
      var lastRow = sheet.getLastRow();
      termoId = lastRow > 1 ? Math.max.apply(null, sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat()) + 1 : 1;
      linhaExistente = lastRow + 1;
    }

    var valorAtualAssinatura = '';
    var statusTermo = 'Em andamento';
    if (linhaExistente <= dadosExistentes.length && linhaExistente - 1 >= 0) {
      var linhaAtual = dadosExistentes[linhaExistente - 1];
      if (linhaAtual) {
        if (linhaAtual.length > 18) {
          valorAtualAssinatura = linhaAtual[18];
        }
        if (linhaAtual.length > 19) {
          var statusExistente = linhaAtual[19];
          if (statusExistente && statusExistente.toString().trim()) {
            statusTermo = statusExistente;
          }
        }
      }
    }

    if (normalizarTextoBasico(statusTermo) !== 'finalizado') {
      statusTermo = 'Em andamento';
    }

    var assinaturasInfo = normalizarAssinaturas(valorAtualAssinatura);
    assinaturasInfo.inicial = dadosTermo.assinaturaBase64 || assinaturasInfo.inicial || '';

    var numeroArmarioOficial = numeroInformado;

    var sheetAcompanhantes = ss.getSheetByName('Acompanhantes');
    var dadosAcompanhantes = [];
    var estruturaAcompanhantes = null;
    var linhaAcompanhante = -1;

    if (sheetAcompanhantes) {
      estruturaAcompanhantes = obterEstruturaPlanilha(sheetAcompanhantes);
      dadosAcompanhantes = sheetAcompanhantes.getDataRange().getValues();
      for (var indiceA = 1; indiceA < dadosAcompanhantes.length; indiceA++) {
        var linha = dadosAcompanhantes[indiceA];
        if (linha && linha[0] == dadosTermo.armarioId) {
          linhaAcompanhante = indiceA;
          if (linha.length > 1 && linha[1]) {
            numeroArmarioOficial = linha[1];
          } else if (!numeroArmarioOficial) {
            numeroArmarioOficial = dadosTermo.armarioId;
          }
          break;
        }
      }
    }

    if (linhaAcompanhante === -1) {
      var sheetVisitantes = ss.getSheetByName('Visitantes');
      if (sheetVisitantes) {
        var dadosVisitantes = sheetVisitantes.getDataRange().getValues();
        for (var indiceV = 1; indiceV < dadosVisitantes.length; indiceV++) {
          var linhaVisitante = dadosVisitantes[indiceV];
          if (linhaVisitante && linhaVisitante[0] == dadosTermo.armarioId) {
            if (linhaVisitante.length > 1 && linhaVisitante[1]) {
              numeroArmarioOficial = linhaVisitante[1];
            } else if (!numeroArmarioOficial) {
              numeroArmarioOficial = dadosTermo.armarioId;
            }
            break;
          }
        }
      }
    }

    if (!numeroArmarioOficial) {
      numeroArmarioOficial = dadosTermo.armarioId || '';
    }

    dadosTermo.numeroArmario = numeroArmarioOficial;

    var linhaDados = [
      termoId,
      dadosTermo.armarioId,
      numeroArmarioOficial,
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
      JSON.stringify(assinaturasInfo),
      statusTermo
    ];

    sheet.getRange(linhaExistente, 1, 1, linhaDados.length).setValues([linhaDados]);

    // 2. Atualizar status do armário na aba "Acompanhantes"
    var cadastroArmario = dadosTermo.cadastroArmario;
    if (typeof cadastroArmario === 'string' && cadastroArmario) {
      try {
        cadastroArmario = JSON.parse(cadastroArmario);
      } catch (erroCadastro) {
        cadastroArmario = null;
      }
    }

    var precisaAtualizarCadastro = cadastroArmario && String(cadastroArmario.id) === String(dadosTermo.armarioId);

    if (linhaAcompanhante > -1 && sheetAcompanhantes && estruturaAcompanhantes) {
      var totalColunasAcompanhantes = estruturaAcompanhantes.ultimaColuna || 12;
      var linhaAtualizada = dadosAcompanhantes[linhaAcompanhante] ? dadosAcompanhantes[linhaAcompanhante].slice() : [];

      while (linhaAtualizada.length < totalColunasAcompanhantes) {
        linhaAtualizada.push('');
      }

      definirValorLinha(linhaAtualizada, estruturaAcompanhantes, 'volumes', totalVolumes);
      definirValorLinha(linhaAtualizada, estruturaAcompanhantes, 'termo aplicado', true);

      if (precisaAtualizarCadastro) {
        var statusAtual = normalizarTextoBasico(obterValorLinha(linhaAtualizada, estruturaAcompanhantes, 'status', ''));
        if (statusAtual && statusAtual !== 'livre') {
          throw new Error('Armário já está em uso. Atualize a lista e tente novamente.');
        }

        var dataHoraAtualCadastro = obterDataHoraAtualFormatada();
        var horaInicioCadastro = dataHoraAtualCadastro.horaCurta;
        var dataRegistroCadastro = dataHoraAtualCadastro.dataHoraIso;
        var unidadeAtual = obterValorLinha(linhaAtualizada, estruturaAcompanhantes, 'unidade', '');
        var whatsappCadastro = cadastroArmario.whatsapp ? cadastroArmario.whatsapp.toString().trim() : '';
        var nomeColunaCadastro = CABECALHOS_NOME_ACOMPANHANTE;

        definirValorLinha(linhaAtualizada, estruturaAcompanhantes, 'status', 'em-uso');
        definirValorLinha(linhaAtualizada, estruturaAcompanhantes, nomeColunaCadastro, cadastroArmario.nomeVisitante || dadosTermo.acompanhante || '');
        definirValorLinha(linhaAtualizada, estruturaAcompanhantes, 'nome paciente', cadastroArmario.nomePaciente || dadosTermo.paciente || '');
        definirValorLinha(linhaAtualizada, estruturaAcompanhantes, 'leito', cadastroArmario.leito || dadosTermo.leito || '');
        definirValorLinha(linhaAtualizada, estruturaAcompanhantes, 'hora inicio', horaInicioCadastro);
        definirValorLinha(linhaAtualizada, estruturaAcompanhantes, 'hora prevista', '');
        definirValorLinha(linhaAtualizada, estruturaAcompanhantes, 'data registro', dataRegistroCadastro);
        definirValorLinha(linhaAtualizada, estruturaAcompanhantes, CABECALHOS_WHATSAPP, whatsappCadastro);

        // Registrar histórico de uso
        var historicoSheet = ss.getSheetByName('Histórico Acompanhantes');
        if (historicoSheet) {
          var historicoLastRow = historicoSheet.getLastRow();
          var historicoId = historicoLastRow > 1
            ? Math.max.apply(null, historicoSheet.getRange(2, 1, historicoLastRow - 1, 1).getValues().flat()) + 1
            : 1;

          var historicoLinha = [
            historicoId,
            dataRegistroCadastro,
            numeroArmarioOficial,
            cadastroArmario.nomeVisitante || dadosTermo.acompanhante || '',
            cadastroArmario.nomePaciente || dadosTermo.paciente || '',
            cadastroArmario.leito || dadosTermo.leito || '',
            totalVolumes,
            horaInicioCadastro,
            '',
            'EM USO',
            'acompanhante',
            unidadeAtual,
            whatsappCadastro
          ];

          historicoSheet.getRange(historicoLastRow + 1, 1, 1, historicoLinha.length).setValues([historicoLinha]);
        }
      }

      sheetAcompanhantes.getRange(linhaAcompanhante + 1, 1, 1, totalColunasAcompanhantes).setValues([linhaAtualizada]);
    }

    limparCacheTermos();
    limparCacheArmarios();

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
    finalizadoEm: '',
    responsavelFinalizacao: ''
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
      info.responsavelFinalizacao = json.responsavelFinalizacao || '';
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
    info.responsavelFinalizacao = valor.responsavelFinalizacao || '';
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

    var numeroArmarioValor = data[i][2] ? data[i][2].toString().trim() : '';
    var statusBruto = data[i][19] || '';
    var statusNormalizado = normalizarTextoBasico(statusBruto);
    var possuiPdf = Boolean(data[i][17]);
    var finalizado = Boolean(possuiPdf || statusNormalizado === 'finalizado' || (assinaturas && assinaturas.finalizadoEm));

    termos.push({
      linha: i + 1,
      id: data[i][0],
      armarioId: data[i][1],
      numeroArmario: numeroArmarioValor,
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
      assinaturas: assinaturas,
      status: statusBruto,
      statusNormalizado: statusNormalizado,
      finalizado: finalizado,
      possuiPdf: possuiPdf
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
    var termoFinalizadoMaisRecente = null;
    var armarioIdInformado = dados.armarioId !== null && dados.armarioId !== undefined
      ? dados.armarioId.toString().trim()
      : '';
    var numeroInformado = normalizarNumeroArmario(dados.numeroArmario);
    var incluirFinalizados = converterParaBoolean(dados.incluirFinalizados);

    for (var i = data.length - 1; i >= 1; i--) {
      var idLinha = data[i][1];
      var idLinhaTexto = idLinha !== null && idLinha !== undefined ? idLinha.toString().trim() : '';
      if (armarioIdInformado && idLinhaTexto !== armarioIdInformado) {
        continue;
      }

      var numeroLinha = data[i][2] ? data[i][2].toString().trim() : '';
      var numeroLinhaNormalizado = normalizarNumeroArmario(numeroLinha);
      if (numeroInformado && numeroLinhaNormalizado !== numeroInformado) {
        continue;
      }

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

      var statusBruto = data[i][19] || '';
      var statusNormalizado = normalizarTextoBasico(statusBruto);
      var possuiPdf = Boolean(data[i][17]);
      var termoFinalizado = Boolean(possuiPdf || statusNormalizado === 'finalizado' || (assinaturas && assinaturas.finalizadoEm));

      var termoAtual = {
        id: data[i][0],
        armarioId: data[i][1],
        numeroArmario: numeroLinha,
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
        cpfConfirmacao: assinaturas.cpfFinal,
        status: statusBruto,
        statusNormalizado: statusNormalizado,
        finalizado: termoFinalizado
      };

      if (termoFinalizado) {
        if (incluirFinalizados && !termoFinalizadoMaisRecente) {
          termoFinalizadoMaisRecente = termoAtual;
        }
        if (!incluirFinalizados) {
          continue;
        }
      }

      termo = termoAtual;
      if (!termoFinalizado) {
        break;
      }
    }

    if (!termo && incluirFinalizados && termoFinalizadoMaisRecente) {
      termo = termoFinalizadoMaisRecente;
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
    var numeroInformado = normalizarNumeroArmario(dados.numeroArmario);
    var tipoTermo = dados && dados.tipo ? dados.tipo.toString() : '';

    var termosInfo = obterTermosRegistrados();
    if (!termosInfo.sheet) {
      return { success: false, error: 'Aba "Termos de Responsabilidade" não encontrada' };
    }

    var termoEncontrado = null;
    var termoFinalizadoMaisRecente = null;
    for (var i = termosInfo.termos.length - 1; i >= 0; i--) {
      var termoAtual = termosInfo.termos[i];
      if (!termoAtual) {
        continue;
      }

      if (termoAtual.armarioId != armarioId) {
        continue;
      }

      var numeroTermo = normalizarNumeroArmario(termoAtual.numeroArmario);
      if (numeroInformado && numeroTermo !== numeroInformado) {
        continue;
      }

      var statusNormalizado = normalizarTextoBasico(termoAtual.status || termoAtual.statusNormalizado || '');
      var finalizado = termoAtual.finalizado;
      if (finalizado === undefined) {
        finalizado = Boolean(termoAtual.pdfUrl || (termoAtual.assinaturas && termoAtual.assinaturas.finalizadoEm) || statusNormalizado === 'finalizado');
      }

      if (finalizado) {
        if (!termoFinalizadoMaisRecente) {
          termoFinalizadoMaisRecente = termoAtual;
        }
        continue;
      }

      if (!termoAtual.pdfUrl) {
        termoEncontrado = termoAtual;
        break;
      }
    }

    if (!termoEncontrado) {
      for (var j = termosInfo.termos.length - 1; j >= 0; j--) {
        var termoAtualBusca = termosInfo.termos[j];
        if (!termoAtualBusca) {
          continue;
        }

        if (termoAtualBusca.armarioId != armarioId) {
          continue;
        }

        var numeroTermoBusca = normalizarNumeroArmario(termoAtualBusca.numeroArmario);
        if (numeroInformado && numeroTermoBusca !== numeroInformado) {
          continue;
        }

        termoEncontrado = termoAtualBusca;
        break;
      }
    }

    if (!termoEncontrado && termoFinalizadoMaisRecente) {
      termoEncontrado = termoFinalizadoMaisRecente;
    }

    if (!termoEncontrado) {
      return { success: false, error: 'Termo não localizado para este armário' };
    }

    var assinaturas = termoEncontrado.assinaturas || normalizarAssinaturas('');
    var finalizacaoInfo = obterDataHoraAtualFormatada();
    var finalizacaoIso = finalizacaoInfo.dataHoraIso;
    assinaturas.metodoFinal = metodo;
    assinaturas.cpfFinal = metodo === 'cpf' ? confirmacao : '';
    assinaturas.finalizadoEm = finalizacaoIso;
    assinaturas.final = metodo === 'assinatura' ? assinaturaFinal : '';
    var responsavelFinalizacao = determinarResponsavelRegistro(dados.usuarioResponsavel);
    assinaturas.responsavelFinalizacao = responsavelFinalizacao;

    var movimentacoesResultado = getMovimentacoes({ armarioId: armarioId, numeroArmario: numeroInformado });
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
      finalizadoEm: finalizacaoIso,
      assinaturaInicial: assinaturas.inicial,
      assinaturaFinal: assinaturas.final,
      metodoFinal: assinaturas.metodoFinal,
      cpfFinal: assinaturas.cpfFinal,
      responsavelFinalizacao: assinaturas.responsavelFinalizacao,
      movimentacoes: movimentacoes
    };

    var resultadoPDF = gerarESalvarTermoPDF(dadosPDF);
    if (!resultadoPDF.success) {
      throw new Error(resultadoPDF.error || 'Falha ao gerar PDF');
    }

    termosInfo.sheet.getRange(termoEncontrado.linha, 18).setValue(resultadoPDF.pdfUrl);
    termosInfo.sheet.getRange(termoEncontrado.linha, 19).setValue(JSON.stringify(assinaturas));
    termosInfo.sheet.getRange(termoEncontrado.linha, 20).setValue('Finalizado');

    termoEncontrado.status = 'Finalizado';

    finalizarMovimentacoesArmario(armarioId, numeroInformado, tipoTermo);

    limparCacheTermos();
    limparCacheArmarios();

    registrarLog('TERMO_FINALIZADO', 'Termo finalizado para armário ' + termoEncontrado.numeroArmario);

    return {
      success: true,
      pdfUrl: resultadoPDF.pdfUrl,
      finalizadoEm: finalizacaoIso
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

  function formatarDataHoraCompleta(data) {
    if (!data) return 'Não informada';
    try {
      var date = new Date(data);
      if (isNaN(date.getTime())) {
        return data;
      }
      return Utilities.formatDate(date, 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm');
    } catch (erro) {
      return data;
    }
  }

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
  var dataDevolucaoTexto = dadosTermo.finalizadoEm
    ? formatarDataHoraCompleta(dadosTermo.finalizadoEm)
    : '__________________________';
  var conferenteTexto = (dadosTermo.responsavelFinalizacao || '').toString().trim();
  if (!conferenteTexto) {
    conferenteTexto = '__________________________';
  }

  var assinaturaInicialHtml = dadosTermo.assinaturaInicial
    ? '<img src="data:image/png;base64,' + dadosTermo.assinaturaInicial + '" class="assinatura-img" alt="Assinatura inicial" />'
    : '<div class="assinatura-linha">Assinatura não registrada digitalmente.</div>';

  var assinaturaFinalHtml = '';
  if (dadosTermo.metodoFinal === 'cpf' && dadosTermo.cpfFinal) {
    assinaturaFinalHtml = '<div class="assinatura-linha">Confirmação por CPF: ' + dadosTermo.cpfFinal + '</div>';
  } else if (dadosTermo.metodoFinal === 'manual') {
    assinaturaFinalHtml = '<div class="assinatura-linha">Finalização manual registrada no sistema.</div>';
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
  partes.push('<div class="devolucao-box">Data: ' + dataDevolucaoTexto + ' &nbsp;&nbsp; Conferente: ' + conferenteTexto + '</div>');
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
  partes.push('  <div style="margin-top:4px; font-size:11px;">Encerrado em: ' + formatarDataHoraCompleta(dadosTermo.finalizadoEm) + '</div>');
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
function garantirEstruturaMovimentacoes(sheet) {
  if (!sheet) {
    return 10;
  }
  var colunaStatus = 10;
  var totalColunas = sheet.getLastColumn();
  if (totalColunas < colunaStatus) {
    sheet.insertColumnsAfter(totalColunas, colunaStatus - totalColunas);
  }
  var cabecalhoStatus = sheet.getRange(1, colunaStatus).getValue();
  if (!cabecalhoStatus) {
    sheet.getRange(1, colunaStatus).setValue('Status');
  }
  return colunaStatus;
}

function garantirEstruturaHistorico(sheet) {
  if (!sheet) {
    return 13;
  }
  var minimoColunas = 14;
  var totalColunas = sheet.getLastColumn();
  if (totalColunas < minimoColunas) {
    sheet.insertColumnsAfter(totalColunas, minimoColunas - totalColunas);
    totalColunas = sheet.getLastColumn();
  }
  var cabecalhos = sheet.getRange(1, 1, 1, Math.max(totalColunas, minimoColunas)).getValues()[0];
  if (!cabecalhos[13]) {
    sheet.getRange(1, 14).setValue('Usuário');
  }
  return Math.max(totalColunas, minimoColunas);
}

function getMovimentacoes(dados) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Movimentações');
  var colunaStatus = garantirEstruturaMovimentacoes(sheet);
  var possuiArmario = dados && dados.armarioId !== undefined && dados.armarioId !== null;
  var armarioId = possuiArmario ? dados.armarioId : null;
  var armarioIdTexto = possuiArmario && armarioId !== null && armarioId !== undefined ? armarioId.toString().trim() : '';
  var numeroInformado = normalizarNumeroArmario(dados ? dados.numeroArmario : '');
  var tipoInformado = dados && dados.tipo ? normalizarTextoBasico(dados.tipo) : '';
  var chaveIdentificacao = possuiArmario ? [armarioIdTexto, numeroInformado, tipoInformado].join('|') : 'todos';
  var chaveCache = montarChaveCache('movimentacoes', chaveIdentificacao);

  return executarComCache(chaveCache, CACHE_TTL_MOVIMENTACOES, function() {
    try {
      if (!sheet || sheet.getLastRow() < 2) {
        return { success: true, data: [] };
      }

      var data = sheet.getDataRange().getValues();
      var movimentacoes = [];

      if (!possuiArmario) {
        return { success: true, data: movimentacoes };
      }

      for (var i = 1; i < data.length; i++) {
        var idLinha = data[i][1];
        if (armarioIdTexto && String(idLinha) !== armarioIdTexto) {
          continue;
        }

        if (numeroInformado) {
          var numeroLinha = data[i][2] ? data[i][2].toString().trim() : '';
          if (normalizarNumeroArmario(numeroLinha) !== numeroInformado) {
            continue;
          }
        }

        var statusLinha = colunaStatus ? data[i][colunaStatus - 1] : '';
        var statusNormalizado = normalizarTextoBasico(statusLinha);
        if (statusNormalizado === 'finalizado') {
          continue;
        }

        movimentacoes.push({
          id: data[i][0],
          armarioId: data[i][1],
          numeroArmario: data[i][2],
          tipo: data[i][3],
          descricao: data[i][4],
          responsavel: data[i][5],
          data: formatarDataPlanilha(data[i][6]),
          hora: formatarHorarioPlanilha(data[i][7]),
          dataHoraRegistro: converterParaDataHoraIso(data[i][8], ''),
          status: statusLinha || ''
        });
      }

      return { success: true, data: movimentacoes };

    } catch (error) {
      registrarLog('ERRO', `Erro ao buscar movimentações: ${error.toString()}`);
      return { success: false, error: error.toString() };
    }
    });
}

function salvarMovimentacao(dados) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Movimentações');

    if (!sheet) {
      return { success: false, error: 'Aba de movimentações não encontrada' };
    }

    var colunaStatus = garantirEstruturaMovimentacoes(sheet);

    // Buscar número do armário
    var tipoArmarioNormalizado = normalizarTextoBasico(dados.tipoArmario);
    var tipoNormalizado = normalizarTextoBasico(dados.tipo);
    var numeroArmario = normalizarNumeroArmario(dados.numeroArmario);
    if (!numeroArmario) {
      var nomeSheetArmario = tipoNormalizado === 'visitante' ? 'Visitantes' : 'Acompanhantes';
      var armarioSheet = ss.getSheetByName(nomeSheetArmario);
      if (armarioSheet) {
        var estruturaArmario = obterEstruturaPlanilha(armarioSheet);
        var totalLinhasArmario = armarioSheet.getLastRow();
        if (totalLinhasArmario > 1) {
          var dadosArmario = armarioSheet.getRange(2, 1, totalLinhasArmario - 1, estruturaArmario.ultimaColuna || (nomeSheetArmario === 'Visitantes' ? 13 : 12)).getValues();
          for (var i = 0; i < dadosArmario.length; i++) {
            var linha = dadosArmario[i];
            if (String(linha[0]) === String(dados.armarioId)) {
              var numeroLinha = obterValorLinha(linha, estruturaArmario, 'numero', linha[1]);
              numeroArmario = numeroLinha ? numeroLinha.toString().trim() : '';
              break;
            }
          }
        }
      }
    }

    var lastRow = sheet.getLastRow();
    var novoId = lastRow > 1 ? Math.max(...sheet.getRange(2, 1, sheet.getLastRow()-1, 1).getValues().flat()) + 1 : 1;
    
    var registroAtual = obterDataHoraAtualFormatada();
    var dataMovimentacao = formatarDataPlanilha(dados.data);
    var horaMovimentacao = formatarHorarioPlanilha(dados.hora);
    var registroMovimento = registroAtual.dataHoraIso;

    var novaLinha = [
      novoId,
      dados.armarioId,
      numeroArmario,
      dados.tipo,
      dados.descricao,
      dados.responsavel,
      dataMovimentacao,
      horaMovimentacao,
      registroMovimento,
      'ativo'
    ];

    sheet.getRange(lastRow + 1, 1, 1, colunaStatus).setValues([novaLinha]);

    registrarLog('MOVIMENTAÇÃO', `Movimentação registrada para armário ${numeroArmario}`);

    limparCacheMovimentacoes(dados.armarioId, numeroArmario, tipoArmarioNormalizado || tipoNormalizado);

    return { success: true, message: 'Movimentação registrada com sucesso', id: novoId };

  } catch (error) {
    registrarLog('ERRO', `Erro ao salvar movimentação: ${error.toString()}`);
    return { success: false, error: error.toString() };
  }
}

function finalizarMovimentacoesArmario(armarioId, numeroArmario, tipo) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Movimentações');

    if (!sheet || sheet.getLastRow() < 2) {
      return;
    }

    var colunaStatus = garantirEstruturaMovimentacoes(sheet);
    var totalLinhas = sheet.getLastRow() - 1;
    if (totalLinhas <= 0) {
      return;
    }

    var idTexto = armarioId !== undefined && armarioId !== null ? armarioId.toString().trim() : '';
    var numeroNormalizado = normalizarNumeroArmario(numeroArmario);
    var largura = Math.max(colunaStatus, sheet.getLastColumn());
    var dados = sheet.getRange(2, 1, totalLinhas, largura).getValues();
    var statusValores = sheet.getRange(2, colunaStatus, totalLinhas, 1).getValues();
    var houveAlteracao = false;

    for (var i = 0; i < dados.length; i++) {
      var linha = dados[i];
      if (!linha) {
        continue;
      }
      var idLinha = linha[1] !== undefined && linha[1] !== null ? linha[1].toString().trim() : '';
      if (idTexto && idLinha !== idTexto) {
        continue;
      }
      if (numeroNormalizado) {
        var numeroLinha = linha[2] ? linha[2].toString().trim() : '';
        if (normalizarNumeroArmario(numeroLinha) !== numeroNormalizado) {
          continue;
        }
      }
      var statusAtual = normalizarTextoBasico(statusValores[i][0] || linha[colunaStatus - 1]);
      if (statusAtual === 'finalizado') {
        continue;
      }
      statusValores[i][0] = 'finalizado';
      houveAlteracao = true;
    }

    if (houveAlteracao) {
      sheet.getRange(2, colunaStatus, totalLinhas, 1).setValues(statusValores);
      limparCacheMovimentacoes(armarioId, numeroArmario, tipo);
    }

  } catch (error) {
    registrarLog('AVISO_MOVIMENTACAO', 'Falha ao finalizar movimentações: ' + error.toString());
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
    
    var dataLog = obterDataHoraAtualFormatada().dataHoraIso;

    var novaLinha = [
      dataLog,
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
