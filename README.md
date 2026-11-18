/**
 * SISTEMA AUTOMÁTICO DE GESTÃO DE CONTRATOS (V.9)
 * Sincroniza novos registros de 'Arquivos_Drive' para 'Contratos' e 'Parcelas'.
 */

// Cria um menu personalizado na barra superior da planilha
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📑 Gestão de Contratos')
    .addItem('Sincronizar Novos Arquivos Agora', 'sincronizarNovosContratos')
    .addToUi();
}

/**
 * Função principal de sincronização
 */
function sincronizarNovosContratos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaArquivos = ss.getSheetByName('Arquivos_Drive');
  const abaContratos = ss.getSheetByName('Contratos');
  const abaParcelas = ss.getSheetByName('Parcelas');

  if (!abaArquivos || !abaContratos || !abaParcelas) {
    Logger.log('Erro: Uma ou mais abas necessárias não foram encontradas.');
    return;
  }

  // 1. Obter dados de Arquivos_Drive (ignora o cabeçalho na linha 1)
  const dadosArquivos = abaArquivos.getDataRange().getValues();
  if (dadosArquivos.length <= 1) return; // Sem dados

  // 2. Mapear File IDs e Nºs de Contratos já existentes para evitar duplicidade
  const dadosContratos = abaContratos.getDataRange().getValues();
  const fileIdsExistentes = new Set();
  const contratosExistentes = new Set();

  for (let i = 1; i < dadosContratos.length; i++) {
    const numContrato = String(dadosContratos[i][0] || '').trim();
    const fileId = String(dadosContratos[i][13] || '').trim(); // Coluna N (índice 13)
    if (fileId) fileIdsExistentes.add(fileId);
    if (numContrato) contratosExistentes.add(numContrato);
  }

  const novosContratos = [];
  const novasParcelas = [];
  const dataHoje = new Date();

  // 3. Varrer Arquivos_Drive em busca de novos itens
  for (let i = 1; i < dadosArquivos.length; i++) {
    const nomeArquivo = String(dadosArquivos[i][0] || '').trim();
    const linkArquivo = String(dadosArquivos[i][1] || '').trim();
    const dataCriacao = dadosArquivos[i][2] || dataHoje;
    const fileId = String(dadosArquivos[i][3] || '').trim();

    if (!fileId && !nomeArquivo) continue;

    // Se já foi cadastrado, pula
    if (fileId && fileIdsExistentes.has(fileId)) continue;

    // Extrair dados do nome do arquivo (Exemplo: "CT-2024-001 - Fornecedor - Equipamento")
    const metadados = extrairMetadadosDoNome(nomeArquivo);

    // Garante que o número de contrato seja único
    let numContrato = metadados.numContrato;
    if (!numContrato || contratosExistentes.has(numContrato)) {
      numContrato = `CT-${Utilities.formatDate(new Date(), 'GMT-3', 'yyyy')}-${String(contratosExistentes.size + 1).padStart(3, '0')}`;
    }

    contratosExistentes.add(numContrato);
    if (fileId) fileIdsExistentes.add(fileId);

    // Monta a linha para a aba 'Contratos' (Colunas A até Q)
    // As colunas K, L e R até Y utilizam as fórmulas matriciais automáticas já configuradas
    const linhaContrato = [
      numContrato,                                     // A: Nº Contrato
      metadados.fornecedor || 'Pendente Cadastro',     // B: Fornecedor
      metadados.equipamento || 'Geral',               // C: Equipamento
      dataCriacao,                                     // D: Início
      adicionarMeses(dataCriacao, 12),                 // E: Vencimento (padrão: 12 meses)
      metadados.valorTotal || 0,                       // F: Valor Total
      Session.getActiveUser().getEmail() || '',        // G: Solicitante (e-mail)
      '',                                              // H: Aprovador
      'Pendente',                                      // I: Status Aprovação
      '',                                              // J: Data Aprovação
      '',                                              // K: Status (calculado via matriz)
      '',                                              // L: Dias p/ Vencer (calculado via matriz)
      linkArquivo,                                     // M: Link PDF
      fileId,                                          // N: File ID
      Utilities.formatDate(dataHoje, 'GMT-3', 'dd/MM/yyyy HH:mm'), // O: Lido em
      1,                                               // P: Qtd Contratada (padrão)
      0                                                // Q: Qtd Recebida (padrão)
    ];

    novosContratos.push(linhaContrato);

    // Se houver valor cadastrado, gera a parcela inicial em 'Parcelas'
    if (metadados.valorTotal > 0) {
      novasParcelas.push([
        numContrato,                        // A: Nº Contrato
        1,                                  // B: Parcela
        adicionarMeses(dataCriacao, 1),     // C: Vencimento da 1ª parcela
        metadados.valorTotal,               // D: Valor
        '',                                 // E: Data Pagamento
        0                                   // F: Valor Pago
      ]);
    }
  }

  // 4. Inserir os novos contratos
  if (novosContratos.length > 0) {
    const ultimaLinhaContratos = abaContratos.getLastRow();
    abaContratos.getRange(ultimaLinhaContratos + 1, 1, novosContratos.length, novosContratos[0].length)
      .setValues(novosContratos);

    // 5. Inserir as novas parcelas
    if (novasParcelas.length > 0) {
      const ultimaLinhaParcelas = abaParcelas.getLastRow();
      abaParcelas.getRange(ultimaLinhaParcelas + 1, 1, novasParcelas.length, novasParcelas[0].length)
        .setValues(novasParcelas);
    }

    Logger.log(`${novosContratos.length} novo(s) contrato(s) sincronizado(s) com sucesso.`);
  } else {
    Logger.log('Nenhum contrato novo para sincronizar.');
  }
}

/**
 * Utilitário: Tenta decompor o nome do arquivo em metadados
 */
function extrairMetadadosDoNome(nome) {
  let resultado = {
    numContrato: '',
    fornecedor: '',
    equipamento: '',
    valorTotal: 0
  };

  if (!nome) return resultado;

  // Regex para identificar padrões como CT-001, CONTRATO-123, etc.
  const regexCT = /(CT[-_\s]?\d+|CONTRATO[-_\s]?\d+)/i;
  const matchCT = nome.match(regexCT);
  if (matchCT) {
    resultado.numContrato = matchCT[1].toUpperCase().replace(/\s+/g, '-');
  }

  // Se o arquivo tiver partes separadas por traço (" - ")
  const partes = nome.replace(/\.[^/.]+$/, '').split(/\s*-\s*/);
  if (partes.length >= 2) {
    if (!resultado.numContrato) resultado.numContrato = partes[0].trim();
    resultado.fornecedor = partes[1].trim();
    if (partes.length >= 3) {
      resultado.equipamento = partes[2].trim();
    }
  } else if (!resultado.fornecedor) {
    resultado.fornecedor = nome.replace(/\.[^/.]+$/, '');
  }

  return resultado;
}

/**
 * Utilitário: Adiciona meses a uma data
 */
function adicionarMeses(dataBase, meses) {
  const d = new Date(dataBase);
  d.setMonth(d.getMonth() + meses);
  return d;
}
