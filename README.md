# Protocolo-digital-gestão-de-contratos-
Aplicação de gestão de contratos
Repósitorio criado para a automação do processo de gestão de contratos

// ============================================================
// PROTOCOLO DIGITAL – GESTÃO DE CONTRATOS | VERSÃO DEFINITIVA (v4.0)
// Integração: Formatação + IA Claude 3.5 + Tratamento de NF e Contratos
// ============================================================

const CONFIG = {
  PASTA_ID: 'COLE_O_ID_REAL_DA_PASTA_AQUI', // <-- ID da pasta do Drive onde os PDFs vão cair
  MODELO: 'claude-3-5-sonnet-20241022',     // Modelo atualizado para leitura avançada
  ABA: 'Contratos',
  ABA_PARCELAS: 'Parcelas',
  ABA_HIST: 'Historico',
  ABA_PAINEL: 'Painel'
};

const CAB = {
  Contratos: ['Nº Contrato','Cliente','Equipamento','Início','Vencimento','Valor Total',
    'Solicitante', 'Aprovador', 'Status Aprovação', 'Data Aprovação',
    'Status','Dias p/ Vencer','Link PDF','File ID','Lido em',
    'Qtd Contratada','Qtd Recebida','Diferença','Status Recebimento',
    'Nº Parcelas','Parcelas Pagas','Saldo Pendente','Status Financeiro'],
  Parcelas: ['Nº Contrato','Parcela','Vencimento','Valor','Data Pagamento','Valor Pago',
    'Status','Dias p/ Vencer','Saldo Parcela'],
  Historico: ['Data/Hora','Usuário','Nº Contrato','Aba','Evento','Campo','Valor Anterior','Valor Novo']
};

// ==================== 0. MENUS E CONFIGURAÇÕES INICIAIS ====================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚙️ Protocolo Digital')
    .addItem('1. Formatar Planilha (Estrutura)', 'configurarPlanilha')
    .addItem('2. Ativar Piloto Automático (Gatilhos)', 'criarGatilhos')
    .addToUi();
}

// Rode UMA vez e apague a chave do código depois por segurança
function salvarChaveAPI() {
  PropertiesService.getScriptProperties().setProperty('ANTHROPIC_KEY', 'COLE_SUA_CHAVE_AQUI');
  SpreadsheetApp.getUi().alert("Chave salva com segurança!");
}

function configurarPlanilha() {
  const ss = SpreadsheetApp.getActive();
  const obter = nome => ss.getSheetByName(nome) || ss.insertSheet(nome);

  // ABA CONTRATOS
  const c = obter(CONFIG.ABA);
  c.getRange(1, 1, 1, CAB.Contratos.length).setValues([CAB.Contratos])
    .setFontWeight('bold').setBackground('#263238').setFontColor('#fff');
  c.setFrozenRows(1);
  c.getRange('D:E').setNumberFormat('dd/MM/yyyy');
  c.getRange('J:J').setNumberFormat('dd/MM/yyyy');
  c.getRange('F:F').setNumberFormat('R$ #,##0.00');
  c.getRange('V:V').setNumberFormat('R$ #,##0.00');
  c.getRange('G:J').setBackground('#e3f2fd'); // Azul claro para Aprovações
  c.getRange('Q:Q').setBackground('#fff8e1'); // Amarelo para Qtd Recebida
  
  // ABA PARCELAS
  const p = obter(CONFIG.ABA_PARCELAS);
  p.getRange(1, 1, 1, CAB.Parcelas.length).setValues([CAB.Parcelas])
    .setFontWeight('bold').setBackground('#263238').setFontColor('#fff');
  p.setFrozenRows(1);
  p.getRange('C:E').setNumberFormat('dd/MM/yyyy');
  p.getRange('D:D').setNumberFormat('R$ #,##0.00');
  p.getRange('F:F').setNumberFormat('R$ #,##0.00');
  p.getRange('I:I').setNumberFormat('R$ #,##0.00');
  p.getRange('E:F').setBackground('#fff8e1'); 

  // ABA HISTORICO
  const h = obter(CONFIG.ABA_HIST);
  h.getRange(1, 1, 1, CAB.Historico.length).setValues([CAB.Historico])
    .setFontWeight('bold').setBackground('#263238').setFontColor('#fff');
  h.setFrozenRows(1);
  h.getRange('A:A').setNumberFormat('dd/MM/yyyy HH:mm:ss');
  
  // Aplica as fórmulas nas primeiras 100 linhas para garantir
  aplicarFormulasContrato(c, 2, 100);
  aplicarFormulasParcela(p, 2, 100);
  
  montarPainel(obter(CONFIG.ABA_PAINEL));
  SpreadsheetApp.getUi().alert("✅ Planilha formatada com sucesso!");
}

function aplicarFormulasContrato(aba, linhaIni, n) {
  const f = (col, r1c1) => aba.getRange(linhaIni, col, n, 1).setFormulaR1C1(r1c1);
  f(11, '=IF(RC5="","Sem data",IF(RC5<TODAY(),"Vencido","No Prazo"))');
  f(12, '=IF(RC5="","",RC5-TODAY())');
  f(18, '=IF(RC16="","",RC16-N(RC17))');
  f(19, '=IF(RC16="","",IF(N(RC17)=0,"Pendente",IF(N(RC17)<RC16,"Parcial",IF(N(RC17)=RC16,"Completo","Excedente"))))');
  f(20, '=COUNTIF(Parcelas!C1,RC1)');
  f(21, '=COUNTIFS(Parcelas!C1,RC1,Parcelas!C7,"Paga")');
  f(22, '=IF(RC20=0,N(RC6),SUMIFS(Parcelas!C9,Parcelas!C1,RC1))');
  f(23, '=IF(RC20=0,"Sem parcelas",IF(RC22<=0,"Liquidado",IF(COUNTIFS(Parcelas!C1,RC1,Parcelas!C7,"Atrasada")>0,"Em atraso","Em dia")))');
}

function aplicarFormulasParcela(aba, linhaIni, n) {
  const f = (col, r1c1) => aba.getRange(linhaIni, col, n, 1).setFormulaR1C1(r1c1);
  f(7, '=IF(AND(RC4<>"",N(RC6)>=RC4),"Paga",IF(RC3<TODAY(),"Atrasada",IF(N(RC6)>0,"Parcial","A vencer")))');
  f(8, '=IF(RC7="Paga","",RC3-TODAY())');
  f(9, '=MAX(0,N(RC4)-N(RC6))');
}

// ==================== 1. MOTOR DE INTELIGÊNCIA ARTIFICIAL (CLAUDE) ====================

function processarNovosPDFs() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return;

  try {
    const ss = SpreadsheetApp.getActive();
    const aba = ss.getSheetByName(CONFIG.ABA);
    const abaP = ss.getSheetByName(CONFIG.ABA_PARCELAS);
    if (!aba || !abaP) return;

    const pastaOrigem = DriveApp.getFolderById(CONFIG.PASTA_ID);
    
    // Cria pasta "Processados" para não ler o mesmo contrato duas vezes
    const pastasProcessados = pastaOrigem.getFoldersByName("Processados");
    const pastaDestino = pastasProcessados.hasNext() ? pastasProcessados.next() : pastaOrigem.createFolder("Processados");

    const arquivos = pastaOrigem.getFilesByType(MimeType.PDF);
    
    while (arquivos.hasNext()) {
      const arq = arquivos.next();

      try {
        const d = extrairDados(arq);
        const id = d.numero_contrato || ('SN-' + arq.getId().slice(0, 8));
        const linha = aba.getLastRow() + 1;
        const valor = Number(d.valor) || 0;

        // LÓGICA DE ALÇADAS (Substitua os valores conforme política da sua empresa)
        let aprovador = "Gerente de Área";
        if (valor > 150000) aprovador = "Conselho/Presidência";
        else if (valor > 50000) aprovador = "Diretoria";

        aba.getRange(linha, 1, 1, 17).setValues([[
          id, d.cliente, d.equipamento,
          paraData(d.data_inicio), paraData(d.data_vencimento), d.valor,
          d.solicitante || 'A Definir', aprovador, 'Pendente', '', 
          '', '', arq.getUrl(), arq.getId(), new Date(),
          d.quantidade_total ?? '', 0 
        ]]);
        
        aplicarFormulasContrato(aba, linha, 1);
        gerarParcelas(abaP, id, d);
        
        // Move o arquivo para limpar a caixa de entrada
        arq.moveTo(pastaDestino);
        registrarHistorico(id, CONFIG.ABA, 'Importado por IA', '', '', `Valor: R$ ${valor} | Aprovador: ${aprovador}`);
      
      } catch (e) {
        registrarHistorico(arq.getName(), CONFIG.ABA, 'Erro IA', '', '', e.message.slice(0, 200));
      }
    }
  } finally {
    lock.releaseLock();
  }
}

function extrairDados(arquivo) {
  const chave = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_KEY');
  if (!chave) throw new Error('Chave da API não configurada.');
  
  const pdf = Utilities.base64Encode(arquivo.getBlob().getBytes());
  
  // PROMPT BLINDADO: Resolve as falhas matemáticas e de distinção entre NF e Contrato
  const prompt =
    'Você é um analista financeiro rigoroso extraindo dados de Contratos de Locação ou Notas Fiscais. ' +
    'Responda EXCLUSIVAMENTE com um JSON válido, sem nenhum texto adicional. Estrutura obrigatória: ' +
    '{"numero_contrato":"","cliente":"","equipamento":"","quantidade_total":0,' +
    '"data_inicio":"AAAA-MM-DD","data_vencimento":"AAAA-MM-DD","valor":0,"solicitante":"",' +
    '"numero_parcelas":0,"valor_parcela":0,"data_primeira_parcela":"AAAA-MM-DD",' +
    '"parcelas":[{"numero":1,"vencimento":"AAAA-MM-DD","valor":0}]}. ' +
    'REGRAS DE EXTRAÇÃO: ' +
    '1. numero_contrato: Se for Nota Fiscal, use o número da NF. Se for Aditivo, use o número do Aditivo. ' +
    '2. equipamento/quantidade: Agrupe os itens (Ex: "2x iPhone 17e"). quantidade_total é a soma de unidades. ' +
    '3. valor: VALOR TOTAL. Se for contrato de locação com prazo (ex: 48 meses) e valor mensal, MULTIPLIQUE o valor mensal pelo prazo para achar o valor total. Se for NF, use o total da nota. ' +
    '4. datas: data_inicio é a data de assinatura ou emissão. Se houver prazo (ex: 48 meses) sem data de vencimento expressa, CALCULE a data_vencimento somando os meses à data_inicio. Na NF de compra definitiva, repita a data de emissão na data_vencimento. ' +
    '5. solicitante: Procure o Destinatário. Se vazio, use o nome do signatário da empresa (ex: quem assinou eletronicamente). Se não achar, "Não identificado". ' +
    '6. parcelas: Na NF à vista ou sem detalhamento, envie [] (lista vazia) e 0 para numero_parcelas. ' +
    'Se não encontrar uma informação, use null. Não invente dados.';

  const resp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    headers: { 
      'x-api-key': chave, 
      'anthropic-version': '2023-06-01',
      'anthropic-beta': 'pdfs-2024-09-25' // Header obrigatório para ler PDFs na Anthropic
    },
    payload: JSON.stringify({
      model: CONFIG.MODELO,
      max_tokens: 1500,
      temperature: 0,
      messages: [{ role: 'user', content: [
        { type: 'document', source: { type: 'base64', media_type: 'application/pdf', data: pdf } },
        { type: 'text', text: prompt }
      ]}]
    }),
    muteHttpExceptions: true
  });

  if (resp.getResponseCode() !== 200) throw new Error(`Erro API: ${resp.getContentText()}`);
  const texto = JSON.parse(resp.getContentText()).content[0].text;
  return JSON.parse(texto.replace(/```json|```/g, '').trim());
}

// ==================== 2. FUNÇÕES AUXILIARES E PAINEL ====================

function gerarParcelas(abaP, id, d) {
  let linhas = [];
  if (Array.isArray(d.parcelas) && d.parcelas.length) {
    linhas = d.parcelas.map((p, i) => [id, p.numero || i + 1, paraData(p.vencimento), Number(p.valor) || '', '', '']);
  } else if (Number(d.numero_parcelas) > 0) {
    const n = Number(d.numero_parcelas);
    const total = Number(d.valor) || 0;
    const vp = Number(d.valor_parcela) || (total ? total / n : 0);
    const base = paraData(d.data_primeira_parcela) || paraData(d.data_inicio);
    if (!base) return 0;

    for (let i = 0; i < n; i++) {
      let v = (i === n - 1 && total) ? (total - vp * (n - 1)) : vp;
      linhas.push([id, i + 1, addMeses(base, i), v || '', '', '']);
    }
  }
  if (!linhas.length) return 0;
  
  const ini = abaP.getLastRow() + 1;
  abaP.getRange(ini, 1, linhas.length, 6).setValues(linhas);
  aplicarFormulasParcela(abaP, ini, linhas.length);
  return linhas.length;
}

function paraData(iso) {
  if (!iso || typeof iso !== 'string') return '';
  const [a, m, d] = iso.split('-').map(Number);
  return a && m && d ? new Date(a, m - 1, d) : '';
}

function addMeses(dt, n) {
  return new Date(dt.getFullYear(), dt.getMonth() + n, dt.getDate());
}

function registrarHistorico(contrato, aba, evento, campo, antes, depois, usuario) {
  const h = SpreadsheetApp.getActive().getSheetByName(CONFIG.ABA_HIST);
  if (!h) return;
  const user = usuario || Session.getEffectiveUser().getEmail() || 'sistema';
  h.appendRow([new Date(), user, contrato, aba, evento, campo, String(antes ?? ''), String(depois ?? '')]);
}

function aoEditar(e) {
  if (!e || !e.range) return;
  const sh = e.range.getSheet();
  if (e.range.getRow() < 2 || !['Contratos', 'Parcelas'].includes(sh.getName())) return;
  const contrato = sh.getRange(e.range.getRow(), 1).getValue();
  const c = e.range.getColumn();
  registrarHistorico(contrato, sh.getName(), 'Alteração manual', `Coluna ${c}`, e.oldValue ?? '', e.range.getDisplayValue());
}

function montarPainel(p) {
  p.clear();
  p.getCharts().forEach(ch => p.removeChart(ch));
  const C = `'${CONFIG.ABA}'!`, P = `'${CONFIG.ABA_PARCELAS}'!`;

  const blocos = [
    ['A1', [['PAINEL GERENCIAL – PROTOCOLO DIGITAL', '']]],
    ['A2', [['Atualizado em', '=NOW()']]],
    ['A4', [
      ['CONTRATOS E APROVAÇÕES', ''],
      ['Total de Contratos', `=COUNTA(${C}A2:A)`],
      ['Pendentes de Aprovação', `=COUNTIF(${C}I2:I,"Pendente")`],
      ['Valor total contratado', `=SUM(${C}F2:F)`]
    ]],
    ['A10', [
      ['EQUIPAMENTOS', ''],
      ['Qtd contratada', `=SUM(${C}P2:P)`],
      ['Qtd recebida', `=SUM(${C}Q2:Q)`],
      ['Pendente de recebimento', `=SUMIF(${C}R2:R,">0")`]
    ]],
    ['D10', [
      ['Status recebimento', 'Contratos'],
      ['Completo', `=COUNTIF(${C}S2:S,D11)`],
      ['Parcial', `=COUNTIF(${C}S2:S,D12)`],
      ['Pendente', `=COUNTIF(${C}S2:S,D13)`]
    ]],
    ['A17', [
      ['PARCELAS', ''],
      ['Total de parcelas', `=COUNTA(${P}A2:A)`],
      ['Valor pago', `=SUM(${P}F2:F)`],
      ['Saldo em aberto (Geral)', `=SUM(${C}V2:V)`],
      ['Valor em atraso', `=SUMIF(${P}G2:G,"Atrasada",${P}I2:I)`]
    ]],
    ['A25', [['TOP 5 SALDOS EM ABERTO', ''], ['Nº Contrato', 'Cliente']]]
  ];

  blocos.forEach(([cel, vals]) => {
    const r = p.getRange(cel);
    p.getRange(r.getRow(), r.getColumn(), vals.length, 2).setValues(vals);
  });
  
  p.getRange('C26:D26').setValues([['Saldo', 'Status']]);
  p.getRange('A27').setFormula(`=IFERROR(QUERY(${C}A2:W,"select A, B, V, W where V > 0 order by V desc limit 5",0),"Nenhum")`);

  p.getRange('A1').setFontSize(16).setFontWeight('bold');
  ['A4:B4', 'A10:B10', 'D10:E10', 'A17:B17', 'A25:D26'].forEach(a => p.getRange(a).setFontWeight('bold').setBackground('#263238').setFontColor('#fff'));
  ['B7', 'B19:B20', 'C27:C31'].forEach(a => p.getRange(a).setNumberFormat('R$ #,##0.00'));
  p.setColumnWidth(1, 220); p.setColumnWidth(4, 160);
}

// ==================== 3. ATIVAÇÃO DO PILOTO AUTOMÁTICO ====================

function criarGatilhos() {
  // Limpa gatilhos antigos para não duplicar
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  
  // Gatilho 1: Acorda o robô a cada 15 minutos para procurar PDFs novos no Drive
  ScriptApp.newTrigger('processarNovosPDFs').timeBased().everyMinutes(15).create();
  
  // Gatilho 2: Registra no histórico qualquer célula que alguém editar manualmente
  ScriptApp.newTrigger('aoEditar').forSpreadsheet(SpreadsheetApp.getActive()).onEdit().create();
  
  SpreadsheetApp.getUi().alert("🚀 Piloto Automático Ativado!\nA planilha agora buscará novos contratos no Drive a cada 15 minutos.");
}
