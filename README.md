# Protocolo-digital-gestão-de-contratos-
Aplicação de gestão de contratos
Repósitorio criado para a automação do processo de gestão de contratos


// ============================================================
// PROTOCOLO DIGITAL – GESTÃO DE CONTRATOS  |  v3
// ============================================================
const CONFIG = {
  PASTA_ID: 'COLE_O_ID_REAL',              // ID da URL da pasta no Drive
  EMAIL: 'administrativo@maino.com.br',
  DIAS_ALERTA: 14,                          // janela para o envio semanal
  MODELO: 'claude-haiku-4-5',               // confira em docs.anthropic.com
  ABA: 'Contratos',
  ABA_PARCELAS: 'Parcelas',
  ABA_HIST: 'Historico',
  ABA_PAINEL: 'Painel',
  TZ: 'America/Sao_Paulo'
};

const CAB = {
  Contratos: ['Nº Contrato','Cliente','Equipamento','Início','Vencimento','Valor Total',
    'Status','Dias p/ Vencer','Link PDF','File ID','Lido em',
    'Qtd Contratada','Qtd Recebida','Diferença','Status Recebimento',
    'Nº Parcelas','Parcelas Pagas','Saldo Pendente','Status Financeiro'],
  Parcelas: ['Nº Contrato','Parcela','Vencimento','Valor','Data Pagamento','Valor Pago',
    'Status','Dias p/ Vencer','Saldo Parcela'],
  Historico: ['Data/Hora','Usuário','Nº Contrato','Aba','Evento','Campo','Valor Anterior','Valor Novo']
};

// Rode UMA vez e depois apague a chave do código
function salvarChave() {
  PropertiesService.getScriptProperties().setProperty('ANTHROPIC_KEY', 'COLE_SUA_CHAVE_AQUI');
}

// ==================== 0. CONFIGURAÇÃO DA PLANILHA (rodar 1 vez) ====================
function configurarPlanilha() {
  const ss = SpreadsheetApp.getActive();
  const obter = nome => ss.getSheetByName(nome) || ss.insertSheet(nome);

  // Contratos: preserva A:K e acrescenta L:S
  const c = obter(CONFIG.ABA);
  c.getRange(1, 1, 1, CAB.Contratos.length).setValues([CAB.Contratos])
    .setFontWeight('bold').setBackground('#263238').setFontColor('#fff');
  c.setFrozenRows(1);
  c.getRange('D:E').setNumberFormat('dd/MM/yyyy');
  c.getRange('F:F').setNumberFormat('R$ #,##0.00');
  c.getRange('R:R').setNumberFormat('R$ #,##0.00');
  c.getRange('M2:M').setBackground('#fff8e1'); // coluna de preenchimento manual
  if (c.getLastRow() > 1) aplicarFormulasContrato(c, 2, c.getLastRow() - 1); // migra as linhas antigas

  const p = obter(CONFIG.ABA_PARCELAS);
  p.getRange(1, 1, 1, CAB.Parcelas.length).setValues([CAB.Parcelas])
    .setFontWeight('bold').setBackground('#263238').setFontColor('#fff');
  p.setFrozenRows(1);
  p.getRange('C:C').setNumberFormat('dd/MM/yyyy');
  p.getRange('E:E').setNumberFormat('dd/MM/yyyy');
  p.getRange('D:D').setNumberFormat('R$ #,##0.00');
  p.getRange('F:F').setNumberFormat('R$ #,##0.00');
  p.getRange('I:I').setNumberFormat('R$ #,##0.00');
  p.getRange('E2:F').setBackground('#fff8e1');

  const h = obter(CONFIG.ABA_HIST);
  h.getRange(1, 1, 1, CAB.Historico.length).setValues([CAB.Historico])
    .setFontWeight('bold').setBackground('#263238').setFontColor('#fff');
  h.setFrozenRows(1);
  h.getRange('A:A').setNumberFormat('dd/MM/yyyy HH:mm:ss');
  // Proteção: só o dono do script edita o histórico
  const prot = h.protect().setDescription('Histórico – somente sistema');
  prot.removeEditors(prot.getEditors());
  prot.addEditor(Session.getEffectiveUser());

  montarPainel(obter(CONFIG.ABA_PAINEL));
  registrarHistorico('-', 'Sistema', 'Planilha configurada (v3)', '', '', '');
}

function aplicarFormulasContrato(aba, linhaIni, n) {
  const f = (col, r1c1) => aba.getRange(linhaIni, col, n, 1).setFormulaR1C1(r1c1);
  f(7,  '=IF(RC5="","Sem data",IF(RC5<TODAY(),"Vencido","No Prazo"))');
  f(8,  '=IF(RC5="","",RC5-TODAY())');
  f(14, '=IF(RC12="","",RC12-N(RC13))');
  f(15, '=IF(RC12="","Sem qtd",IF(N(RC13)=0,"Pendente",IF(N(RC13)<RC12,"Parcial",IF(N(RC13)=RC12,"Completo","Excedente"))))');
  f(16, '=COUNTIF(Parcelas!C1,RC1)');
  f(17, '=COUNTIFS(Parcelas!C1,RC1,Parcelas!C7,"Paga")');
  f(18, '=IF(RC16=0,N(RC6),SUMIFS(Parcelas!C9,Parcelas!C1,RC1))');
  f(19, '=IF(RC16=0,"Sem parcelas",IF(RC18<=0,"Liquidado",IF(COUNTIFS(Parcelas!C1,RC1,Parcelas!C7,"Atrasada")>0,"Em atraso","Em dia")))');
}

function aplicarFormulasParcela(aba, linhaIni, n) {
  const f = (col, r1c1) => aba.getRange(linhaIni, col, n, 1).setFormulaR1C1(r1c1);
  f(7, '=IF(AND(RC4<>"",N(RC6)>=RC4),"Paga",IF(RC3<TODAY(),"Atrasada",IF(N(RC6)>0,"Parcial","A vencer")))');
  f(8, '=IF(RC7="Paga","",RC3-TODAY())');
  f(9, '=MAX(0,N(RC4)-N(RC6))');
}

// ==================== 1. LEITURA DOS PDFs ====================
function processarNovosPDFs() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return;

  try {
    const ss = SpreadsheetApp.getActive();
    const aba = ss.getSheetByName(CONFIG.ABA);
    const abaP = ss.getSheetByName(CONFIG.ABA_PARCELAS);
    if (!aba || !abaP) throw new Error('Rode configurarPlanilha() antes.');

    const ultima = aba.getLastRow();
    const jaLidos = ultima > 1 ? aba.getRange(2, 10, ultima - 1, 1).getValues().flat() : [];
    const arquivos = DriveApp.getFolderById(CONFIG.PASTA_ID).getFilesByType(MimeType.PDF);
    const inicio = Date.now();

    while (arquivos.hasNext()) {
      if (Date.now() - inicio > 4.5 * 60 * 1000) break;
      const arq = arquivos.next();
      if (jaLidos.includes(arq.getId())) continue;

      try {
        const d = extrairDados(arq);
        const id = d.numero_contrato || ('SN-' + arq.getId().slice(0, 8));
        const linha = aba.getLastRow() + 1;

        aba.getRange(linha, 1, 1, 13).setValues([[
          id, d.cliente, d.equipamento,
          paraData(d.data_inicio), paraData(d.data_vencimento), d.valor,
          '', '', arq.getUrl(), arq.getId(), new Date(),
          d.quantidade_total ?? '', 0
        ]]);
        aplicarFormulasContrato(aba, linha, 1);

        const nParc = gerarParcelas(abaP, id, d);
        registrarHistorico(id, CONFIG.ABA, 'Contrato importado por IA', '', '',
          `Valor: ${d.valor ?? 'n/d'} | Qtd: ${d.quantidade_total ?? 'n/d'} | Parcelas: ${nParc}`);
        if (!nParc) registrarHistorico(id, CONFIG.ABA_PARCELAS, 'Parcelas não identificadas',
          '', '', 'Lançar manualmente');
      } catch (e) {
        aba.appendRow(['ERRO', arq.getName(), e.message.slice(0, 200),
          '', '', '', 'Revisar', '', arq.getUrl(), arq.getId(), new Date()]);
        registrarHistorico(arq.getName(), CONFIG.ABA, 'Falha na leitura', '', '', e.message.slice(0, 200));
        MailApp.sendEmail(CONFIG.EMAIL, '⚠️ Falha ao ler contrato',
          `Arquivo: ${arq.getName()}\nErro: ${e.message}\nLink: ${arq.getUrl()}`);
      }
    }
  } finally {
    lock.releaseLock();
  }
}

function extrairDados(arquivo) {
  const chave = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_KEY');
  if (!chave) throw new Error('Chave da API não configurada (rode salvarChave).');
  const pdf = Utilities.base64Encode(arquivo.getBlob().getBytes());

  const prompt =
    'Você extrai dados de contratos de locação de equipamentos. ' +
    'Responda SOMENTE com JSON válido, sem texto extra, neste formato: ' +
    '{"numero_contrato":"","cliente":"","equipamento":"","quantidade_total":0,' +
    '"data_inicio":"AAAA-MM-DD","data_vencimento":"AAAA-MM-DD","valor":0,' +
    '"numero_parcelas":0,"valor_parcela":0,"data_primeira_parcela":"AAAA-MM-DD",' +
    '"parcelas":[{"numero":1,"vencimento":"AAAA-MM-DD","valor":0}]}. ' +
    'valor = valor TOTAL do contrato (não a parcela), número sem R$ e com ponto decimal. ' +
    'quantidade_total = soma das unidades de todos os equipamentos. ' +
    'Se houver vários equipamentos, liste-os em "equipamento" separados por "; " com a quantidade (ex: "2x Notebook; 1x Impressora"). ' +
    'Preencha "parcelas" somente se o contrato listar explicitamente cada parcela; senão use []. ' +
    'Se um dado não existir, use null. Não invente informações.';

  const resp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-api-key': chave, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify({
      model: CONFIG.MODELO,
      max_tokens: 2048,
      messages: [{ role: 'user', content: [
        { type: 'document', source: { type: 'base64', media_type: 'application/pdf', data: pdf } },
        { type: 'text', text: prompt }
      ]}]
    }),
    muteHttpExceptions: true
  });

  if (resp.getResponseCode() !== 200) {
    throw new Error(`API ${resp.getResponseCode()}: ${resp.getContentText()}`);
  }
  const texto = JSON.parse(resp.getContentText()).content[0].text;
  return JSON.parse(texto.replace(/```json|```/g, '').trim());
}

// Prioridade: lista explícita > (nº parcelas + valor/1º venc.) > nenhuma
function gerarParcelas(abaP, id, d) {
  let linhas = [];

  if (Array.isArray(d.parcelas) && d.parcelas.length) {
    linhas = d.parcelas.map((p, i) => [id, p.numero || i + 1, paraData(p.vencimento), Number(p.valor) || '', '', '']);
  } else if (Number(d.numero_parcelas) > 0) {
    const n = Number(d.numero_parcelas);
    const total = Number(d.valor) || 0;
    const vp = Number(d.valor_parcela) || (total ? Math.round(total / n * 100) / 100 : 0);
    const base = paraData(d.data_primeira_parcela) || (paraData(d.data_inicio) ? addMeses(paraData(d.data_inicio), 1) : null);
    if (!base) return 0;

    for (let i = 0; i < n; i++) {
      let v = vp;
      if (!d.valor_parcela && total && i === n - 1) v = Math.round((total - vp * (n - 1)) * 100) / 100; // ajuste de arredondamento
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
  const r = new Date(dt.getFullYear(), dt.getMonth() + n, 1);
  const ultimoDia = new Date(r.getFullYear(), r.getMonth() + 1, 0).getDate();
  r.setDate(Math.min(dt.getDate(), ultimoDia)); // evita 31/01 → 03/03
  return r;
}

// ==================== 2. HISTÓRICO ====================
function registrarHistorico(contrato, aba, evento, campo, antes, depois, usuario) {
  const h = SpreadsheetApp.getActive().getSheetByName(CONFIG.ABA_HIST);
  if (!h) return;
  const user = usuario || Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || 'sistema';
  h.appendRow([new Date(), user, contrato, aba, evento, campo, String(antes ?? ''), String(depois ?? '')]);
}

// Gatilho instalável de edição (criado em criarGatilhos)
function aoEditar(e) {
  if (!e || !e.range) return;
  const sh = e.range.getSheet();
  const nome = sh.getName();
  const monitoradas = {
    [CONFIG.ABA]: [1, 2, 3, 4, 5, 6, 12, 13],
    [CONFIG.ABA_PARCELAS]: [3, 4, 5, 6]
  };
  if (!monitoradas[nome] || e.range.getRow() < 2) return;

  const r = e.range.getRow(), c = e.range.getColumn();
  const contrato = sh.getRange(r, 1).getValue();
  const user = e.user && e.user.getEmail();

  if (e.range.getNumRows() > 1 || e.range.getNumColumns() > 1) {
    registrarHistorico(contrato, nome, 'Edição em lote', e.range.getA1Notation(), '', '(múltiplas células)', user);
    return;
  }
  if (!monitoradas[nome].includes(c)) return;

  const campo = sh.getRange(1, c).getValue();
  const parcela = nome === CONFIG.ABA_PARCELAS ? ` (parcela ${sh.getRange(r, 2).getValue()})` : '';
  registrarHistorico(contrato, nome, 'Alteração manual' + parcela, campo,
    e.oldValue ?? '', e.range.getDisplayValue() || '(apagado)', user);
}

// ==================== 3. PAINEL ====================
function montarPainel(p) {
  p.clear();
  p.getCharts().forEach(ch => p.removeChart(ch));
  const C = `'${CONFIG.ABA}'!`, P = `'${CONFIG.ABA_PARCELAS}'!`;

  const blocos = [
    ['A1', [['PAINEL – GESTÃO DE CONTRATOS', '']]],
    ['A2', [['Atualizado em', '=NOW()']]],
    ['A4', [
      ['CONTRATOS', ''],
      ['Contratos lidos', `=COUNTIFS(${C}J2:J,"<>",${C}A2:A,"<>ERRO")`],
      ['Valor total contratado', `=SUMIFS(${C}F2:F,${C}A2:A,"<>ERRO")`],
      ['Contratos vencidos', `=COUNTIF(${C}G2:G,"Vencido")`],
      ['Leituras com erro', `=COUNTIF(${C}A2:A,"ERRO")`]
    ]],
    ['A10', [
      ['EQUIPAMENTOS', ''],
      ['Qtd contratada', `=SUM(${C}L2:L)`],
      ['Qtd recebida', `=SUM(${C}M2:M)`],
      ['Pendente de recebimento', `=SUMIF(${C}N2:N,">0")`],
      ['% recebido', '=IFERROR(B12/B11,0)']
    ]],
    ['D10', [
      ['Status recebimento', 'Contratos'],
      ['Completo', `=COUNTIF(${C}O2:O,D11)`],
      ['Parcial', `=COUNTIF(${C}O2:O,D12)`],
      ['Pendente', `=COUNTIF(${C}O2:O,D13)`],
      ['Excedente', `=COUNTIF(${C}O2:O,D14)`]
    ]],
    ['A17', [
      ['PARCELAS', ''],
      ['Total de parcelas', `=COUNTA(${P}A2:A)`],
      ['Pagas', `=COUNTIF(${P}G2:G,"Paga")`],
      ['Pendentes', '=B18-B19'],
      ['Valor pago', `=SUM(${P}F2:F)`],
      ['Saldo em aberto (contratos)', `=SUMIFS(${C}R2:R,${C}A2:A,"<>ERRO")`],
      ['Valor em atraso', `=SUMIF(${P}G2:G,"Atrasada",${P}I2:I)`],
      ['Contratos liquidados', `=COUNTIF(${C}S2:S,"Liquidado")`]
    ]],
    ['D17', [
      ['Status parcelas', 'Qtd'],
      ['Paga', `=COUNTIF(${P}G2:G,D18)`],
      ['A vencer', `=COUNTIF(${P}G2:G,D19)`],
      ['Atrasada', `=COUNTIF(${P}G2:G,D20)`],
      ['Parcial', `=COUNTIF(${P}G2:G,D21)`]
    ]],
    ['A27', [['TOP 10 SALDOS EM ABERTO', ''], ['Nº Contrato', 'Cliente']]]
  ];

  blocos.forEach(([cel, vals]) => {
    const r = p.getRange(cel);
    p.getRange(r.getRow(), r.getColumn(), vals.length, 2).setValues(vals);
  });

  p.getRange('C28:D28').setValues([['Saldo', 'Status']]);
  p.getRange('A29').setFormula(
    `=IFERROR(QUERY(${C}A2:S,"select A, B, R, S where R > 0 and A <> 'ERRO' order by R desc limit 10",0),"Nenhum saldo em aberto")`);

  // Formatação
  p.getRange('A1').setFontSize(16).setFontWeight('bold');
  p.getRange('B2').setNumberFormat('dd/MM/yyyy HH:mm');
  ['A4:B4', 'A10:B10', 'D10:E10', 'A17:B17', 'D17:E17', 'A27:D28'].forEach(a =>
    p.getRange(a).setFontWeight('bold').setBackground('#263238').setFontColor('#fff'));
  ['B6', 'B21:B23', 'C29:C38'].forEach(a => p.getRange(a).setNumberFormat('R$ #,##0.00'));
  p.getRange('B14').setNumberFormat('0.0%');
  p.getRange('B23').setFontColor('#c62828').setFontWeight('bold');
  p.setColumnWidth(1, 220); p.setColumnWidth(4, 160);

  // Gráficos
  p.insertChart(p.newChart().setChartType(Charts.ChartType.PIE)
    .addRange(p.getRange('D11:E14')).setPosition(3, 7, 0, 0)
    .setOption('title', 'Recebimento de equipamentos').build());
  p.insertChart(p.newChart().setChartType(Charts.ChartType.PIE)
    .addRange(p.getRange('D18:E21')).setPosition(22, 7, 0, 0)
    .setOption('title', 'Status das parcelas').build());
}

// ==================== 4. ALERTA SEMANAL ====================
function enviarAlertaVencimentos() {
  const ss = SpreadsheetApp.getActive();
  const aba = ss.getSheetByName(CONFIG.ABA);
  const abaP = ss.getSheetByName(CONFIG.ABA_PARCELAS);
  if (!aba || aba.getLastRow() < 2) return;

  SpreadsheetApp.flush();
  const hoje = new Date(); hoje.setHours(0, 0, 0, 0);
  const dias = d => Math.round((d - hoje) / 86400000);
  const fmt = d => d instanceof Date ? Utilities.formatDate(d, CONFIG.TZ, 'dd/MM/yyyy') : '';
  const moeda = v => Number(v || 0).toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
  const esc = s => String(s ?? '').replace(/[&<>"]/g, m => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[m]));

  const contratos = aba.getRange(2, 1, aba.getLastRow() - 1, 19).getValues().filter(r => r[0] && r[0] !== 'ERRO');
  const link = {}; contratos.forEach(r => link[r[0]] = r[8]);

  const vencidos = [], proximos = [], equipPend = [];
  contratos.forEach(r => {
    if (r[4] instanceof Date) {
      const dd = dias(r[4]);
      const item = [r[0], r[1], r[2], fmt(r[4]), dd, moeda(r[17])];
      if (dd < 0 && r[18] !== 'Liquidado') vencidos.push(item);
      else if (dd >= 0 && dd <= CONFIG.DIAS_ALERTA) proximos.push(item);
    }
    if (['Pendente', 'Parcial', 'Excedente'].includes(r[14]) && r[3] instanceof Date && r[3] <= hoje)
      equipPend.push([r[0], r[1], r[2], r[11], r[12], r[13], r[14]]);
  });

  const parcAtras = [], parcProx = [];
  if (abaP && abaP.getLastRow() > 1) {
    abaP.getRange(2, 1, abaP.getLastRow() - 1, 9).getValues().forEach(r => {
      if (!(r[2] instanceof Date) || r[6] === 'Paga') return;
      const dd = dias(r[2]);
      const item = [r[0], r[1], fmt(r[2]), dd, moeda(r[8])];
      if (dd < 0) parcAtras.push(item);
      else if (dd <= CONFIG.DIAS_ALERTA) parcProx.push(item);
    });
  }

  if (![vencidos, proximos, equipPend, parcAtras, parcProx].some(l => l.length)) return;

  const tabela = (cols, linhas, cor) => `
    <table border="1" cellpadding="6" style="border-collapse:collapse;font-family:Arial;font-size:13px">
      <tr style="background:${cor};color:#fff">${cols.map(c => `<th>${c}</th>`).join('')}<th>PDF</th></tr>
      ${linhas.map(l => `<tr>${l.map(v => `<td>${esc(v)}</td>`).join('')}
        <td>${link[l[0]] ? `<a href="${link[l[0]]}">Abrir</a>` : ''}</td></tr>`).join('')}
    </table>`;
  const secao = (titulo, cor, cols, lista) => lista.length
    ? `<h3 style="color:${cor}">${titulo} (${lista.length})</h3>${tabela(cols, lista, cor)}` : '';

  const colC = ['Nº', 'Cliente', 'Equipamento', 'Vencimento', 'Dias', 'Saldo'];
  const colP = ['Nº', 'Parcela', 'Vencimento', 'Dias', 'Saldo'];
  let html = `<p>Resumo semanal de contratos – ${fmt(hoje)}</p>`;
  html += secao('🔴 Parcelas em atraso', '#c62828', colP, parcAtras);
  html += secao(`🟡 Parcelas a vencer em até ${CONFIG.DIAS_ALERTA} dias`, '#f9a825', colP, parcProx);
  html += secao('🔴 Contratos vencidos com saldo', '#c62828', colC, vencidos);
  html += secao(`🟡 Contratos que vencem em até ${CONFIG.DIAS_ALERTA} dias`, '#f9a825', colC, proximos);
  html += secao('📦 Equipamentos com divergência', '#1565c0',
    ['Nº', 'Cliente', 'Equipamento', 'Contratada', 'Recebida', 'Diferença', 'Status'], equipPend);
  html += `<p><a href="${ss.getUrl()}">Abrir planilha</a></p>`;

  MailApp.sendEmail({
    to: CONFIG.EMAIL,
    subject: `Contratos: ${parcAtras.length} parcela(s) atrasada(s), ${vencidos.length} vencido(s), ${equipPend.length} divergência(s)`,
    htmlBody: html
  });
  registrarHistorico('-', 'Sistema', 'Alerta semanal enviado', '', '',
    `Atrasadas: ${parcAtras.length} | Vencidos: ${vencidos.length} | Equip.: ${equipPend.length}`);
}

// ==================== 5. GATILHOS (rodar 1 vez) ====================
function criarGatilhos() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('processarNovosPDFs')
    .timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(7).create();
  ScriptApp.newTrigger('enviarAlertaVencimentos')
    .timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(9).create();
  ScriptApp.newTrigger('aoEditar')
    .forSpreadsheet(SpreadsheetApp.getActive()).onEdit().create();
}
