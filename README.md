# Protocolo-digital-gestão-de-contratos-
Aplicação de gestão de contratos
Repósitorio criado para a automação do processo de gestão de contratos

// ============================================================
// PROTOCOLO DIGITAL – GESTÃO DE CONTRATOS | MASTER (v5.0)
// União: ArrayFormulas + Workflow Rigoroso + IA Claude 3.5 Autônoma
// ============================================================

const CONFIG = {
  PASTA_ID: 'COLE_O_ID_REAL_DA_PASTA_AQUI', // <-- ID da pasta do Drive onde os PDFs caem
  MODELO: 'claude-3-5-sonnet-20241022',
  ABA: 'Contratos', ABA_PARCELAS: 'Parcelas', ABA_HIST: 'Historico',
  ABA_PAINEL: 'Painel', ABA_ALCADAS: 'Alcadas', TZ: 'America/Sao_Paulo'
};

const STATUS_APROV = ['Pendente', 'Aprovado', 'Reprovado'];
const COL = { NUM: 1, FORN: 2, INI: 4, VALOR: 6, SOLIC: 7, APROV: 8, ST_APROV: 9, DT_APROV: 10, QTD_C: 16, QTD_R: 17 };
const CP = { NUM: 1, DT_PAG: 5, VL_PAG: 6 };

const CAB = {
  Contratos: ['Nº Contrato','Fornecedor','Equipamento','Início','Vencimento','Valor Total',
    'Solicitante (e-mail)','Aprovador','Status Aprovação','Data Aprovação',
    'Status','Dias p/ Vencer','Link PDF','File ID','Lido em',
    'Qtd Contratada','Qtd Recebida','Diferença','Status Recebimento',
    'Nº Parcelas','Parcelas Pagas','Saldo Pendente','Status Financeiro',
    'Alçada Exigida','Pendências'],
  Parcelas: ['Nº Contrato','Parcela','Vencimento','Valor','Data Pagamento','Valor Pago',
    'Status','Dias p/ Vencer','Saldo Parcela'],
  Historico: ['Data/Hora','Usuário','Nº Contrato','Aba','Evento','Campo','Valor Anterior','Valor Novo'],
  Alcadas: ['Valor Mínimo (R$)','Nível','E-mails dos Aprovadores (separar por ;)']
};

const FORM_C = {
  K: '=ARRAYFORMULA(IF(A2:A="","",IF(E2:E="","Sem data",IF(E2:E<TODAY(),"Vencido","No prazo"))))',
  L: '=ARRAYFORMULA(IF((A2:A="")+(E2:E=""),"",E2:E-TODAY()))',
  R: '=ARRAYFORMULA(IF((A2:A="")+(P2:P=""),"",P2:P-Q2:Q))',
  S: '=ARRAYFORMULA(IF(A2:A="","",IF(P2:P="","Sem qtd",IF(Q2:Q=0,"Pendente",IF(Q2:Q<P2:P,"Parcial",IF(Q2:Q=P2:P,"Completo","Excedente"))))))',
  T: '=ARRAYFORMULA(IF(A2:A="","",COUNTIF(Parcelas!A2:A,A2:A)))',
  U: '=ARRAYFORMULA(IF(A2:A="","",COUNTIFS(Parcelas!A2:A,A2:A,Parcelas!G2:G,"Paga")))',
  V: '=ARRAYFORMULA(IF(A2:A="","",IF(T2:T=0,F2:F,SUMIFS(Parcelas!I2:I,Parcelas!A2:A,A2:A))))',
  W: '=ARRAYFORMULA(IF(A2:A="","",IF(T2:T=0,"Sem parcelas",IF(V2:V<=0.009,"Liquidado",IF(COUNTIFS(Parcelas!A2:A,A2:A,Parcelas!G2:G,"Atrasada")>0,"Em atraso","Em dia")))))',
  X: '=ARRAYFORMULA(IF((A2:A="")+(F2:F=""),"",IFERROR(VLOOKUP(F2:F,SORT(FILTER(Alcadas!A2:B,Alcadas!A2:A<>""),1,TRUE),2,TRUE),"Sem faixa")))',
  Y: '=ARRAYFORMULA(IF(A2:A="","",REGEXREPLACE(IF(COUNTIF(A2:A,A2:A)>1,"Nº duplicado; ","")&IF(F2:F="","Sem valor; ","")&IF(G2:G="","Sem solicitante; ","")&IF(I2:I<>"Aprovado","Aguardando aprovação; ","")&IF((T2:T>0)*(ABS(SUMIFS(Parcelas!D2:D,Parcelas!A2:A,A2:A)-F2:F)>0.01),"Parcelas ≠ total; ","")&IF((S2:S="Pendente")+(S2:S="Parcial"),"Recebimento pendente; ",""),"; $","")))'
};
const FORM_P = {
  G: '=ARRAYFORMULA(IF(A2:A="","",IF((D2:D<>"")*(F2:F>=D2:D),"Paga",IF(C2:C="","Sem data",IF(C2:C<TODAY(),"Atrasada",IF(F2:F>0,"Parcial","A vencer"))))))',
  H: '=ARRAYFORMULA(IF((A2:A="")+(G2:G="Paga")+(C2:C=""),"",C2:C-TODAY()))',
  I: '=ARRAYFORMULA(IF(A2:A="","",IF(D2:D-F2:F>0,D2:D-F2:F,0)))'
};

// ==================== MENUS & SETUP ====================
function onOpen() {
  SpreadsheetApp.getUi().createMenu('⚙️ Protocolo Digital')
    .addItem('1. Configurar Estrutura (ArrayFormulas)', 'configurarPlanilha')
    .addItem('2. Ativar Automação IA e Bloqueios (1x)', 'ativarGatilhosGlobais')
    .addSeparator()
    .addItem('Gerar parcelas (Manual p/ Selecionado)', 'gerarParcelasSelecionado')
    .addItem('Recriar Painel', 'recriarPainel')
    .addToUi();
}

function salvarChaveAPI() {
  PropertiesService.getScriptProperties().setProperty('ANTHROPIC_KEY', 'COLE_SUA_CHAVE_AQUI');
  SpreadsheetApp.getUi().alert("Chave API salva com segurança no cofre do Google!");
}

function configurarPlanilha() {
  const ss = SpreadsheetApp.getActive();
  const obter = n => ss.getSheetByName(n) || ss.insertSheet(n);
  const cab = (sh, h) => {
    sh.getRange(1, 1, 1, h.length).setValues([h])
      .setFontWeight('bold').setBackground('#263238').setFontColor('#fff').setWrap(true);
    sh.setFrozenRows(1);
  };
  const dv = () => SpreadsheetApp.newDataValidation().setAllowInvalid(false);

  // --- CONTRATOS ---
  const c = obter(CONFIG.ABA);
  cab(c, CAB.Contratos);
  Object.keys(FORM_C).forEach(l => c.getRange(`${l}2:${l}`).clearContent());
  Object.entries(FORM_C).forEach(([l, f]) => c.getRange(`${l}2`).setFormula(f));

  c.getRange('D2:E').setNumberFormat('dd/MM/yyyy');
  c.getRange('J2:J').setNumberFormat('dd/MM/yyyy HH:mm');
  ['F2:F', 'V2:V'].forEach(a => c.getRange(a).setNumberFormat('R$ #,##0.00'));
  c.getRange('A2:Y').setBackground(null);
  ['G2:G', 'I2:I'].forEach(a => c.getRange(a).setBackground('#e3f2fd'));
  ['H2:H', 'J2:J'].forEach(a => c.getRange(a).setBackground('#eceff1'));
  c.getRange('Q2:Q').setBackground('#fff8e1');

  c.getRange('I2:I').setDataValidation(dv().requireValueInList(STATUS_APROV, true).build());
  c.getRange('D2:E').setDataValidation(dv().requireDate().build());
  c.getRange('F2:F').setDataValidation(dv().requireNumberGreaterThanOrEqualTo(0).build());
  c.getRange('P2:Q').setDataValidation(dv().requireNumberGreaterThanOrEqualTo(0).build());
  proteger(c, ['H2:H', 'J2:J', 'K2:L', 'R2:Y']);

  // --- PARCELAS ---
  const p = obter(CONFIG.ABA_PARCELAS);
  cab(p, CAB.Parcelas);
  Object.keys(FORM_P).forEach(l => p.getRange(`${l}2:${l}`).clearContent());
  Object.entries(FORM_P).forEach(([l, f]) => p.getRange(`${l}2`).setFormula(f));
  ['C2:C', 'E2:E'].forEach(a => p.getRange(a).setNumberFormat('dd/MM/yyyy'));
  ['D2:D', 'F2:F', 'I2:I'].forEach(a => p.getRange(a).setNumberFormat('R$ #,##0.00'));
  p.getRange('A2:I').setBackground(null);
  p.getRange('E2:F').setBackground('#fff8e1');
  p.getRange('C2:C').setDataValidation(dv().requireDate().build());
  p.getRange('E2:E').setDataValidation(dv().requireDate().build());
  p.getRange('D2:D').setDataValidation(dv().requireNumberGreaterThanOrEqualTo(0).build());
  p.getRange('F2:F').setDataValidation(dv().requireNumberGreaterThanOrEqualTo(0).build());
  proteger(p, ['G2:I']);

  // --- ALÇADAS ---
  const a = obter(CONFIG.ABA_ALCADAS);
  cab(a, CAB.Alcadas);
  if (ultimaLinha(a, 1) < 2) a.getRange(2, 1, 3, 3).setValues([
    [0, 'Nível 1 – Coordenação', 'coordenador@maino.com.br'],
    [10000, 'Nível 2 – Gerência', 'gerente@maino.com.br'],
    [50000, 'Nível 3 – Diretoria', 'diretor@maino.com.br']
  ]);
  a.getRange('A2:A').setNumberFormat('R$ #,##0.00');
  a.setColumnWidth(3, 380);

  // --- HISTÓRICO ---
  const h = obter(CONFIG.ABA_HIST);
  cab(h, CAB.Historico);
  h.getRange('A2:A').setNumberFormat('dd/MM/yyyy HH:mm:ss');
  h.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(x => x.remove());
  const prot = h.protect().setDescription('PD-hist');
  prot.removeEditors(prot.getEditors());
  if (prot.canDomainEdit()) prot.setDomainEdit(false);

  montarPainel(obter(CONFIG.ABA_PAINEL));
  registrarHistorico('-', 'Sistema', 'Estrutura configurada (v5.0)', '', '', '');
  SpreadsheetApp.getUi().alert('✅ Estrutura Master v5.0 aplicada.\nArrayFormulas ativadas.');
}

function proteger(sh, areas) {
  sh.getProtections(SpreadsheetApp.ProtectionType.RANGE).filter(x => x.getDescription() === 'PD-calc').forEach(x => x.remove());
  areas.forEach(a => sh.getRange(a).protect().setDescription('PD-calc').setWarningOnly(true));
}

function ativarGatilhosGlobais() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('processarNovosPDFs').timeBased().everyMinutes(15).create();
  ScriptApp.newTrigger('aoEditar').forSpreadsheet(SpreadsheetApp.getActive()).onEdit().create();
  SpreadsheetApp.getUi().alert('✅ IA e Workflow Ativados!\nA IA buscará PDFs a cada 15m. O bloqueio de aprovações/edições manuais está rodando.');
}

// ==================== MOTOR DE IA (CLAUDE) ====================
function processarNovosPDFs() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return;

  try {
    const ss = SpreadsheetApp.getActive();
    const aba = ss.getSheetByName(CONFIG.ABA);
    const abaP = ss.getSheetByName(CONFIG.ABA_PARCELAS);
    if (!aba || !abaP) return;

    const pastaOrigem = DriveApp.getFolderById(CONFIG.PASTA_ID);
    const pastasProcessados = pastaOrigem.getFoldersByName("Processados");
    const pastaDestino = pastasProcessados.hasNext() ? pastasProcessados.next() : pastaOrigem.createFolder("Processados");

    const arquivos = pastaOrigem.getFilesByType(MimeType.PDF);
    
    while (arquivos.hasNext()) {
      const arq = arquivos.next();

      try {
        const d = extrairDados(arq);
        const id = d.numero_contrato || ('SN-' + arq.getId().slice(0, 8));
        const linha = aba.getLastRow() + 1;
        
        // Pula colunas da ARRAYFORMULA para não quebrar a planilha, insere as pontas
        aba.getRange(linha, 1, 1, 10).setValues([[
          id, d.cliente, d.equipamento, paraData(d.data_inicio), paraData(d.data_vencimento), 
          d.valor, d.solicitante || '', '', 'Pendente', ''
        ]]);
        aba.getRange(linha, 13, 1, 5).setValues([[
          arq.getUrl(), arq.getId(), new Date(), d.quantidade_total ?? '', 0 
        ]]);
        
        gerarParcelasIA(abaP, id, d);
        arq.moveTo(pastaDestino);
        registrarHistorico(id, CONFIG.ABA, 'Importado por IA', '', '', `Valor: R$ ${d.valor}`);
      
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
  const prompt =
    'Você é um analista financeiro. Extraia dados do Contrato ou NF anexa em JSON exato: ' +
    '{"numero_contrato":"","cliente":"","equipamento":"","quantidade_total":0,' +
    '"data_inicio":"AAAA-MM-DD","data_vencimento":"AAAA-MM-DD","valor":0,"solicitante":"",' +
    '"numero_parcelas":0,"valor_parcela":0,"data_primeira_parcela":"AAAA-MM-DD",' +
    '"parcelas":[{"numero":1,"vencimento":"AAAA-MM-DD","valor":0}]}. ' +
    '1. numero_contrato: NF = Num da NF. Aditivo = Num do Aditivo. ' +
    '2. equipamento: Agrupe. quantidade_total: Soma total. ' +
    '3. valor: VALOR TOTAL. Se contrato de locação com prazo (ex: 48 meses), MULTIPLIQUE valor mensal pelo prazo. ' +
    '4. datas: Se NF de compra, repita a emissão no vencimento. Se locação sem vencimento, SOME os meses. ' +
    '5. solicitante: Destinatário ou signatário eletrônico. ' +
    '6. parcelas: NF à vista ou sem detalhe = [] e 0 parcelas. ' +
    'Responda APENAS JSON. Use null para o que não achar.';

  const resp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post', contentType: 'application/json',
    headers: { 'x-api-key': chave, 'anthropic-version': '2023-06-01', 'anthropic-beta': 'pdfs-2024-09-25' },
    payload: JSON.stringify({
      model: CONFIG.MODELO, max_tokens: 1500, temperature: 0,
      messages: [{ role: 'user', content: [
        { type: 'document', source: { type: 'base64', media_type: 'application/pdf', data: pdf } },
        { type: 'text', text: prompt }
      ]}]
    }),
    muteHttpExceptions: true
  });
  if (resp.getResponseCode() !== 200) throw new Error(resp.getContentText());
  const texto = JSON.parse(resp.getContentText()).content[0].text;
  return JSON.parse(texto.replace(/```json|```/g, '').trim());
}

function gerarParcelasIA(abaP, id, d) {
  let linhas = [];
  if (Array.isArray(d.parcelas) && d.parcelas.length) {
    linhas = d.parcelas.map((p, i) => [id, p.numero || i + 1, paraData(p.vencimento), Number(p.valor) || '', '', '']);
  } else if (Number(d.numero_parcelas) > 0) {
    const n = Number(d.numero_parcelas), total = Number(d.valor) || 0;
    const vp = Number(d.valor_parcela) || (total ? total / n : 0);
    const base = paraData(d.data_primeira_parcela) || paraData(d.data_inicio);
    if (!base) return;

    for (let i = 0; i < n; i++) {
      let v = (i === n - 1 && total) ? (total - vp * (n - 1)) : vp;
      linhas.push([id, i + 1, addMeses(base, i), v || '', '', '']);
    }
  }
  if (!linhas.length) return;
  const ini = ultimaLinha(abaP, 1) + 1;
  abaP.getRange(ini, 1, linhas.length, 6).setValues(linhas); // Seta só A-F para não quebrar ArrayFormula G-I
}

// ==================== WORKFLOW, ALÇADAS E BLOQUEIOS (MANUAIS) ====================
function aoEditar(e) {
  if (!e || !e.range || e.range.getRow() < 2) return;
  const sh = e.range.getSheet(), nome = sh.getName();
  const user = ((e.user && e.user.getEmail()) || '').toLowerCase();
  const multi = e.range.getNumRows() > 1 || e.range.getNumColumns() > 1;

  if (nome === CONFIG.ABA) editarContrato(sh, e, user, multi);
  else if (nome === CONFIG.ABA_PARCELAS) editarParcela(sh, e, user, multi);
  else if (nome === CONFIG.ABA_ALCADAS)
    registrarHistorico('-', nome, 'Alçada alterada', e.range.getA1Notation(), e.oldValue, e.value, user);
}

function editarContrato(sh, e, user, multi) {
  const r = e.range.getRow(), c = e.range.getColumn();
  if (multi) {
    const c2 = c + e.range.getNumColumns() - 1, n = e.range.getNumRows();
    if (c <= COL.DT_APROV && c2 >= COL.APROV) {
      sh.getRange(r, COL.APROV, n, 3).setValues(Array.from({ length: n }, () => ['', 'Pendente', '']));
      avisar(e.range, 'Aprovação em lote bloqueada. Aprove linha a linha.');
      registrarHistorico('-', sh.getName(), 'BLOQUEADO: lote', '', '', 'Revertido', user);
    }
    return;
  }

  const num = sh.getRange(r, COL.NUM).getValue();
  const campo = sh.getRange(1, c).getValue(), novo = e.range.getValue();
  const bloquear = msg => { e.range.setValue(e.oldValue ?? ''); avisar(e.range, msg); registrarHistorico(num, sh.getName(), 'BLOQUEADO', campo, e.value, e.oldValue, user); };

  if (c === COL.APROV || c === COL.DT_APROV) return bloquear('Automático.');
  if (c === COL.ST_APROV) {
    if (novo === 'Aprovado') {
      const valor = Number(sh.getRange(r, COL.VALOR).getValue()), solic = String(sh.getRange(r, COL.SOLIC).getValue()).trim().toLowerCase();
      if (!num) return bloquear('Informe o Nº.');
      if (!user) return bloquear('Login não identificado.');
      if (!(valor > 0)) return bloquear('Sem Valor Total.');
      if (!solic) return bloquear('Informe solicitante.');
      if (solic === user) return bloquear('Você não pode aprovar seu próprio contrato.');
      const faixa = buscarAlcada(valor);
      if (!faixa || !faixa.emails.includes(user)) return bloquear(faixa ? `${faixa.nivel} exige: ${faixa.emails.join(', ')}` : 'Alçada não cadastrada.');
      
      sh.getRange(r, COL.APROV).setValue(user);
      sh.getRange(r, COL.DT_APROV).setValue(new Date());
      e.range.clearNote();
      registrarHistorico(num, sh.getName(), 'Aprovado', faixa.nivel, e.oldValue, 'R$ ' + valor.toFixed(2), user);
    } else {
      const rep = novo === 'Reprovado';
      sh.getRange(r, COL.APROV).setValue(rep ? user : '');
      sh.getRange(r, COL.DT_APROV).setValue(rep ? new Date() : '');
      registrarHistorico(num, sh.getName(), 'Status alterado', campo, e.oldValue, novo, user);
    }
    return;
  }
  if ([COL.NUM, COL.VALOR, COL.SOLIC].includes(c) && sh.getRange(r, COL.ST_APROV).getValue() === 'Aprovado') {
    sh.getRange(r, COL.APROV, 1, 3).setValues([['', 'Pendente', '']]);
    avisar(e.range, 'Dado-chave alterado. Aprovação revogada.');
    registrarHistorico(num, sh.getName(), 'Revogada', campo, e.oldValue, e.range.getDisplayValue(), user);
    return;
  }
  if ([1,2,3,4,5,6,7,13,16,17].includes(c)) registrarHistorico(num, sh.getName(), 'Alteração', campo, e.oldValue, e.range.getDisplayValue(), user);
}

function editarParcela(sh, e, user, multi) {
  const r = e.range.getRow(), c = e.range.getColumn();
  if (multi) return;
  const num = sh.getRange(r, CP.NUM).getValue(), campo = sh.getRange(1, c).getValue(), parc = sh.getRange(r, 2).getValue();

  if ((c === CP.DT_PAG || c === CP.VL_PAG) && e.range.getValue() !== '') {
    const st = statusAprovacao(num);
    if (st !== 'Aprovado') {
      e.range.setValue(e.oldValue ?? '');
      avisar(e.range, `Contrato ${num} não está Aprovado.`);
      registrarHistorico(num, sh.getName(), `BLOQUEADO pgto parcela ${parc}`, campo, e.value, '', user);
      return;
    }
  }
  if (c <= 6) registrarHistorico(num, sh.getName(), `Alt parcela ${parc}`, campo, e.oldValue, e.range.getDisplayValue(), user);
}

// ==================== AUXILIARES & PAINEL ====================
function paraData(iso) { if (!iso || typeof iso !== 'string') return ''; const [a,m,d] = iso.split('-').map(Number); return a&&m&&d ? new Date(a,m-1,d) : ''; }
function addMeses(dt, n) { const r = new Date(dt.getFullYear(), dt.getMonth() + n, 1); r.setDate(Math.min(dt.getDate(), new Date(r.getFullYear(), r.getMonth() + 1, 0).getDate())); return r; }
function avisar(r, msg) { r.setNote(`⛔ ${msg}`); SpreadsheetApp.getActive().toast(msg, 'Protocolo', 8); }
function registrarHistorico(c, aba, ev, cp, a, d, u) { const h = SpreadsheetApp.getActive().getSheetByName(CONFIG.ABA_HIST); if(h) h.appendRow([new Date(), u||Session.getActiveUser().getEmail(), c, aba, ev, cp, String(a??''), String(d??'')]); }
function ultimaLinha(sh, c) { const v = sh.getRange(1, c, sh.getMaxRows(), 1).getValues(); for (let i = v.length - 1; i >= 0; i--) if (v[i][0] !== '') return i + 1; return 1; }
function statusAprovacao(num) { const c = SpreadsheetApp.getActive().getSheetByName(CONFIG.ABA), cel = c.getRange('A2:A').createTextFinder(String(num)).matchEntireCell(true).findNext(); return cel ? c.getRange(cel.getRow(), COL.ST_APROV).getValue() : ''; }
function buscarAlcada(v) { const a = SpreadsheetApp.getActive().getSheetByName(CONFIG.ABA_ALCADAS), u = ultimaLinha(a, 1); if(u<2) return null; const f = a.getRange(2,1,u-1,3).getValues().filter(l=>l[0]!==''&&!isNaN(l[0])).map(l=>({min:Number(l[0]),nivel:l[1],emails:String(l[2]).toLowerCase().split(/[;,]/).map(s=>s.trim()).filter(Boolean)})).sort((x,y)=>y.min-x.min); return f.find(x=>v>=x.min) || null; }

function gerarParcelasSelecionado() {
  const ss = SpreadsheetApp.getActive(), ui = SpreadsheetApp.getUi(), sh = ss.getActiveSheet();
  if (sh.getName() !== CONFIG.ABA) return ui.alert('Selecione a linha na aba Contratos.');
  const r = sh.getActiveRange().getRow(); if (r < 2) return;
  const [num, , , ini, , total] = sh.getRange(r, 1, 1, 6).getValues()[0];
  if (!num || !(ini instanceof Date) || !(Number(total) > 0)) return ui.alert('Preencha Nº, Início e Valor Total.');
  
  const p = ss.getSheetByName(CONFIG.ABA_PARCELAS), ult = ultimaLinha(p, 1);
  if (ult > 1 && p.getRange(2, 1, ult - 1, 1).getValues().flat().map(String).includes(String(num))) return ui.alert('Já possui parcelas.');
  const resp = ui.prompt('Gerar', `Contrato ${num} - Total R$ ${Number(total).toFixed(2)}\nQuantas parcelas?`, ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const n = parseInt(resp.getResponseText(), 10);
  if (!(n > 0)) return;

  const vp = Math.floor(total / n * 100) / 100, linhas = Array.from({ length: n }, (_, i) => [num, i+1, addMeses(ini, i+1), i===n-1 ? Math.round((total-vp*(n-1))*100)/100 : vp, '', '']);
  p.getRange(ult + 1, 1, n, 6).setValues(linhas);
  ui.alert(`✅ ${n} parcelas geradas.`);
}

function recriarPainel() { montarPainel(SpreadsheetApp.getActive().getSheetByName(CONFIG.ABA_PAINEL)); }
function montarPainel(p) {
  p.clear(); p.getCharts().forEach(ch => p.removeChart(ch));
  const C = `'${CONFIG.ABA}'!`, P = `'${CONFIG.ABA_PARCELAS}'!`;
  const blocos = [
    ['A1', [['PAINEL GERENCIAL', '']]], ['A2', [['Atualizado em', '=NOW()']]],
    ['A4', [['CONTRATOS', ''], ['Total de contratos', `=COUNTA(${C}A2:A)`], ['Aguardando aprovação', `=COUNTIFS(${C}A2:A,"<>",${C}I2:I,"<>Aprovado",${C}I2:I,"<>Reprovado")`], ['Contratos vencidos', `=COUNTIF(${C}K2:K,"Vencido")`], ['Valor cadastrado', `=SUM(${C}F2:F)`]]],
    ['A11', [['EQUIPAMENTOS', ''], ['Qtd contratada', `=SUM(${C}P2:P)`], ['Qtd recebida', `=SUM(${C}Q2:Q)`], ['Pendente de receb.', `=SUMIF(${C}R2:R,">0")`]]],
    ['D11', [['Status recebimento', 'Contratos'], ['Completo', `=COUNTIF(${C}S2:S,D12)`], ['Pendente', `=COUNTIF(${C}S2:S,D14)`]]],
    ['A17', [['FINANCEIRO', ''], ['Total de parcelas', `=COUNTA(${P}A2:A)`], ['Valor pago', `=SUM(${P}F2:F)`], ['Saldo em aberto', `=SUM(${C}V2:V)`], ['Valor em atraso', `=SUMIF(${P}G2:G,"Atrasada",${P}I2:I)`]]],
    ['D17', [['Status parcelas', 'Qtd'], ['Paga', `=COUNTIF(${P}G2:G,D18)`], ['A vencer', `=COUNTIF(${P}G2:G,D19)`], ['Atrasada', `=COUNTIF(${P}G2:G,D20)`]]]
  ];
  blocos.forEach(([cel, vals]) => { const r = p.getRange(cel); p.getRange(r.getRow(), r.getColumn(), vals.length, 2).setValues(vals); });
  p.getRange('A23:D24').setValues([['TOP 10 SALDOS EM ABERTO', '', '', ''], ['Nº Contrato', 'Fornecedor', 'Saldo', 'Status']]);
  p.getRange('A25').setFormula(`=IFERROR(QUERY(${C}A2:Y,"select A, B, V, W where A is not null and V > 0 order by V desc limit 10",0),"Nenhum")`);
  p.getRange('A1').setFontSize(16).setFontWeight('bold');
  ['A4:B4', 'A11:B11', 'D11:E11', 'A17:B17', 'D17:E17', 'A23:D24'].forEach(a => p.getRange(a).setFontWeight('bold').setBackground('#263238').setFontColor('#fff'));
  ['B8:B9', 'B19:B21', 'C25:C34'].forEach(a => p.getRange(a).setNumberFormat('R$ #,##0.00'));
  p.setColumnWidth(1, 230); p.setColumnWidth(4, 160);
}
