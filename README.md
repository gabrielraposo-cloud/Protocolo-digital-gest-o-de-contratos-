// PROTOCOLO DIGITAL – GESTÃO DE CONTRATOS |  (v7.0)

const CONFIG = {
  PASTA_ID: '1nAB3-CrOBI2dq7dmPq8eKQ942uqTFNSd',
  MODELO: 'gemini-2.5-flash',
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
    .addItem('1. Configurar Estrutura', 'configurarPlanilha')
    .addItem('2. Inserir Chave do Gemini (AIza...)', 'salvarChaveAPI')
    .addItem('3. Ativar Robô (15 min)', 'ativarGatilhosGlobais')
    .addSeparator()
    .addItem('▶️ 4. Forçar Leitura Agora', 'forcarProcessamento')
    .addToUi();
}

function salvarChaveAPI() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt('Segurança da Chave', 'Cole sua chave do Google AI Studio (DEVE começar com AIza...):', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() === ui.Button.OK) {
    const chave = resp.getResponseText().trim();
    if(!chave.startsWith('AIza')) {
      ui.alert('❌ Erro: Chave inválida. Ela deve começar com "AIza".');
      return;
    }
    PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', chave);
    ui.alert("✅ Chave do Gemini salva com segurança no PropertiesService!");
  }
}

function forcarProcessamento() {
  SpreadsheetApp.getActive().toast("Varrendo Drive e acionando IA. Aguarde...", "Protocolo", 5);
  const qtd = processarNovosPDFs();
  SpreadsheetApp.getUi().alert(`✅ Leitura concluída! ${qtd} arquivo(s) processado(s).\nErros estarão no Historico.`);
}

function configurarPlanilha() {
  const ss = SpreadsheetApp.getActive();
  const obter = n => ss.getSheetByName(n) || ss.insertSheet(n);
  const cab = (sh, h) => { sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold').setBackground('#263238').setFontColor('#fff').setWrap(true); sh.setFrozenRows(1); };
  const dv = () => SpreadsheetApp.newDataValidation().setAllowInvalid(false);

  const c = obter(CONFIG.ABA); cab(c, CAB.Contratos);
  Object.keys(FORM_C).forEach(l => c.getRange(`${l}2:${l}`).clearContent());
  Object.entries(FORM_C).forEach(([l, f]) => c.getRange(`${l}2`).setFormula(f));
  c.getRange('D2:E').setNumberFormat('dd/MM/yyyy'); c.getRange('J2:J').setNumberFormat('dd/MM/yyyy HH:mm');
  ['F2:F', 'V2:V'].forEach(a => c.getRange(a).setNumberFormat('R$ #,##0.00'));
  ['G2:G', 'I2:I'].forEach(a => c.getRange(a).setBackground('#e3f2fd'));
  ['H2:H', 'J2:J'].forEach(a => c.getRange(a).setBackground('#eceff1'));
  c.getRange('Q2:Q').setBackground('#fff8e1');
  c.getRange('I2:I').setDataValidation(dv().requireValueInList(STATUS_APROV, true).build());
  proteger(c, ['H2:H', 'J2:J', 'K2:L', 'R2:Y']);

  const p = obter(CONFIG.ABA_PARCELAS); cab(p, CAB.Parcelas);
  Object.keys(FORM_P).forEach(l => p.getRange(`${l}2:${l}`).clearContent());
  Object.entries(FORM_P).forEach(([l, f]) => p.getRange(`${l}2`).setFormula(f));
  ['C2:C', 'E2:E'].forEach(a => p.getRange(a).setNumberFormat('dd/MM/yyyy'));
  ['D2:D', 'F2:F', 'I2:I'].forEach(a => p.getRange(a).setNumberFormat('R$ #,##0.00'));
  p.getRange('E2:F').setBackground('#fff8e1');
  proteger(p, ['G2:I']);

  const a = obter(CONFIG.ABA_ALCADAS); cab(a, CAB.Alcadas);
  if (ultimaLinha(a, 1) < 2) a.getRange(2, 1, 3, 3).setValues([[0, 'Coordenação', 'coord@maino.com.br'], [10000, 'Gerência', 'gerente@maino.com.br'], [50000, 'Diretoria', 'diretor@maino.com.br']]);
  
  const h = obter(CONFIG.ABA_HIST); cab(h, CAB.Historico);
  h.getRange('A2:A').setNumberFormat('dd/MM/yyyy HH:mm:ss');
  SpreadsheetApp.getUi().alert('✅ Estrutura Master v7.0 aplicada!');
}

function proteger(sh, areas) { sh.getProtections(SpreadsheetApp.ProtectionType.RANGE).filter(x => x.getDescription() === 'PD-calc').forEach(x => x.remove()); areas.forEach(a => sh.getRange(a).protect().setDescription('PD-calc').setWarningOnly(true)); }
function ativarGatilhosGlobais() { ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t)); ScriptApp.newTrigger('processarNovosPDFs').timeBased().everyMinutes(15).create(); ScriptApp.newTrigger('aoEditar').forSpreadsheet(SpreadsheetApp.getActive()).onEdit().create(); SpreadsheetApp.getUi().alert('🚀 Robô Ativado para rodar a cada 15 min!'); }

// ==================== MOTOR DE IA (GEMINI API) ====================
function processarNovosPDFs() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return 0;
  let processados = 0;

  try {
    const ss = SpreadsheetApp.getActive();
    const aba = ss.getSheetByName(CONFIG.ABA);
    const abaP = ss.getSheetByName(CONFIG.ABA_PARCELAS);
    if (!aba || !abaP) return 0;

    const pastaOrigem = DriveApp.getFolderById(CONFIG.PASTA_ID);
    const pastasProcessados = pastaOrigem.getFoldersByName("Processados");
    const pastaDestino = pastasProcessados.hasNext() ? pastasProcessados.next() : pastaOrigem.createFolder("Processados");
    
    const arquivos = pastaOrigem.getFiles(); // FILTRO FLEXÍVEL: Pega todos e valida depois
    
    while (arquivos.hasNext()) {
      const arq = arquivos.next();
      const mime = arq.getMimeType();
      const nome = arq.getName().toLowerCase();
      
      // Valida se é PDF nativo ou se tem .pdf no nome (evita erro de octet-stream)
      if (mime === MimeType.PDF || nome.endsWith('.pdf')) {
        try {
          const d = extrairDadosGemini(arq);
          const id = d.numero_contrato || ('SN-' + arq.getId().slice(0, 8));
          const linha = ultimaLinha(aba, 1) + 1;
          
          aba.getRange(linha, 1, 1, 10).setValues([[
            id, d.cliente, d.equipamento, paraData(d.data_inicio), paraData(d.data_vencimento), 
            d.valor, d.solicitante || '', '', 'Pendente', ''
          ]]);
          aba.getRange(linha, 13, 1, 5).setValues([[
            arq.getUrl(), arq.getId(), new Date(), d.quantidade_total ?? '', 0 
          ]]);
          
          gerarParcelasIA(abaP, id, d);
          arq.moveTo(pastaDestino);
          registrarHistorico(id, CONFIG.ABA, 'Importado pelo Gemini', '', '', `Valor: R$ ${d.valor}`);
          processados++;
        } catch (e) {
          registrarHistorico(arq.getName(), CONFIG.ABA, 'Erro IA', '', '', e.message.slice(0, 200));
        }
      }
    }
  } finally {
    lock.releaseLock();
  }
  return processados;
}

function extrairDadosGemini(arquivo) {
  const chave = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!chave || !chave.startsWith('AIza')) throw new Error('Chave API ausente ou inválida. Insira via Menu > Inserir Chave.');
  
  const pdfBase64 = Utilities.base64Encode(arquivo.getBlob().getBytes());
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${CONFIG.MODELO}:generateContent?key=${chave}`;
  
  const prompt = 
    'Retorne APENAS um JSON válido. Regras: NF de compra não tem parcelas (repita emissão no venc. e array parcelas vazio). Locação sem vencimento expresso: some prazo à data inicio para o vencimento, valor total = prazo * valor mensal. Formato exato:\n' +
    '{"numero_contrato":"","cliente":"","equipamento":"","quantidade_total":0,"data_inicio":"AAAA-MM-DD","data_vencimento":"AAAA-MM-DD","valor":0,"solicitante":"","numero_parcelas":0,"valor_parcela":0,"data_primeira_parcela":"AAAA-MM-DD","parcelas":[{"numero":1,"vencimento":"AAAA-MM-DD","valor":0}]}';

  const payload = {
    "contents": [{
      "parts": [
        {"text": prompt},
        {"inlineData": {"mimeType": "application/pdf", "data": pdfBase64}}
      ]
    }],
    "generationConfig": {
      "responseMimeType": "application/json", // BLINDAGEM DE JSON NATIVA
      "temperature": 0
    }
  };

  const resp = UrlFetchApp.fetch(url, {
    method: 'post', contentType: 'application/json',
    payload: JSON.stringify(payload), muteHttpExceptions: true
  });
  
  if (resp.getResponseCode() !== 200) throw new Error(resp.getContentText());
  const jsonOutput = JSON.parse(resp.getContentText());
  let textoGemini = jsonOutput.candidates[0].content.parts[0].text;
  
  // Limpeza extra caso o Gemini ainda tente mandar Markdown
  textoGemini = textoGemini.replace(/```json/gi, '').replace(/```/g, '').trim();
  return JSON.parse(textoGemini);
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
  abaP.getRange(ini, 1, linhas.length, 6).setValues(linhas);
}

// ==================== WORKFLOW MANUAIS (ALÇADAS) ====================
function aoEditar(e) {
  if (!e || !e.range || e.range.getRow() < 2) return;
  const sh = e.range.getSheet(), nome = sh.getName(), user = ((e.user && e.user.getEmail()) || '').toLowerCase();
  const multi = e.range.getNumRows() > 1 || e.range.getNumColumns() > 1;
  if (nome === CONFIG.ABA) editarContrato(sh, e, user, multi);
  else if (nome === CONFIG.ABA_PARCELAS) editarParcela(sh, e, user, multi);
}

function editarContrato(sh, e, user, multi) {
  const r = e.range.getRow(), c = e.range.getColumn();
  if (multi) return;
  const num = sh.getRange(r, COL.NUM).getValue(), campo = sh.getRange(1, c).getValue(), novo = e.range.getValue();
  const bloquear = msg => { e.range.setValue(e.oldValue ?? ''); avisar(e.range, msg); registrarHistorico(num, sh.getName(), 'BLOQUEADO', campo, e.value, e.oldValue, user); };

  if (c === COL.APROV || c === COL.DT_APROV) return bloquear('Automático.');
  if (c === COL.ST_APROV) {
    if (novo === 'Aprovado') {
      const valor = Number(sh.getRange(r, COL.VALOR).getValue()), solic = String(sh.getRange(r, COL.SOLIC).getValue()).trim().toLowerCase();
      if (!num || !user || !(valor > 0) || !solic) return bloquear('Preencha Nº, Valor e Solicitante.');
      if (solic === user) return bloquear('Auto-aprovação proibida.');
      const faixa = buscarAlcada(valor);
      if (!faixa || !faixa.emails.includes(user)) return bloquear(faixa ? `Requer: ${faixa.emails.join(', ')}` : 'Alçada inválida.');
      sh.getRange(r, COL.APROV, 1, 2).setValues([[user, new Date()]]);
      registrarHistorico(num, sh.getName(), 'Aprovado', faixa.nivel, e.oldValue, 'R$ ' + valor.toFixed(2), user);
    } else {
      sh.getRange(r, COL.APROV, 1, 2).setValues([[novo === 'Reprovado' ? user : '', novo === 'Reprovado' ? new Date() : '']]);
      registrarHistorico(num, sh.getName(), 'Status alterado', campo, e.oldValue, novo, user);
    }
    return;
  }
  if ([COL.NUM, COL.VALOR, COL.SOLIC].includes(c) && sh.getRange(r, COL.ST_APROV).getValue() === 'Aprovado') {
    sh.getRange(r, COL.APROV, 1, 3).setValues([['', 'Pendente', '']]);
    avisar(e.range, 'Aprovação revogada por alteração de dados.');
  }
}

function editarParcela(sh, e, user, multi) {
  const r = e.range.getRow(), c = e.range.getColumn();
  if (multi) return;
  const num = sh.getRange(r, CP.NUM).getValue();
  if ((c === CP.DT_PAG || c === CP.VL_PAG) && e.range.getValue() !== '') {
    if (statusAprovacao(num) !== 'Aprovado') {
      e.range.setValue(e.oldValue ?? '');
      avisar(e.range, `Contrato não Aprovado.`);
    }
  }
}

function paraData(iso) { if (!iso || typeof iso !== 'string') return ''; const [a,m,d] = iso.split('-').map(Number); return a&&m&&d ? new Date(a,m-1,d) : ''; }
function addMeses(dt, n) { const r = new Date(dt.getFullYear(), dt.getMonth() + n, 1); r.setDate(Math.min(dt.getDate(), new Date(r.getFullYear(), r.getMonth() + 1, 0).getDate())); return r; }
function avisar(r, msg) { r.setNote(`⛔ ${msg}`); SpreadsheetApp.getActive().toast(msg, 'Erro', 8); }
function registrarHistorico(c, aba, ev, cp, a, d, u) { const h = SpreadsheetApp.getActive().getSheetByName(CONFIG.ABA_HIST); if(h) h.appendRow([new Date(), u||'sistema', c, aba, ev, cp, String(a??''), String(d??'')]); }
function ultimaLinha(sh, c) { const v = sh.getRange(1, c, sh.getMaxRows(), 1).getValues(); for (let i = v.length - 1; i >= 0; i--) if (v[i][0] !== '') return i + 1; return 1; }
function statusAprovacao(num) { const c = SpreadsheetApp.getActive().getSheetByName(CONFIG.ABA), cel = c.getRange('A2:A').createTextFinder(String(num)).matchEntireCell(true).findNext(); return cel ? c.getRange(cel.getRow(), COL.ST_APROV).getValue() : ''; }
function buscarAlcada(v) { const a = SpreadsheetApp.getActive().getSheetByName(CONFIG.ABA_ALCADAS), u = ultimaLinha(a, 1); if(u<2) return null; const f = a.getRange(2,1,u-1,3).getValues().filter(l=>l[0]!==''&&!isNaN(l[0])).map(l=>({min:Number(l[0]),nivel:l[1],emails:String(l[2]).toLowerCase().split(/[;,]/).map(s=>s.trim()).filter(Boolean)})).sort((x,y)=>y.min-x.min); return f.find(x=>v>=x.min) || null; }
