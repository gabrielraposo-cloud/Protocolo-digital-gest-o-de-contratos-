# Protocolo-digital-gestão-de-contratos-
Aplicação de gestão de contratos
Repósitorio criado para a automação do processo de gestão de contratos


// ===================== CONFIGURAÇÃO =====================
const CONFIG = {
  PASTA_ID: 'Protocolo Digital- Gestão de Contratos',
  EMAIL: 'administrativo@maino.com.br',
  DIAS_ALERTA: 7,
  MODELO: 'claude-haiku-4-5' // confirme o nome vigente na documentação
};

// Rode UMA vez para salvar a chave com segurança (depois apague a chave daqui)
function salvarChave() {
  PropertiesService.getScriptProperties()
    .setProperty('ANTHROPIC_KEY', 'COLE_SUA_CHAVE_AQUI');
}

// ===================== 1. LEITURA DOS PDFs =====================
function processarNovosPDFs() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return; // evita execuções simultâneas

  try {
    const aba = SpreadsheetApp.getActive().getSheetByName(CONFIG.ABA);
    const ultima = aba.getLastRow();
    const jaLidos = ultima > 1
      ? aba.getRange(2, 10, ultima - 1, 1).getValues().flat()
      : [];

    const arquivos = DriveApp.getFolderById(CONFIG.PASTA_ID)
      .getFilesByType(MimeType.PDF);
    const inicio = Date.now();

    while (arquivos.hasNext()) {
      if (Date.now() - inicio > 4.5 * 60 * 1000) break; // margem do limite de 6 min
      const arq = arquivos.next();
      if (jaLidos.includes(arq.getId())) continue;

      try {
        const d = extrairDados(arq);
        const linha = aba.getLastRow() + 1;
        aba.getRange(linha, 1, 1, 11).setValues([[
          d.numero_contrato, d.cliente, d.equipamento,
          paraData(d.data_inicio), paraData(d.data_vencimento), d.valor,
          '', '', arq.getUrl(), arq.getId(), new Date()
        ]]);
        // Status e dias calculados por fórmula (atualizam sozinhos todo dia)
        aba.getRange(linha, 7).setFormula(
          `=IF(E${linha}="","Sem data",IF(E${linha}<TODAY(),"Vencido","No Prazo"))`);
        aba.getRange(linha, 8).setFormula(
          `=IF(E${linha}="","",E${linha}-TODAY())`);
      } catch (e) {
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
  const pdf = Utilities.base64Encode(arquivo.getBlob().getBytes());

  const prompt =
    'Você extrai dados de contratos de locação de equipamentos. ' +
    'Responda SOMENTE com JSON válido, sem texto extra, neste formato: ' +
    '{"numero_contrato":"","cliente":"","equipamento":"",' +
    '"data_inicio":"AAAA-MM-DD","data_vencimento":"AAAA-MM-DD","valor":0}. ' +
    'valor = valor total do contrato, número sem R$ e com ponto decimal. ' +
    'Se houver vários equipamentos, liste-os separados por "; ". ' +
    'Se um dado não existir, use null. Não invente informações.';

  const resp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-api-key': chave, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify({
      model: CONFIG.MODELO,
      max_tokens: 1024,
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

function paraData(iso) {
  if (!iso) return '';
  const [a, m, d] = iso.split('-').map(Number);
  return new Date(a, m - 1, d);
}

// ===================== 2. ALERTA DE VENCIMENTO =====================
function enviarAlertaVencimentos() {
  const aba = SpreadsheetApp.getActive().getSheetByName(CONFIG.ABA);
  if (aba.getLastRow() < 2) return;
  const dados = aba.getRange(2, 1, aba.getLastRow() - 1, 9).getValues();

  const hoje = new Date(); hoje.setHours(0, 0, 0, 0);
  const vencidos = [], proximos = [];

  dados.forEach(r => {
    if (!(r[4] instanceof Date)) return;
    const dias = Math.round((r[4] - hoje) / 86400000);
    const item = { num: r[0], cli: r[1], eq: r[2], venc: r[4], valor: r[5], link: r[8], dias };
    if (dias < 0) vencidos.push(item);
    else if (dias <= CONFIG.DIAS_ALERTA) proximos.push(item);
  });

  if (!vencidos.length && !proximos.length) return; // não envia e-mail vazio

  const fmt = d => Utilities.formatDate(d, 'America/Sao_Paulo', 'dd/MM/yyyy');
  const moeda = v => Number(v || 0).toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
  const tabela = (lista, cor) => `
    <table border="1" cellpadding="6" style="border-collapse:collapse;font-family:Arial;font-size:13px">
      <tr style="background:${cor};color:#fff">
        <th>Nº</th><th>Cliente</th><th>Equipamento</th><th>Vencimento</th><th>Dias</th><th>Valor</th><th>PDF</th>
      </tr>
      ${lista.map(i => `<tr><td>${i.num}</td><td>${i.cli}</td><td>${i.eq}</td>
        <td>${fmt(i.venc)}</td><td>${i.dias}</td><td>${moeda(i.valor)}</td>
        <td><a href="${i.link}">Abrir</a></td></tr>`).join('')}
    </table>`;

  let html = `<p>Resumo de contratos em ${fmt(hoje)}:</p>`;
  if (vencidos.length) html += `<h3 style="color:#c62828">🔴 Vencidos (${vencidos.length})</h3>${tabela(vencidos, '#c62828')}`;
  if (proximos.length) html += `<h3 style="color:#f9a825">🟡 Vencem em até ${CONFIG.DIAS_ALERTA} dias (${proximos.length})</h3>${tabela(proximos, '#f9a825')}`;
  html += `<p><a href="${SpreadsheetApp.getActive().getUrl()}">Abrir planilha</a></p>`;

  MailApp.sendEmail({
    to: CONFIG.EMAIL,
    subject: `Contratos: ${vencidos.length} vencido(s), ${proximos.length} a vencer`,
    htmlBody: html
  });
}

// ===================== 3. GATILHOS (rodar 1 vez) =====================
function criarGatilhos() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('processarNovosPDFs').timeBased().everyHours(1).create();
  ScriptApp.newTrigger('enviarAlertaVencimentos').timeBased().everyDays(1).atHour(8).create();
}
