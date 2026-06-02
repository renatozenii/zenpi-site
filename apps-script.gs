// Cole este código em script.google.com → Novo projeto
// Depois: Implantar → Nova implantação → Tipo: App da Web
// Executar como: Eu | Quem tem acesso: Qualquer pessoa
// Copie a URL gerada para index.html e dev.html

function doGet(e) {
  const action = e.parameter.action;
  try {
    if (action === 'getVisits') return jsonOut(getVisits());
    if (action === 'getMessages') return jsonOut(getMessages());
  } catch(err) { return jsonOut({error: err.message}); }
  return jsonOut([]);
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.action === 'saveVisit') saveVisit(data);
    else if (data.action === 'updateVisit') updateVisit(data);
    else if (data.action === 'saveMessage') saveMessage(data);
    else if (data.action === 'markRead') markRead(data);
  } catch(err) {}
  return jsonOut({ok: true});
}

function jsonOut(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (name === 'Visitas') sheet.appendRow(['Session ID','Data/Hora','Dispositivo','OS','Navegador','Fonte','Referrer','Cidade','Região','País','Tempo (s)','Cliques']);
    if (name === 'Mensagens') sheet.appendRow(['ID','Data/Hora','Nome','WhatsApp','Ramo','Lida']);
  }
  return sheet;
}

function saveVisit(d) {
  getSheet('Visitas').appendRow([d.sessionId, new Date().toISOString(), d.device||'', d.os||'', d.browser||'', d.source||'', d.referrer||'', d.city||'', d.region||'', d.country||'', 0, '[]']);
}

function updateVisit(d) {
  const sheet = getSheet('Visitas');
  const vals = sheet.getDataRange().getValues();
  for (let i = 1; i < vals.length; i++) {
    if (vals[i][0] === d.sessionId) {
      if (d.timeOnSite !== undefined) sheet.getRange(i+1, 11).setValue(d.timeOnSite);
      if (d.clicks !== undefined) sheet.getRange(i+1, 12).setValue(JSON.stringify(d.clicks));
      break;
    }
  }
}

function saveMessage(d) {
  getSheet('Mensagens').appendRow([Date.now().toString(), new Date().toISOString(), d.nome||'', d.whatsapp||'', d.ramo||'', false]);
}

function markRead(d) {
  const sheet = getSheet('Mensagens');
  const vals = sheet.getDataRange().getValues();
  for (let i = 1; i < vals.length; i++) {
    if (vals[i][0] === d.id) { sheet.getRange(i+1, 6).setValue(true); break; }
  }
}

function getVisits() {
  const vals = getSheet('Visitas').getDataRange().getValues();
  return vals.slice(1).map(r => ({
    sessionId: r[0], created_at: r[1], device: r[2], os: r[3], browser: r[4],
    source: r[5], referrer: r[6], city: r[7], region: r[8], country: r[9],
    time_on_site: r[10],
    clicks: (() => { try { return JSON.parse(r[11]||'[]'); } catch(e) { return []; } })()
  }));
}

function getMessages() {
  const vals = getSheet('Mensagens').getDataRange().getValues();
  return vals.slice(1).map(r => ({
    id: r[0], created_at: r[1], nome: r[2], whatsapp: r[3], ramo: r[4], lida: r[5]
  }));
}
