/** @OnlyCurrentDoc */
// Recebe os leads da LP e grava na aba "Leads" da planilha.
// Implantar > Nova implantação > App da Web
//   Executar como: Eu | Quem pode acessar: Qualquer pessoa
// A URL /exec gerada vai em LEADS_ENDPOINT no lead.js.

const ABA = 'Leads';
const HORAS_SEM_DUPLICAR = 24; // mesmo telefone dentro desse prazo não vira linha nova

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    return ContentService.createTextOutput(gravarLead(d));
  } finally {
    lock.releaseLock();
  }
}

// Abrir a URL /exec no navegador mostra esta mensagem (serve pra conferir se está no ar)
function doGet() {
  return ContentService.createTextOutput('Endpoint de leads funcionando ✅');
}

function gravarLead(d) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA);
  if (!sheet) return 'aba Leads não encontrada';

  if (jaRegistrado(sheet, d.telefone)) return 'duplicado';

  const origem = (d.utm_source || '').toLowerCase();
  let plataforma = 'Orgânico';
  if (d.gclid || origem === 'google') plataforma = 'Google';
  else if (d.fbclid || ['facebook', 'instagram', 'meta', 'fb', 'ig'].includes(origem)) plataforma = 'Meta';
  else if (origem) plataforma = 'Outro';

  const obs = [d.interesse, d.pagina ? 'Página: ' + d.pagina : ''].filter(Boolean).join(' | ');

  sheet.appendRow([
    new Date(),                 // A Data e hora
    d.nome || '',               // B Nome
    d.telefone || '',           // C Telefone
    plataforma,                 // D Plataforma
    'Site / LP',                // E Captura
    '', '', '',                 // F Quando precisa, G Pessoas, H Motivo (a recepção preenche)
    d.utm_source || '',         // I utm_source
    d.utm_medium || '',         // J utm_medium
    d.utm_campaign || '',       // K utm_campaign
    d.utm_content || '',        // L utm_content
    d.utm_term || '',           // M utm_term
    d.gclid || d.fbclid || '',  // N Click ID
    'Novo',                     // O Status
    '', '', '',                 // P Data do pernoite, Q Valor, R Responsável
    obs                         // S Observações
  ]);
  sheet.getRange(sheet.getLastRow(), 1).setNumberFormat('dd/MM/yyyy HH:mm');
  return 'ok';
}

function jaRegistrado(sheet, telefone) {
  const ultima = sheet.getLastRow();
  if (!telefone || ultima < 2) return false;
  const inicio = Math.max(2, ultima - 300);
  const linhas = sheet.getRange(inicio, 1, ultima - inicio + 1, 3).getValues();
  const limite = Date.now() - HORAS_SEM_DUPLICAR * 3600 * 1000;
  const alvo = String(telefone).replace(/\D/g, '');
  return linhas.some(r => r[0] instanceof Date && r[0].getTime() > limite &&
    String(r[2]).replace(/\D/g, '') === alvo);
}

// Teste manual: selecione "testar" no editor e clique em Executar.
function testar() {
  const r = gravarLead({
    nome: 'Teste',
    telefone: '(17) 90000-0000',
    interesse: 'Teste do Apps Script',
    pagina: '/teste',
    utm_source: 'facebook',
    utm_medium: 'cpc',
    utm_campaign: 'teste'
  });
  Logger.log(r);
}
