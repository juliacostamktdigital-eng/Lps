/**
 * V4 Company — Cockpit Operacional
 * Google Apps Script: proxy de leitura das planilhas GrowthPack e Comercial
 *
 * COMO USAR:
 * 1. Abra script.google.com
 * 2. Crie um novo projeto
 * 3. Cole todo este código
 * 4. Clique em "Implantar" → "Nova implantação"
 * 5. Tipo: App da Web
 * 6. Executar como: Eu (sua conta)
 * 7. Quem pode acessar: Qualquer pessoa
 * 8. Clique em "Implantar" e copie a URL gerada
 * 9. No painel HTML, substitua APPS_SCRIPT_URL pela URL copiada
 *
 * DEPOIS DE FAZER ISSO:
 * No painel/index.html, edite a linha:
 *   const APPS_SCRIPT_URL = '';
 * E coloque a URL do Apps Script.
 * O CORS vai funcionar automaticamente.
 */

const SHEET_IDS = {
  FCE_PACE:     '1naAmFNzrycM5cddVQDhGduDQx0nZFe0GlgtajOA2fUI',
  FCE_COM:      '1duCquoJj-NEf2sGF_25oD6lHWZegCOoiGgoJ6F1N5TU',
  ACCE_PACE:    '1DJYnq3qXC9Td8Fip1LJvFy4qn8XBIIkcGhT8QC2m_M8',
  MSZB_PACE:    '1Trpq5YTb4e_spNIG94bE12z1u6ZWPkPEr_wDPlcz29Y',
  TAWV_PACE:    '1qGWXaxS2yLu3jX9hTIxX87jwnT1ONi59kFNSSXVIaDY',
  TECN1_PACE:   '1e2PKWFMxEHoc-9K_0_omXVLiLIT8bASte0J_s6KCiMg',
  TECN1_COM:    '1Q4cgA_XG2Q6TxPzIB5oKb0vHGUyaQa-bFsJwFc41hrQ',
  FARS_PACE:    '1y_veFQsSb7q8Nzz_Q0ZD68D29vWM20c578JZDVZLkpU',
  FARS_COM:     '1X-tCWiJeRVLuRiCXjuGlpXBVbhcQRROQJRn_QefJeIo',
  SOTB_PACE:    '1GDiRcey6G4kmx0ybSoIhVJSMRx-LGM5Ae7LepLNMoB0'
};

const GID_MAP = {
  FCE_PACE:  617612824,  ACCE_PACE: 2105044438, MSZB_PACE: 368538734,
  TAWV_PACE: 2105044438, TECN1_PACE:2105044438, FARS_PACE: 105734657,
  SOTB_PACE: 617612824,  FCE_COM:   0,          TECN1_COM: 0,
  FARS_COM:  285026569
};

function doGet(e) {
  const key = e.parameter.key;
  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);

  try {
    if (!key || !SHEET_IDS[key]) {
      output.setContent(JSON.stringify({ error: 'Chave invalida: ' + key }));
      return output;
    }

    const ss = SpreadsheetApp.openById(SHEET_IDS[key]);
    const gid = GID_MAP[key];
    let sheet;

    if (gid === 0 || gid === undefined) {
      sheet = ss.getSheets()[0];
    } else {
      const sheets = ss.getSheets();
      sheet = sheets.find(s => s.getSheetId() === gid) || sheets[0];
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(String);
    const rows = data.slice(1).map(row => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = row[i] !== undefined ? String(row[i]) : ''; });
      return obj;
    }).filter(r => Object.values(r).some(v => v.trim && v.trim()));

    output.setContent(JSON.stringify({ key, rows, headers, total: rows.length, ok: true }));
  } catch (err) {
    output.setContent(JSON.stringify({ error: err.message, key }));
  }

  return output;
}

function getAllPace() {
  const result = {};
  const paceKeys = ['FCE_PACE','ACCE_PACE','MSZB_PACE','TAWV_PACE','TECN1_PACE','FARS_PACE','SOTB_PACE'];
  paceKeys.forEach(key => {
    try {
      const ss = SpreadsheetApp.openById(SHEET_IDS[key]);
      const gid = GID_MAP[key];
      const sheets = ss.getSheets();
      const sheet = sheets.find(s => s.getSheetId() === gid) || sheets[0];
      const data = sheet.getDataRange().getValues();
      result[key] = { rows: data.map(r => r.map(String)), ok: true };
    } catch(e) {
      result[key] = { error: e.message, ok: false };
    }
  });
  return result;
}
