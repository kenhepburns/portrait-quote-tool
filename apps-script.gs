/**
 * 形象寫真報價 → Google Drive 歸檔
 *
 * 部署設定：
 *   1. 開 https://script.google.com → 新增專案
 *   2. 把整份檔案內容貼上、儲存（檔名隨意，例：形象寫真報價歸檔）
 *   3. 修改下方 ROOT_FOLDER_ID 為你想要存放的 Drive 資料夾 ID
 *      （從資料夾分享網址末段擷取，例：https://drive.google.com/drive/folders/XXXXX）
 *   4. 點右上角「部署 → 新增部署」
 *      - 類型：網頁應用程式 (Web app)
 *      - 說明：形象寫真報價歸檔
 *      - 執行身分：我自己
 *      - 存取權：所有人 (Anyone)
 *   5. 第一次部署會跳授權 → 同意（Drive、Spreadsheet 權限）
 *   6. 複製拿到的 Web app URL，貼回前端設定的「Apps Script 網址」欄位
 *
 * 收到的 payload（form POST，欄位 name="payload"，值為 JSON 字串）：
 *   {
 *     quoteNumber, quoteDate, validUntil, status, statusLabel,
 *     clientName, clientPhone, clientEmail, shootDate,
 *     discountGroup, subtotal, discountTotal, tax, total, deposit,
 *     simplePdfBase64, detailPdfBase64,
 *     simpleHtml, detailHtml,
 *     stateJson, baseFileName
 *   }
 *
 * 行為：
 *   - 在 ROOT_FOLDER_ID 下建立/找到 客戶名 子資料夾
 *   - 存入 簡潔版/詳細版 PDF + HTML + 原始 JSON
 *   - 在根目錄維護一個「形象寫真報價總覽」Google Sheet，每次歸檔追加一列
 */

// ⚠️ Drive 根資料夾 ID（kenhepburns@gmail.com 的「形象寫真報價」資料夾）
// 資料夾網址：https://drive.google.com/drive/folders/1AiLZ-8h7eyBOOLrRGVss8VqlzrzrOSB7
const ROOT_FOLDER_ID = '1AiLZ-8h7eyBOOLrRGVss8VqlzrzrOSB7'
const SHEET_NAME = '形象寫真報價總覽'

function doPost(e) {
  try {
    const payload = JSON.parse(e.parameter.payload)
    const root = DriveApp.getFolderById(ROOT_FOLDER_ID)

    const clientName = sanitize(payload.clientName || '未命名客戶')
    const clientFolder = getOrCreateFolder(root, clientName)

    const base = sanitize(payload.baseFileName || payload.quoteNumber)
    const out = {}

    if (payload.simplePdfBase64) {
      out.simplePdfUrl = saveFile(
        clientFolder, `${base}_簡潔版.pdf`,
        Utilities.base64Decode(payload.simplePdfBase64), 'application/pdf'
      )
    }

    if (payload.detailPdfBase64) {
      out.detailPdfUrl = saveFile(
        clientFolder, `${base}_詳細版.pdf`,
        Utilities.base64Decode(payload.detailPdfBase64), 'application/pdf'
      )
    }

    if (payload.simpleHtml) {
      out.simpleHtmlUrl = saveFile(
        clientFolder, `${base}_簡潔版.html`,
        payload.simpleHtml, 'text/html'
      )
    }

    if (payload.detailHtml) {
      out.detailHtmlUrl = saveFile(
        clientFolder, `${base}_詳細版.html`,
        payload.detailHtml, 'text/html'
      )
    }

    if (payload.stateJson) {
      out.stateUrl = saveFile(
        clientFolder, `${base}_資料.json`,
        payload.stateJson, 'application/json'
      )
    }

    appendToMasterSheet(root, payload, out, clientFolder)

    return jsonOut({ ok: true, folder: clientFolder.getUrl(), files: out })
  } catch (err) {
    return jsonOut({ ok: false, error: String(err), stack: err.stack || '' })
  }
}

function doGet(e) {
  const action = e && e.parameter && e.parameter.action
  if (action === 'list') return handleList()
  if (action === 'load') return handleLoad(e.parameter.fileId)
  return ContentService.createTextOutput(
    '✅ 形象寫真報價歸檔服務運作中\n根資料夾：' + ROOT_FOLDER_ID + '\n總覽 Sheet：' + SHEET_NAME
  )
}

function handleList() {
  try {
    const root = DriveApp.getFolderById(ROOT_FOLDER_ID)
    const ss = getOrCreateMasterSheet(root)
    const data = ss.getActiveSheet().getDataRange().getValues()
    const quotes = []
    // 欄位索引：0歸檔時間 1編號 2建立者 3代碼 4最後編輯者 5最後編輯時間 6狀態 7狀態文字 8客戶 9電話 10Email 11拍攝日 12報價日 13有效期 14族群 15小計 16折扣 17稅 18總額 19訂金 20簡潔PDF 21詳細PDF 22簡潔HTML 23詳細HTML 24JSON 25stateFileId 26資料夾
    for (let i = data.length - 1; i >= 1; i--) {
      const r = data[i]
      const fileId = r[25]
      if (!fileId) continue
      quotes.push({
        archivedAt:    r[0],
        quoteNumber:   r[1],
        creator:       r[2],
        creatorCode:   r[3],
        lastEditedBy:  r[4],
        lastEditedAt:  r[5] ? (r[5] instanceof Date ? r[5].toISOString() : String(r[5])) : '',
        status:        r[6],
        statusLabel:   r[7],
        clientName:    r[8],
        shootDate:     r[11],
        quoteDate:     r[12],
        total:         r[18],
        fileId:        fileId
      })
      if (quotes.length >= 50) break
    }
    return jsonOut({ ok: true, quotes })
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) })
  }
}

function handleLoad(fileId) {
  try {
    if (!fileId) return jsonOut({ ok: false, error: '缺少 fileId' })
    const content = DriveApp.getFileById(fileId).getBlob().getDataAsString()
    return jsonOut({ ok: true, state: JSON.parse(content) })
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) })
  }
}

// ─── helpers ─────────────────────────────────────────────────────

function saveFile(folder, name, content, mime) {
  const it = folder.getFilesByName(name)
  while (it.hasNext()) it.next().setTrashed(true)
  const blob = Utilities.newBlob(content, mime, name)
  return folder.createFile(blob).getUrl()
}

function getOrCreateFolder(parent, name) {
  const it = parent.getFoldersByName(name)
  return it.hasNext() ? it.next() : parent.createFolder(name)
}

function getOrCreateMasterSheet(parent) {
  const it = parent.getFilesByName(SHEET_NAME)
  let ss
  if (it.hasNext()) {
    ss = SpreadsheetApp.open(it.next())
  } else {
    ss = SpreadsheetApp.create(SHEET_NAME)
    DriveApp.getFileById(ss.getId()).moveTo(parent)
    initMasterSheet(ss)
  }
  return ss
}

function initMasterSheet(ss) {
  const sheet = ss.getActiveSheet()
  sheet.setName('歸檔記錄')
  const headers = [
    '歸檔時間', '報價單編號', '建立者', '使用者代碼',
    '最後編輯者', '最後編輯時間',
    '狀態', '狀態文字',
    '客戶姓名', '電話', 'Email', '拍攝日期',
    '報價日期', '有效期限', '優惠族群',
    '小計(NT$)', '折扣(NT$)', '稅(NT$)', '總計(NT$)', '訂金(NT$)',
    '簡潔版 PDF', '詳細版 PDF', '簡潔版 HTML', '詳細版 HTML',
    '原始資料 JSON', 'stateFileId', '客戶資料夾'
  ]
  sheet.appendRow(headers)
  sheet.setFrozenRows(1)
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#1a3560')
    .setFontColor('white')
  sheet.setColumnWidth(1, 150)
  sheet.setColumnWidth(2, 140)
  sheet.setColumnWidth(3, 100)
  sheet.setColumnWidth(7, 160)
}

function extractFileId(url) {
  const m = String(url || '').match(/\/d\/([a-zA-Z0-9_-]+)/)
  return m ? m[1] : ''
}

function appendToMasterSheet(root, payload, out, clientFolder) {
  const ss = getOrCreateMasterSheet(root)
  const sheet = ss.getActiveSheet()
  const linkOrEmpty = (url, label) => url ? `=HYPERLINK("${url}","${label}")` : ''
  sheet.appendRow([
    new Date(),
    payload.quoteNumber || '',
    payload.creator || '',
    payload.creatorCode || '',
    payload.lastEditedBy || '',
    payload.lastEditedAt || '',
    payload.status || '',
    payload.statusLabel || '',
    payload.clientName || '',
    payload.clientPhone || '',
    payload.clientEmail || '',
    payload.shootDate || '',
    payload.quoteDate || '',
    payload.validUntil || '',
    payload.discountGroup || '',
    Number(payload.subtotal) || 0,
    Number(payload.discountTotal) || 0,
    Number(payload.tax) || 0,
    Number(payload.total) || 0,
    Number(payload.deposit) || 0,
    linkOrEmpty(out.simplePdfUrl, '開啟'),
    linkOrEmpty(out.detailPdfUrl, '開啟'),
    linkOrEmpty(out.simpleHtmlUrl, '開啟'),
    linkOrEmpty(out.detailHtmlUrl, '開啟'),
    linkOrEmpty(out.stateUrl, '下載'),
    extractFileId(out.stateUrl),
    linkOrEmpty(clientFolder.getUrl(), '進入')
  ])
}

function sanitize(name) {
  return String(name).replace(/[\\/:*?"<>|]/g, '_').trim() || '未命名'
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON)
}
