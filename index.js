// index.js
import express from 'express'
import bodyParser from 'body-parser'
import dotenv from 'dotenv'
import { google } from 'googleapis'
import cron from 'node-cron'

dotenv.config()

// ─── Env Vars ─────────────────────────────────────────────────
const {
  SHEET_ID,
  KEYWORD_SHEET_ID,
  SHEET_NAME = 'Leads',
  WHATSAPP_VERIFY_TOKEN,
  GOOGLE_SERVICE_ACCOUNT_EMAIL,
  GOOGLE_PRIVATE_KEY,
  PORT = 3000
} = process.env

for (let v of [
  'SHEET_ID',
  'KEYWORD_SHEET_ID',
  'WHATSAPP_VERIFY_TOKEN',
  'GOOGLE_SERVICE_ACCOUNT_EMAIL',
  'GOOGLE_PRIVATE_KEY'
]) {
  if (!process.env[v]) {
    console.error(`Missing env var: ${v}`)
    process.exit(1)
  }
}

// ─── Google Sheets setup ──────────────────────────────────────
const auth = new google.auth.JWT(
  GOOGLE_SERVICE_ACCOUNT_EMAIL,
  null,
  GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n'),
  ['https://www.googleapis.com/auth/spreadsheets']
)
const sheets = google.sheets({ version: 'v4', auth })

// ─── Keyword patterns ─────────────────────────────────────────
let leadPatterns = []    // [{ label, pattern }]
let typePatterns = []    // [{ label, pattern }]
let statusRules  = []    // [{ status, pattern }]
let remarkPatterns = []  // [{ label, pattern }]

async function loadKeywords() {
  const ls = await sheets.spreadsheets.values.get({
    spreadsheetId: KEYWORD_SHEET_ID,
    range: 'LeadSources!A2:B'
  })
  leadPatterns = (ls.data.values||[])
    .map(([label, pat]) => ({ label, pattern: pat||label }))

  const pt = await sheets.spreadsheets.values.get({
    spreadsheetId: KEYWORD_SHEET_ID,
    range: 'ProductTypes!A2:B'
  })
  typePatterns = (pt.data.values||[])
    .map(([label, pat]) => ({ label, pattern: pat||label }))

  const sr = await sheets.spreadsheets.values.get({
    spreadsheetId: KEYWORD_SHEET_ID,
    range: 'StatusRules!A2:B'
  })
  statusRules = (sr.data.values||[])
    .map(([status, pat]) => ({ status, pattern: pat }))

  const rp = await sheets.spreadsheets.values.get({
    spreadsheetId: KEYWORD_SHEET_ID,
    range: 'RemarkPatterns!A2:B'
  })
  remarkPatterns = (rp.data.values||[])
    .map(([label, pat]) => ({ label, pattern: pat }))
}

// ─── Text analysis ────────────────────────────────────────────
function analyzeText(text) {
  const leads = leadPatterns.find(r => new RegExp(r.pattern, 'i').test(text))
                ?.label || 'Unknown Leads'
  const status = statusRules.find(r => new RegExp(r.pattern, 'i').test(text))
                 ?.status || 'in progress'
  const type = typePatterns.find(r => new RegExp(r.pattern, 'i').test(text))
               ?.label || 'others'
  let remarks = ''
  for (let { pattern } of remarkPatterns) {
    const m = new RegExp(pattern, 'i').exec(text)
    if (m && m[1]) { remarks = m[1].trim(); break }
  }
  return { leads, status, type, remarks }
}

// ─── Sheet upsert helpers ─────────────────────────────────────
async function findRowByContact(phone) {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: `${SHEET_NAME}!C2:C`
  })
  const rows = res.data.values || []
  return rows.findIndex(r => r[0] === phone)
}

async function upsertLead({ date, name, phone, leads, status, remarks, type }) {
  const row = [date, name, phone, leads, status, remarks, type, '', '']
  const idx = await findRowByContact(phone)
  if (idx >= 0) {
    const rowNum = idx + 2
    await sheets.spreadsheets.values.update({
      spreadsheetId: SHEET_ID,
      range: `${SHEET_NAME}!A${rowNum}:I${rowNum}`,
      valueInputOption: 'RAW',
      requestBody: { values: [row] }
    })
  } else {
    await sheets.spreadsheets.values.append({
      spreadsheetId: SHEET_ID,
      range: `${SHEET_NAME}!A2:I`,
      valueInputOption: 'RAW',
      insertDataOption: 'INSERT_ROWS',
      requestBody: { values: [row] }
    })
  }
}

// ─── Process one WhatsApp message ─────────────────────────────
async function processMessage(msg, contacts) {
  const phone = msg.from
  const date  = new Date(Number(msg.timestamp)*1000).toISOString().split('T')[0]
  const name  = contacts.find(c=>c.wa_id===phone)?.profile.name || ''
  const text  = msg.text?.body || ''
  const { leads, status, type, remarks } = analyzeText(text)
  await upsertLead({ date, name, phone, leads, status, remarks, type })
}

// ─── Cron: mark “in progress” >7 days as “no reply” ───────────
cron.schedule('0 0 * * *', async ()=>{
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: `${SHEET_NAME}!A2:E`
  })
  const rows = res.data.values||[]
  const now = Date.now()
  for (let i=0; i<rows.length; i++){
    const [date,, , , status] = rows[i]
    if (status==='in progress') {
      const ts = new Date(date).getTime()
      if (now - ts > 7*24*3600*1000) {
        const rowNum = i+2
        await sheets.spreadsheets.values.update({
          spreadsheetId: SHEET_ID,
          range: `${SHEET_NAME}!E${rowNum}:E${rowNum}`,
          valueInputOption: 'RAW',
          requestBody:{ values:[['no reply']] }
        })
      }
    }
  }
})

// ─── Express server ───────────────────────────────────────────
const app = express()
app.use(bodyParser.json())

// Verify token handshake
app.get('/webhook',(req,res)=>{
  if (req.query['hub.verify_token']===WHATSAPP_VERIFY_TOKEN)
    return res.status(200).send(req.query['hub.challenge'])
  res.status(403).send('Forbidden')
})

// Ingest messages
app.post('/webhook',async(req,res)=>{
  try {
    for (let entry of req.body.entry||[]) {
      for (let change of entry.changes||[]) {
        for (let msg of change.value.messages||[]) {
          await processMessage(msg, change.value.contacts||[])
        }
      }
    }
    res.status(200).send('OK')
  } catch(e) {
    console.error(e)
    res.status(500).send('Error')
  }
})

// ─── Start ────────────────────────────────────────────────────
;(async ()=>{
  await loadKeywords()
  app.listen(PORT,()=>console.log(`Listening on port ${PORT}`))
})()
