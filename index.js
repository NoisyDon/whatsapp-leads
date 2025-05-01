// index.js
import express from 'express'
import bodyParser from 'body-parser'
import dotenv from 'dotenv'
import { google } from 'googleapis'
import cron from 'node-cron'

dotenv.config()

// ─── Env vars & validation ─────────────────────────────────────
const {
  SHEET_ID,
  KEYWORD_SHEET_ID,
  SHEET_NAME = 'Leads',
  WHATSAPP_VERIFY_TOKEN,
  GOOGLE_SERVICE_ACCOUNT_EMAIL,
  GOOGLE_PRIVATE_KEY,
  PORT = 3000
} = process.env

for (const key of [
  'SHEET_ID',
  'KEYWORD_SHEET_ID',
  'WHATSAPP_VERIFY_TOKEN',
  'GOOGLE_SERVICE_ACCOUNT_EMAIL',
  'GOOGLE_PRIVATE_KEY'
]) {
  if (!process.env[key]) {
    console.error(`❌ Missing env var: ${key}`)
    process.exit(1)
  }
}

// ─── Google Sheets client ──────────────────────────────────────
const auth = new google.auth.JWT(
  GOOGLE_SERVICE_ACCOUNT_EMAIL,
  null,
  GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n'),
  ['https://www.googleapis.com/auth/spreadsheets']
)
const sheets = google.sheets({ version: 'v4', auth })

// ─── Keyword patterns (loaded once) ────────────────────────────
let loaded = false
let leadPatterns = []
let typePatterns = []
let statusRules = []
let remarkPatterns = []

async function ensureKeywords() {
  if (loaded) return
  const [ls, pt, sr, rp] = await Promise.all([
    sheets.spreadsheets.values.get({
      spreadsheetId: KEYWORD_SHEET_ID,
      range: 'LeadSources!A2:B'
    }),
    sheets.spreadsheets.values.get({
      spreadsheetId: KEYWORD_SHEET_ID,
      range: 'ProductTypes!A2:B'
    }),
    sheets.spreadsheets.values.get({
      spreadsheetId: KEYWORD_SHEET_ID,
      range: 'StatusRules!A2:B'
    }),
    sheets.spreadsheets.values.get({
      spreadsheetId: KEYWORD_SHEET_ID,
      range: 'RemarkPatterns!A2:B'
    })
  ])
  leadPatterns  = (ls.data.values  || []).map(([label, pat]) => ({ label, pattern: pat || label }))
  typePatterns  = (pt.data.values  || []).map(([label, pat]) => ({ label, pattern: pat || label }))
  statusRules   = (sr.data.values  || []).map(([status, pat]) => ({ status, pattern: pat }))
  remarkPatterns= (rp.data.values  || []).map(([label, pat]) => ({ label, pattern: pat }))
  loaded = true
}

// ─── Text analysis ────────────────────────────────────────────
function analyzeText(text) {
  const leads  = leadPatterns.find(r => new RegExp(r.pattern, 'i').test(text))?.label || 'Unknown Leads'
  const status = statusRules.find(r => new RegExp(r.pattern, 'i').test(text))?.status || 'in progress'
  const type   = typePatterns.find(r => new RegExp(r.pattern, 'i').test(text))?.label  || 'others'
  let remarks  = ''
  for (const { pattern } of remarkPatterns) {
    const m = new RegExp(pattern, 'i').exec(text)
    if (m?.[1]) { remarks = m[1].trim(); break }
  }
  return { leads, status, type, remarks }
}

// ─── Sheet upsert helpers ──────────────────────────────────────
async function findRowByContact(phone) {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: `${SHEET_NAME}!C2:C`
  })
  const rows = res.data.values || []
  return rows.findIndex(r => r[0] === phone)
}

async function upsertLead({ date, name, phone, leads, status, remarks, type }) {
  const row = [ date, name, phone, leads, status, remarks, type, '', '' ]
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

// ─── Cron job: mark stale leads “no reply” ──────────────────────
cron.schedule('0 0 * * *', async () => {
  try {
    const now = Date.now()
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: `${SHEET_NAME}!A2:E`
    })
    const rows = res.data.values || []
    for (let i = 0; i < rows.length; i++) {
      const [ date, , , , stat ] = rows[i]
      if (stat === 'in progress') {
        const ts = new Date(date).getTime()
        if (now - ts > 7 * 24 * 3600 * 1000) {
          const rowNum = i + 2
          await sheets.spreadsheets.values.update({
            spreadsheetId: SHEET_ID,
            range: `${SHEET_NAME}!E${rowNum}:E${rowNum}`,
            valueInputOption: 'RAW',
            requestBody: { values: [['no reply']] }
          })
        }
      }
    }
  } catch (err) {
    console.error('Cron job error:', err)
  }
})

// ─── Express setup ────────────────────────────────────────────
const app = express()
app.use(bodyParser.json())

// Health-check
app.get('/', (_req, res) => {
  res.status(200).send('✅ WhatsApp-Leads webhook is up')
})

// Webhook verification
app.get('/webhook', (req, res) => {
  const mode      = req.query['hub.mode']
  const token     = req.query['hub.verify_token']
  const challenge = req.query['hub.challenge']
  if (mode === 'subscribe' && token === WHATSAPP_VERIFY_TOKEN) {
    return res.status(200).send(challenge)
  }
  res.status(403).send('Forbidden')
})

// Incoming messages
app.post('/webhook', async (req, res) => {
  try {
    await ensureKeywords()
    for (const entry of req.body.entry || []) {
      for (const change of entry.changes || []) {
        const contacts = change.value.contacts || []
        for (const msg of change.value.messages || []) {
          const date  = new Date(Number(msg.timestamp) * 1000).toISOString().split('T')[0]
          const phone = msg.from
          const name  = contacts.find(c => c.wa_id === phone)?.profile.name || ''
          const text  = msg.text?.body || ''
          const { leads, status, type, remarks } = analyzeText(text)
          await upsertLead({ date, name, phone, leads, status, remarks, type })
        }
      }
    }
    res.status(200).send('OK')
  } catch (err) {
    console.error('POST /webhook error:', err)
    res.status(500).send('Error')
  }
})

export default app
