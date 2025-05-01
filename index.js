// index.js
import express from 'express'
import bodyParser from 'body-parser'
import dotenv from 'dotenv'
import { google } from 'googleapis'
import cron from 'node-cron'

dotenv.config()

// ─── Env Vars & Validation ─────────────────────────────────────
const {
  SHEET_ID,
  KEYWORD_SHEET_ID,
  SHEET_NAME = 'Leads',
  WHATSAPP_VERIFY_TOKEN,
  GOOGLE_SERVICE_ACCOUNT_EMAIL,
  GOOGLE_PRIVATE_KEY,
  PORT = '3000'
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

// ─── Google Sheets Client ──────────────────────────────────────
const auth = new google.auth.JWT(
  GOOGLE_SERVICE_ACCOUNT_EMAIL,
  null,
  GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n'),
  ['https://www.googleapis.com/auth/spreadsheets']
)
const sheets = google.sheets({ version: 'v4', auth })

// ─── Keyword Patterns (load once) ──────────────────────────────
let loaded = false
let leadPatterns = [], typePatterns = [], statusRules = [], remarkPatterns = []

async function ensureKeywords() {
  if (loaded) return
  const [ ls, pt, sr, rp ] = await Promise.all([
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

  leadPatterns   = (ls.data.values   || []).map(([label,pat]) => ({ label, pattern: (pat||label).trim() }))
  typePatterns   = (pt.data.values   || []).map(([label,pat]) => ({ label, pattern: (pat||label).trim() }))
  statusRules    = (sr.data.values   || []).map(([status,pat])=> ({ status, pattern: pat.trim() }))
  remarkPatterns = (rp.data.values   || []).map(([label,pat])=> ({ label, pattern: (pat||'').trim().replace(/`/g,'') }))

  loaded = true
}

// ─── Text Analysis ─────────────────────────────────────────────
function analyzeText(text) {
  function find(arr, defaultVal, prop='label') {
    for (const entry of arr) {
      try {
        if (new RegExp(entry.pattern, 'i').test(text)) {
          return entry[prop]
        }
      } catch (_) {
        continue
      }
    }
    return defaultVal
  }

  const leads  = find(leadPatterns,  'Unknown Leads', 'label')
  const status = find(statusRules,   'in progress', 'status')
  const type   = find(typePatterns,   'others',      'label')

  let remarks = ''
  for (const { pattern } of remarkPatterns) {
    try {
      const m = new RegExp(pattern, 'i').exec(text)
      if (m?.[1]) {
        remarks = m[1].trim()
        break
      }
    } catch (_) { continue }
  }

  return { leads, status, type, remarks }
}

// ─── Sheet Upsert Helpers ──────────────────────────────────────
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
  console.log(`   ↪ [Upsert] phone=${phone}, idx=${idx}, row=${JSON.stringify(row)}`)
  if (idx >= 0) {
    console.log(`     ✏️  Updating row ${idx+2}`)
    await sheets.spreadsheets.values.update({
      spreadsheetId: SHEET_ID,
      range: `${SHEET_NAME}!A${idx+2}:I${idx+2}`,
      valueInputOption: 'RAW',
      requestBody: { values: [row] }
    })
  } else {
    console.log('     ➕ Appending new row')
    await sheets.spreadsheets.values.append({
      spreadsheetId: SHEET_ID,
      range: `${SHEET_NAME}!A2:I`,
      valueInputOption: 'RAW',
      insertDataOption: 'INSERT_ROWS',
      requestBody: { values: [row] }
    })
  }
}

// ─── Cron Job: Mark Stale Leads “no reply” ─────────────────────
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
          console.log(`🔄 Marking row ${rowNum} as no reply`)
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

// ─── Express App & Routes ─────────────────────────────────────
const app = express()
app.use(bodyParser.json())

// Health check
app.get('/', (_req, res) => {
  res.status(200).send('✅ WhatsApp‐Leads webhook is up')
})

// Webhook handshake
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
  console.log('🔔 [Webhook] payload:', JSON.stringify(req.body))
  try {
    await ensureKeywords()
    for (const entry of req.body.entry || []) {
      for (const change of entry.changes || []) {
        const contacts = change.value.contacts || []
        for (const msg of change.value.messages || []) {
          console.log(`   • [Msg] from=${msg.from}, body="${msg.text?.body}"`)
          const date  = new Date(Number(msg.timestamp) * 1000).toISOString().split('T')[0]
          const phone = msg.from
          const name  = contacts.find(c => c.wa_id === phone)?.profile.name || ''
          const text  = msg.text?.body || ''
          const parsed = analyzeText(text)
          await upsertLead({ date, name, phone, ...parsed })
        }
      }
    }
    res.status(200).send('OK')
  } catch (err) {
    console.error('[Webhook] Error:', err)
    res.status(500).send(err.stack || err.toString())
  }
})

// ─── Local‐only HTTP Listener ──────────────────────────────────
if (!process.env.VERCEL) {
  const p = Number(PORT) || 3000
  app.listen(p, () => {
    console.log(`✅ Local server listening on http://localhost:${p}`)
  })
}

export default app
