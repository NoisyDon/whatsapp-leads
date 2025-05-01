import express from 'express'
import bodyParser from 'body-parser'
import dotenv from 'dotenv'
import { google } from 'googleapis'
import cron from 'node-cron'

dotenv.config()

// ─── Env ───────────────────────────────────────────────────────
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

// ─── Sheets client ─────────────────────────────────────────────
const auth = new google.auth.JWT(
  GOOGLE_SERVICE_ACCOUNT_EMAIL,
  null,
  GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n'),
  ['https://www.googleapis.com/auth/spreadsheets']
)
const sheets = google.sheets({ version:'v4', auth })

// ─── Load keywords *every time* ─────────────────────────────────
async function loadKeywords() {
  const [ls, pt, sr, rp] = await Promise.all([
    sheets.spreadsheets.values.get({ spreadsheetId: KEYWORD_SHEET_ID, range: 'LeadSources!A2:B' }),
    sheets.spreadsheets.values.get({ spreadsheetId: KEYWORD_SHEET_ID, range: 'ProductTypes!A2:B' }),
    sheets.spreadsheets.values.get({ spreadsheetId: KEYWORD_SHEET_ID, range: 'StatusRules!A2:B' }),
    sheets.spreadsheets.values.get({ spreadsheetId: KEYWORD_SHEET_ID, range: 'RemarkPatterns!A2:B' })
  ])

  const leadPatterns   = (ls.data.values||[]).map(([l,p])=>({ label:l, pattern:(p||l).trim() }))
  const typePatterns   = (pt.data.values||[]).map(([l,p])=>({ label:l, pattern:(p||l).trim() }))
  const statusRules    = (sr.data.values||[]).map(([s,p])=>({ status:s, pattern:p.trim() }))
  const remarkPatterns = (rp.data.values||[]).map(([l,p])=>({ label:l, pattern:(p||'').trim() }))

  console.log('🔑 leadPatterns =', leadPatterns)
  console.log('🔑 typePatterns =', typePatterns)
  console.log('🔑 statusRules  =', statusRules)
  console.log('🔑 remarkPatterns=', remarkPatterns)

  return { leadPatterns, typePatterns, statusRules, remarkPatterns }
}

// ─── Analyze one message ───────────────────────────────────────
function analyzeText(text, { leadPatterns, typePatterns, statusRules, remarkPatterns }) {
  console.log('🔍 analyzeText on:', text)

  const findOne = (arr, def, prop='label', useStatus=false) => {
    for (const e of arr) {
      try {
        if (new RegExp(`\\b${e.pattern}\\b`, 'i').test(text))
          return useStatus ? e.status : e[prop]
      } catch{}
    }
    return def
  }

  const leads  = findOne(leadPatterns,  'Unknown Leads')
  const status = findOne(statusRules,   'in progress', 'status', true)

  const matched = typePatterns
    .filter(({ pattern }) => {
      try { return new RegExp(`\\b${pattern}\\b`, 'i').test(text) }
      catch { return false }
    })
    .map(({ label }) => label)

  console.log('✅ matched types =', matched)
  const type = matched.length ? matched.join(', ') : 'others'

  let remarks = ''
  for (const { pattern } of remarkPatterns) {
    try {
      const m = new RegExp(pattern, 'i').exec(text)
      if (m?.[1]) { remarks = m[1].trim(); break }
    } catch{}
  }

  return { leads, status, type, remarks }
}

// ─── Upsert helpers ────────────────────────────────────────────
async function findRowByContact(phone) {
  const res = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: `${SHEET_NAME}!C2:C` })
  const rows = res.data.values||[]
  return rows.findIndex(r=>r[0]===phone)
}

async function upsertLead({ date, name, phone, leads, status, remarks, type }) {
  const row = [ date, name, phone, leads, status, remarks, type, '', '' ]
  const idx = await findRowByContact(phone)
  if (idx >= 0) {
    console.log(`✏️ Updating row ${idx+2}`)
    await sheets.spreadsheets.values.update({
      spreadsheetId: SHEET_ID,
      range: `${SHEET_NAME}!A${idx+2}:I${idx+2}`,
      valueInputOption:'RAW',
      requestBody:{ values:[row] }
    })
  } else {
    console.log('➕ Appending new row')
    await sheets.spreadsheets.values.append({
      spreadsheetId: SHEET_ID,
      range: `${SHEET_NAME}!A2:I`,
      valueInputOption:'RAW',
      insertDataOption:'INSERT_ROWS',
      requestBody:{ values:[row] }
    })
  }
}

// ─── Cron (unchanged) ─────────────────────────────────────────
cron.schedule('0 0 * * *', async () => {
  try {
    const now = Date.now()
    const res = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: `${SHEET_NAME}!A2:E` })
    const rows = res.data.values||[]
    for (let i=0; i<rows.length; i++) {
      const [ d,, , , st ] = rows[i]
      if (st==='in progress' && now - new Date(d).getTime() > 7*24*3600*1000) {
        console.log(`🔄 Mark row ${i+2} no reply`)
        await sheets.spreadsheets.values.update({
          spreadsheetId: SHEET_ID,
          range: `${SHEET_NAME}!E${i+2}:E${i+2}`,
          valueInputOption:'RAW',
          requestBody:{ values:[['no reply']] }
        })
      }
    }
  } catch(e){ console.error('Cron failed',e) }
})

// ─── Express & Webhook ────────────────────────────────────────
const app = express()
app.use(bodyParser.json())

app.get('/', (_,res) => res.send('✅ up'))

app.get('/webhook', (req,res) => {
  const { 'hub.mode':m, 'hub.verify_token':t, 'hub.challenge':c } = req.query
  if (m==='subscribe' && t===WHATSAPP_VERIFY_TOKEN) return res.send(c)
  res.sendStatus(403)
})

app.post('/webhook', async (req,res) => {
  console.log('🔔 payload:', JSON.stringify(req.body))
  try {
    const { leadPatterns, typePatterns, statusRules, remarkPatterns } = await loadKeywords()
    for (const e of req.body.entry||[]) {
      for (const ch of e.changes||[]) {
        for (const msg of ch.value.messages||[]) {
          const date  = new Date(Number(msg.timestamp)*1000).toISOString().split('T')[0]
          const phone = msg.from
          const name  = (ch.value.contacts||[]).find(c=>c.wa_id===phone)?.profile.name||''
          const parsed = analyzeText(msg.text?.body||'', { leadPatterns, typePatterns, statusRules, remarkPatterns })
          await upsertLead({ date, name, phone, ...parsed })
        }
      }
    }
    res.sendStatus(200)
  } catch(err) {
    console.error('❌ webhook ERROR stack:', err.stack || err)
    res.sendStatus(500)
  }
})

if (!process.env.VERCEL) {
  app.listen(PORT, ()=> console.log(`✅ localhost:${PORT}`))
}

export default app
