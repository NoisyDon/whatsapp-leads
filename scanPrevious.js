// scanPrevious.js
import fs from 'fs'
import dotenv from 'dotenv'
import { google } from 'googleapis'

dotenv.config()

// Reuse the same auth & parsing code from index.js:
const {
  SHEET_ID,
  KEYWORD_SHEET_ID,
  SHEET_NAME = 'Leads',
  GOOGLE_SERVICE_ACCOUNT_EMAIL,
  GOOGLE_PRIVATE_KEY
} = process.env

const auth = new google.auth.JWT(
  GOOGLE_SERVICE_ACCOUNT_EMAIL,
  null,
  GOOGLE_PRIVATE_KEY.replace(/\\n/g,'\n'),
  ['https://www.googleapis.com/auth/spreadsheets']
)
const sheets = google.sheets({version:'v4', auth})

// Paste loadKeywords, analyzeText, findRowByContact, upsertLead, processMessage from index.js here...
// For brevity, assume you copy those functions verbatim.

async function main() {
  const data = JSON.parse(fs.readFileSync(process.argv[2],'utf8'))
  await loadKeywords()
  for (let { msg, contacts } of data) {
    try {
      await processMessage(msg, contacts)
    } catch(e) {
      console.error('Backfill error:', e)
    }
  }
}

main()
