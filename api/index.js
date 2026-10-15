const express = require('express');
const cors = require('cors');
const { google } = require('googleapis');
require('dotenv').config();

const app = express();
app.use(cors());
app.use(express.json());

const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
const LEADER_PASSWORD = process.env.LEADER_PASSWORD || '1366';

async function getSheetsClient() {
    if (!process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL) throw new Error("Environment Variable တွင် GOOGLE_SERVICE_ACCOUNT_EMAIL မရှိပါ။");
    if (!process.env.GOOGLE_PRIVATE_KEY) throw new Error("Environment Variable တွင် GOOGLE_PRIVATE_KEY မရှိပါ။");
    if (!SPREADSHEET_ID) throw new Error("Environment Variable တွင် SPREADSHEET_ID မရှိပါ။");
    
    let privateKey = process.env.GOOGLE_PRIVATE_KEY;
    // Fix private key formatting issues
    privateKey = privateKey.replace(/^["'](.*)["']$/, '$1').replace(/\\n/g, '\n');

    const auth = new google.auth.JWT(
        process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
        null,
        privateKey,
        ['https://www.googleapis.com/auth/spreadsheets']
    );
    return google.sheets({ version: 'v4', auth });
}

async function getSheetData(sheetName) {
    const sheets = await getSheetsClient();
    const res = await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: `${sheetName}!A1:Z1000`,
    });
    const rows = res.data.values || [];
    if (rows.length === 0) return [];
    
    const headers = rows[0];
    return rows.slice(1).map(row => {
        let obj = {};
        headers.forEach((h, index) => {
            obj[h] = row[index] !== undefined ? row[index] : '';
        });
        return obj;
    });
}

const router = express.Router();

router.get('/categories', async (req, res) => {
    try {
        const data = await getSheetData('Categories');
        res.json({ success: true, data });
    } catch (error) {
        res.status(500).json({ success: false, error: error.message, details: "Categories sheet ကို ဖတ်၍မရပါ။ Sheet နာမည် (သို့) Permission ကို စစ်ဆေးပါ။" });
    }
});

router.get('/processes', async (req, res) => {
    try {
        const data = await getSheetData('Processes');
        const published = data.filter(p => p.Status === 'Published' && p.Status !== 'Archived');
        res.json({ success: true, data: published });
    } catch (error) {
        res.status(500).json({ success: false, error: error.message, details: "Processes sheet ကို ဖတ်၍မရပါ။" });
    }
});

router.get('/promotions', async (req, res) => {
    try {
        const data = await getSheetData('Promotions');
        const today = new Date().toISOString().split('T')[0];

        const processed = data.map(promo => {
            if (promo.Status === 'Cancelled' || promo.Status === 'Archived') return null;

            let computedStatus = promo.Status;
            if (promo.StartDate > today) {
                computedStatus = 'Upcoming';
            } else if (promo.StartDate <= today && promo.EndDate >= today) {
                computedStatus = 'Active';
            } else if (promo.EndDate < today) {
                computedStatus = 'Expired';
            }
            return { ...promo, ComputedStatus: computedStatus };
        }).filter(p => p !== null);

        res.json({ success: true, data: processed });
    } catch (error) {
        res.status(500).json({ success: false, error: error.message, details: "Promotions sheet ကို ဖတ်၍မရပါ။" });
    }
});

router.post('/login', (req, res) => {
    const { password } = req.body;
    if (password === LEADER_PASSWORD) {
        res.json({ success: true, message: 'Login successful' });
    } else {
        res.status(401).json({ success: false, message: 'Invalid Password' });
    }
});

app.use('/api', router);
app.use('/', router);

module.exports = app;
