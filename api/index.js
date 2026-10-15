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
    if (!process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL || !process.env.GOOGLE_PRIVATE_KEY || !SPREADSHEET_ID) {
        throw new Error("Missing required Environment Variables in Vercel!");
    }
    
    const privateKey = process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n');
    
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

const generateId = (prefix) => `${prefix}_${Date.now()}_${Math.random().toString(36).substr(2, 4)}`;

const router = express.Router();

router.get('/categories', async (req, res) => {
    try {
        const data = await getSheetData('Categories');
        res.json({ success: true, data });
    } catch (error) {
        console.error("Error in /categories:", error);
        res.status(500).json({ success: false, error: error.message });
    }
});

router.get('/processes', async (req, res) => {
    try {
        const data = await getSheetData('Processes');
        const published = data.filter(p => p.Status === 'Published' && p.Status !== 'Archived');
        res.json({ success: true, data: published });
    } catch (error) {
        console.error("Error in /processes:", error);
        res.status(500).json({ success: false, error: error.message });
    }
});

router.get('/process-detail', async (req, res) => {
    try {
        const { id } = req.query;
        const data = await getSheetData('Processes');
        const row = data.find(p => p.ProcessID === id);
        if (!row) return res.status(404).json({ success: false, message: 'Process not found' });
        res.json({ success: true, data: row });
    } catch (error) {
        console.error("Error in /process-detail:", error);
        res.status(500).json({ success: false, error: error.message });
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
        console.error("Error in /promotions:", error);
        res.status(500).json({ success: false, error: error.message });
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

const verifyLeaderAuth = (req, res, next) => {
    const pwd = req.headers['x-leader-password'];
    if (pwd !== LEADER_PASSWORD) {
        return res.status(401).json({ success: false, message: 'Unauthorized: Invalid Leader Password' });
    }
    next();
};

router.post('/process', verifyLeaderAuth, async (req, res) => {
    try {
        const sheets = await getSheetsClient();
        const resValues = await sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: 'Processes!A1:Z1' });
        const headers = resValues.data.values[0];

        const newObj = {
            ProcessID: generateId('PROC'),
            Title: req.body.Title || '',
            CategoryID: req.body.CategoryID || '',
            Summary: req.body.Summary || '',
            ProcessContent: req.body.ProcessContent || '',
            AttachmentLink: req.body.AttachmentLink || '',
            SearchKeywords: req.body.SearchKeywords || '',
            Status: req.body.Status || 'Published',
            EffectiveDate: req.body.EffectiveDate || new Date().toISOString().split('T')[0],
            CreatedDate: new Date().toISOString(),
            UpdatedDate: new Date().toISOString()
        };

        const rowArray = headers.map(h => newObj[h] !== undefined ? newObj[h] : '');
        await sheets.spreadsheets.values.append({
            spreadsheetId: SPREADSHEET_ID,
            range: 'Processes!A1',
            valueInputOption: 'USER_ENTERED',
            requestBody: { values: [rowArray] }
        });

        res.json({ success: true, message: 'Process created successfully', data: newObj });
    } catch (error) {
        console.error("Error in POST /process:", error);
        res.status(500).json({ success: false, error: error.message });
    }
});

router.post('/promotion', verifyLeaderAuth, async (req, res) => {
    try {
        const sheets = await getSheetsClient();
        const resValues = await sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: 'Promotions!A1:Z1' });
        const headers = resValues.data.values[0];

        const newObj = {
            PromotionID: generateId('PROMO'),
            PromotionName: req.body.PromotionName || '',
            CategoryID: req.body.CategoryID || '',
            Description: req.body.Description || '',
            PromotionRules: req.body.PromotionRules || '',
            ApplicableProducts: req.body.ApplicableProducts || '',
            StartDate: req.body.StartDate || '',
            EndDate: req.body.EndDate || '',
            AttachmentLink: req.body.AttachmentLink || '',
            Status: req.body.Status || 'Published',
            CreatedDate: new Date().toISOString(),
            UpdatedDate: new Date().toISOString()
        };

        const rowArray = headers.map(h => newObj[h] !== undefined ? newObj[h] : '');
        await sheets.spreadsheets.values.append({
            spreadsheetId: SPREADSHEET_ID,
            range: 'Promotions!A1',
            valueInputOption: 'USER_ENTERED',
            requestBody: { values: [rowArray] }
        });

        res.json({ success: true, message: 'Promotion created successfully', data: newObj });
    } catch (error) {
        console.error("Error in POST /promotion:", error);
        res.status(500).json({ success: false, error: error.message });
    }
});

app.use('/api', router);
app.use('/', router);

module.exports = app;
