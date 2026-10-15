const express = require('express');
const cors = require('cors');
const { GoogleSpreadsheet } = require('google-spreadsheet');
const { JWT } = require('google-auth-library');
require('dotenv').config();

const app = express();
app.use(cors());
app.use(express.json());

const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
const LEADER_PASSWORD = process.env.LEADER_PASSWORD || '1366';

// Initialize Google Sheet Connection
async function getSheet(sheetName) {
    const serviceAccountAuth = new JWT({
        email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
        key: process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n'),
        scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });

    const doc = new GoogleSpreadsheet(SPREADSHEET_ID, serviceAccountAuth);
    await doc.loadInfo();
    return doc.sheetsByTitle[sheetName];
}

// Helper for Unique ID Generation
const generateId = (prefix) => `${prefix}_${Date.now()}_${Math.random().toString(36).substr(2, 4)}`;

// --- PUBLIC & GENERAL ENDPOINTS ---

// 1. Get Categories
app.get('/api/categories', async (req, res) => {
    try {
        const sheet = await getSheet('Categories');
        const rows = await sheet.getRows();
        const data = rows.map(r => r.toObject());
        res.json({ success: true, data });
    } catch (error) {
        res.status(500).json({ success: false, error: error.message });
    }
});

// 2. Get Processes
app.get('/api/processes', async (req, res) => {
    try {
        const sheet = await getSheet('Processes');
        const rows = await sheet.getRows();
        const data = rows.map(r => r.toObject()).filter(p => p.Status === 'Published');
        res.json({ success: true, data });
    } catch (error) {
        res.status(500).json({ success: false, error: error.message });
    }
});

// 3. Get Process Details
app.get('/api/process-detail', async (req, res) => {
    try {
        const { id } = req.query;
        const sheet = await getSheet('Processes');
        const rows = await sheet.getRows();
        const row = rows.find(r => r.get('ProcessID') === id);
        if (!row) return res.status(404).json({ success: false, message: 'Process not found' });
        res.json({ success: true, data: row.toObject() });
    } catch (error) {
        res.status(500).json({ success: false, error: error.message });
    }
});

// 4. Get Promotions (Auto calculate Status: Active, Upcoming, Expired)
app.get('/api/promotions', async (req, res) => {
    try {
        const sheet = await getSheet('Promotions');
        const rows = await sheet.getRows();
        const today = new Date().toISOString().split('T')[0];

        const data = rows.map(r => {
            const promo = r.toObject();
            if (promo.Status === 'Cancelled') return null;

            let computedStatus = promo.Status;
            if (promo.StartDate > today) {
                computedStatus = 'Upcoming';
            } else if (promo.StartDate <= today && promo.EndDate >= today) {
                computedStatus = 'Active';
            } else if (promo.EndDate < today) {
                computedStatus = 'Expired';
            }
            return { ...promo, ComputedStatus: computedStatus };
        }).filter(p => p !== null && p.Status !== 'Archived');

        res.json({ success: true, data });
    } catch (error) {
        res.status(500).json({ success: false, error: error.message });
    }
});

// --- LEADER AUTHENTICATION ---
app.post('/api/login', (req, res) => {
    const { password } = req.body;
    if (password === LEADER_PASSWORD) {
        res.json({ success: true, message: 'Login successful' });
    } else {
        res.status(401).json({ success: false, message: 'Invalid Password' });
    }
});

// --- LEADER CONTENT MANAGEMENT ENDPOINTS ---

// Middleware to check Leader Password via Header
const verifyLeaderAuth = (req, res, next) => {
    const pwd = req.headers['x-leader-password'];
    if (pwd !== LEADER_PASSWORD) {
        return res.status(401).json({ success: false, message: 'Unauthorized: Invalid Leader Password' });
    }
    next();
};

// 5. Create Process
app.post('/api/process', verifyLeaderAuth, async (req, res) => {
    try {
        const sheet = await getSheet('Processes');
        const { Title, CategoryID, Summary, ProcessContent, AttachmentLink, SearchKeywords, EffectiveDate, Status } = req.body;
        const newRow = {
            ProcessID: generateId('PROC'),
            Title, CategoryID, Summary, ProcessContent, AttachmentLink, SearchKeywords,
            Status: Status || 'Published',
            EffectiveDate: EffectiveDate || new Date().toISOString().split('T')[0],
            CreatedDate: new Date().toISOString(),
            UpdatedDate: new Date().toISOString()
        };
        await sheet.addRow(newRow);
        res.json({ success: true, message: 'Process created successfully', data: newRow });
    } catch (error) {
        res.status(500).json({ success: false, error: error.message });
    }
});

// 6. Update Process
app.put('/api/process', verifyLeaderAuth, async (req, res) => {
    try {
        const { ProcessID, ...updateData } = req.body;
        const sheet = await getSheet('Processes');
        const rows = await sheet.getRows();
        const row = rows.find(r => r.get('ProcessID') === ProcessID);
        if (!row) return res.status(404).json({ success: false, message: 'Process not found' });

        for (const key in updateData) {
            row.set(key, updateData[key]);
        }
        row.set('UpdatedDate', new Date().toISOString());
        await row.save();
        res.json({ success: true, message: 'Process updated successfully' });
    } catch (error) {
        res.status(500).json({ success: false, error: error.message });
    }
});

// 7. Archive Process
app.post('/api/process-archive', verifyLeaderAuth, async (req, res) => {
    try {
        const { ProcessID } = req.body;
        const sheet = await getSheet('Processes');
        const rows = await sheet.getRows();
        const row = rows.find(r => r.get('ProcessID') === ProcessID);
        if (!row) return res.status(404).json({ success: false, message: 'Process not found' });

        row.set('Status', 'Archived');
        row.set('UpdatedDate', new Date().toISOString());
        await row.save();
        res.json({ success: true, message: 'Process archived successfully' });
    } catch (error) {
        res.status(500).json({ success: false, error: error.message });
    }
});

// 8. Create Promotion
app.post('/api/promotion', verifyLeaderAuth, async (req, res) => {
    try {
        const sheet = await getSheet('Promotions');
        const { PromotionName, CategoryID, Description, PromotionRules, ApplicableProducts, StartDate, EndDate, AttachmentLink, Status } = req.body;
        const newRow = {
            PromotionID: generateId('PROMO'),
            PromotionName, CategoryID, Description, PromotionRules, ApplicableProducts, StartDate, EndDate, AttachmentLink,
            Status: Status || 'Published',
            CreatedDate: new Date().toISOString(),
            UpdatedDate: new Date().toISOString()
        };
        await sheet.addRow(newRow);
        res.json({ success: true, message: 'Promotion created successfully', data: newRow });
    } catch (error) {
        res.status(500).json({ success: false, error: error.message });
    }
});

// 9. Update Promotion
app.put('/api/promotion', verifyLeaderAuth, async (req, res) => {
    try {
        const { PromotionID, ...updateData } = req.body;
        const sheet = await getSheet('Promotions');
        const rows = await sheet.getRows();
        const row = rows.find(r => r.get('PromotionID') === PromotionID);
        if (!row) return res.status(404).json({ success: false, message: 'Promotion not found' });

        for (const key in updateData) {
            row.set(key, updateData[key]);
        }
        row.set('UpdatedDate', new Date().toISOString());
        await row.save();
        res.json({ success: true, message: 'Promotion updated successfully' });
    } catch (error) {
        res.status(500).json({ success: false, error: error.message });
    }
});

// 10. Cancel Promotion
app.post('/api/promotion-cancel', verifyLeaderAuth, async (req, res) => {
    try {
        const { PromotionID } = req.body;
        const sheet = await getSheet('Promotions');
        const rows = await sheet.getRows();
        const row = rows.find(r => r.get('PromotionID') === PromotionID);
        if (!row) return res.status(404).json({ success: false, message: 'Promotion not found' });

        row.set('Status', 'Cancelled');
        row.set('UpdatedDate', new Date().toISOString());
        await row.save();
        res.json({ success: true, message: 'Promotion cancelled successfully' });
    } catch (error) {
        res.status(500).json({ success: false, error: error.message });
    }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
