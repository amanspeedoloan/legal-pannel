// server.js
const express = require('express');
const multer = require('multer');
const exceljs = require('exceljs');
const PDFDocument = require('pdfkit');
const fs = require('fs');
const path = require('path');
const moment = require('moment');

const app = express();
const upload = multer({ dest: 'uploads/' });

// Ensure required folders exist
['uploads', 'generated', 'public/images', 'public/css'].forEach(dir => {
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
});

app.use(express.static('public'));
// CRITICAL: Ensure express can parse form data sent by the browser
app.use(express.urlencoded({ extended: true })); 

// -------------------------------
// ROUTES
// -------------------------------

// Dashboard (loads index.html from public)
app.get('/', (req, res) => {
    // You must have a public/index.html file for this to work
    res.sendFile(path.resolve(__dirname, 'public/index.html')); 
});

// Download Excel template
app.get('/download-template', async (req, res) => {
    const wb = new exceljs.Workbook();
    const ws = wb.addWorksheet('Template');

    const headers = [
        'Organization Name','RBI Reg No','Address','Email','Contact','Signatory',
        'Borrower Name','Mobile','Borrower Email','Borrower Address','Loan ID','Loan Amount',
        'Disbursal Date','Bank/UPI','Aadhaar Last 4','PAN Last 4','Facts'
    ];
    // NOTE: Add a leading empty string to align Excel columns with data indices (1-based)
    ws.addRow(['', ...headers]); 
    ws.addRow([
        '','Agrim Fincap Pvt Ltd','07AAACV...','Address line','info@agrim.com','+91-XXXXXX','Authorized Person',
        'Rahul Kumar','9876543210','rahul@mail.com','Delhi','LN10203','15000',
        moment().format('YYYY-MM-DD'),'HDFC/UPI','1234','ABCD','Suspected fake KYC'
    ]);

    res.setHeader('Content-Disposition', 'attachment; filename=agrim_template.xlsx');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    await wb.xlsx.write(res);
    res.end();
});

// Upload Excel -> Preview
app.post('/upload-excel', upload.single('excel'), async (req, res) => {
    try {
        if (!req.file) return res.status(400).send('No file uploaded');
        const workbook = new exceljs.Workbook();
        await workbook.xlsx.readFile(req.file.path);
        const sheet = workbook.worksheets[0];
        const rows = [];
        sheet.eachRow(r => rows.push(r.values)); 
        const headers = rows[0] || [];
        const dataRows = rows.slice(1).filter(r => r.length > 1);

        // Build preview HTML
        let html = `<html><head><meta charset="utf-8"><title>Preview</title><link rel="stylesheet" href="/css/style.css"></head><body>`;
        html += `<div class="card"><h2>Excel Preview (${dataRows.length} rows)</h2>`;
        html += `<table class="preview-table"><thead><tr>`;
        
        for (let i = 1; i < headers.length; i++) html += `<th>${headers[i] || `Col ${i}`}</th>`;
        html += `</tr></thead><tbody>`;

        dataRows.forEach(r => {
            html += `<tr>`;
            for (let i = 1; i < headers.length; i++) html += `<td>${r[i] || '-'}</td>`;
            html += `</tr>`;
        });
        
        html += `</tbody></table>`;
        // CRITICAL: Ensure the hidden field is correctly populated with the file path
        html += `<form method="POST" action="/generate-all" style="margin-top:20px;">
                    <input type="hidden" name="filePath" value="${req.file.path}">
                    <button class="btn primary">Generate All PDFs (${dataRows.length} letters)</button>
                 </form>`;
        html += `<div style="margin-top:12px"><a href="/">← Back to Dashboard</a></div></div></body></html>`;
        res.send(html);
    } catch (err) {
        console.error(err);
        if (req.file && fs.existsSync(req.file.path)) {
            try { fs.unlinkSync(req.file.path); } catch (e) {}
        }
        res.status(500).send(`Excel read error. Check file format/path. Details: ${err.message}`);
    }
});

// Generate all PDFs from uploaded Excel
app.post('/generate-all', async (req, res) => {
    let successCount = 0;
    const filePath = req.body.filePath;
    let dataRows = []; // Wider scope for logging

    try {
        if (!filePath || !fs.existsSync(filePath)) {
            console.error('ERROR: Missing or invalid filePath received for generation:', filePath);
            return res.status(400).send('Uploaded file not found. Please go back and re-upload the file.');
        }
        
        const workbook = new exceljs.Workbook();
        await workbook.xlsx.readFile(filePath);
        const sheet = workbook.worksheets[0];
        const rows = [];
        sheet.eachRow(r => rows.push(r.values));
        dataRows = rows.slice(1).filter(r => r.length > 1);

        console.log(`Starting PDF generation for ${dataRows.length} rows...`); 

        // CRITICAL FIX: Internal try-catch to ensure one failure doesn't stop the whole process
        for (const r of dataRows) {
            try { 
                const data = extractExcelRow(r);
                // Use Loan ID and Date for filename
                const name = `Complaint_${sanitize(data.loanId || moment().format('YYYYMMDDHHmmss'))}.pdf`; 
                await createPDF(data, path.join('generated', name));
                successCount++;
            } catch (innerErr) {
                const loanId = r[11] || 'N/A';
                console.error(`Failed to generate PDF for Loan ID: ${loanId}. Error: ${innerErr.message}`);
                // Continue to the next row
            }
        }

        // remove uploaded file
        try { fs.unlinkSync(filePath); } catch (e) { console.warn('Failed to delete uploaded file:', e.message); }
        
        console.log(`Finished PDF generation. Success count: ${successCount}`); 

        // Provide generation feedback and redirect
        return res.redirect(`/view-letters?generated=${successCount}`);

    } catch (err) {
        console.error('Bulk generation setup/file error:', err);
        if (filePath && fs.existsSync(filePath)) {
            try { fs.unlinkSync(filePath); } catch (e) {}
        }
        return res.status(500).send(`Generation error (processed ${successCount} files). Details: ${err.message}. Check server console for log.`);
    }
});

// View generated letters (list)
app.get('/view-letters', (req, res) => {
    const files = fs.existsSync('generated') ? fs.readdirSync('generated').filter(f => f.endsWith('.pdf')).sort().reverse() : [];
    const generatedCount = req.query.generated;
    
    let html = `<html><head><meta charset="utf-8"><title>Generated</title><link rel="stylesheet" href="/css/style.css"></head><body>`;
    html += `<div class="card"><h2>Generated Letters (${files.length})</h2>`;
    
    if (generatedCount > 0) {
        html += `<div class="alert success">${generatedCount} Complaint letters successfully generated!</div>`;
    } else if (generatedCount === '0') {
         html += `<div class="alert error">0 letters generated. Check server logs or source Excel data.</div>`;
    }

    if (!files.length) html += `<p class="muted">No letters generated yet.</p>`;
    else {
        html += `<table class="preview-table"><thead><tr><th>PDF Name</th><th>View</th><th>Download</th></tr></thead><tbody>`;
        files.forEach(f => {
            html += `<tr>
                <td>${f}</td>
                <td><a href="/generated/${encodeURIComponent(f)}" target="_blank" class="btn small view">View</a></td>
                <td><a href="/download/${encodeURIComponent(f)}" class="btn small download">Download</a></td>
            </tr>`;
        });
        html += `</tbody></table>`;
    }
    html += `<div style="margin-top:12px"><a href="/">← Back to Dashboard</a> | <a href="/download-logs">View Download Logs</a> | <a href="/generation-logs">View Generation Logs</a></div></div></body></html>`;
    res.send(html);
});

// Serve generated PDFs for viewing
app.use('/generated', express.static(path.join(__dirname, 'generated')));

// Download (tracks downloads)
app.get('/download/:filename', (req, res) => {
    try {
        const filename = req.params.filename;
        const decoded = decodeURIComponent(filename);
        const filePath = path.join(__dirname, 'generated', decoded);
        if (!fs.existsSync(filePath)) return res.status(404).send('File not found');

        // Track download log
        const dlog = path.join(__dirname, 'download_logs.json');
        let logs = [];
        if (fs.existsSync(dlog)) {
            try { logs = JSON.parse(fs.readFileSync(dlog, 'utf8')); } catch (e) { logs = []; }
        }
        logs.push({ 
            filename: decoded, 
            downloadedAt: moment().format('YYYY-MM-DD HH:mm:ss'),
        });
        const safeLogs = logs.slice(-500); 
        fs.writeFileSync(dlog, JSON.stringify(safeLogs, null, 2));

        // Send file
        res.download(filePath, decoded); 
    } catch (e) {
        console.error(e);
        res.status(500).send('Download error');
    }
});

// Download logs viewer (omitted for brevity)
app.get('/download-logs', (req, res) => {
    // ... (Your existing log viewing code) ...
    const logFile = path.join(__dirname, "download_logs.json");
    let logs = [];
    if (fs.existsSync(logFile)) {
        try { logs = JSON.parse(fs.readFileSync(logFile, "utf8")).reverse(); } catch (e) { logs = []; } 
    }

    let html = `<html><head><meta charset="utf-8"><title>Download Logs</title><link rel="stylesheet" href="/css/style.css"></head><body>`;
    html += `<div class="card"><h2>Download Logs (${logs.length} entries)</h2>`;
    if (!logs.length) html += `<p>No downloads yet.</p>`;
    else {
        html += `<table class="preview-table"><thead><tr><th>PDF Name</th><th>Downloaded At</th><th>View</th></tr></thead><tbody>`;
        logs.forEach(l => {
            html += `<tr><td>${l.filename}</td><td>${l.downloadedAt}</td><td><a href="/generated/${encodeURIComponent(l.filename)}" target="_blank">View</a></td></tr>`;
        });
        html += `</tbody></table>`;
    }
    html += `<br><a href="/">← Back to Dashboard</a></div></body></html>`;
    res.send(html);
});

// Generation logs viewer (omitted for brevity)
app.get('/generation-logs', (req, res) => {
    // ... (Your existing log viewing code) ...
    const logFile = path.join(__dirname, "logs.json");
    let logs = [];
    if (fs.existsSync(logFile)) {
        try { logs = JSON.parse(fs.readFileSync(logFile, "utf8")).reverse(); } catch (e) { logs = []; }
    }

    let html = `<html><head><meta charset="utf-8"><title>Generation Logs</title><link rel="stylesheet" href="/css/style.css"></head><body>`;
    html += `<div class="card"><h2>Generation Logs (${logs.length} entries)</h2>`;
    if (!logs.length) html += `<p>No generations yet.</p>`;
    else {
        html += `<table class="preview-table"><thead><tr><th>Loan ID</th><th>Borrower</th><th>Generated At</th><th>Status</th></tr></thead><tbody>`;
        logs.forEach(l => {
            html += `<tr><td>${l.loanId}</td><td>${l.borrower}</td><td>${l.generatedAt}</td><td>${l.status}</td></tr>`;
        });
        html += `</tbody></table>`;
    }
    html += `<br><a href="/">← Back to Dashboard</a></div></body></html>`;
    res.send(html);
});


// -------------------------------
// HELPERS
// -------------------------------
function sanitize(s) {
    return String(s || '').replace(/[\\/:*?"<>|]/g, '_');
}

function extractExcelRow(r) {
    const formatDisbursalDate = (val) => {
        if (!val) return '';
        if (moment.isDate(val)) {
            return moment(val).format('DD-MM-YYYY');
        }
        let dateVal = String(val);
        // Handle Excel date serial numbers (like 45000)
        if (!isNaN(dateVal) && Number(dateVal) > 40000 && Number(dateVal) < 70000) {
            return moment(new Date(1899, 11, 30)).add(Number(dateVal), 'days').format('DD-MM-YYYY');
        }
        return dateVal;
    };

    return {
        // Excel Index 1-based (r[index])
        email: r[4] || 'support@agrim.com',     // 4
        contact: r[5] || '9876543210',          // 5
        authSignatory: r[6] || 'Aman Mishra',   // 6
        borrowerName: r[7] || '',           // 7
        mobile: r[8] || '',                 // 8
        borrowerEmail: r[9] || '',          // 9
        address: r[10] || '',               // 10
        loanId: r[11] || '',                // 11
        loanAmount: r[12] || '',            // 12
        disbursalDate: formatDisbursalDate(r[13]), // 13 (FIXED formatting)
        bankDetails: r[14] || '',           // 14
        aadhaarLast4: r[15] || '',          // 15
        panLast4: r[16] || '',              // 16
        facts: r[17] || 'Suspected cyber fraud with details below.', // 17
        
        // Key Indicator flags - Default to empty string (or your checkmark value) for clean list rendering
        // You would typically map these to specific columns (e.g., r[18], r[19], etc.)
        forgedKyc: r[18] || '',
        impersonation: r[19] || '',
        fraudCred: r[20] || '',
        suspiciousIp: r[21] || '',
        bankMismatch: r[22] || '',
        immediateWithdraw: r[23] || '',
        nonCooperation: r[24] || '',
    };
}

// -------------------------------
// PDF CREATION
// -------------------------------
// -------------------------------
// PDF CREATION (using photo1.png for header and photo2.png for ENDING footer)
// -------------------------------
// -------------------------------
// PDF CREATION (using photo1.png for header and photo2.png for ENDING footer)
// -------------------------------
const HEADER_HEIGHT = 110; 
// Retaining the buffer definition, though we now draw at doc.y
const FOOTER_SAFE_HEIGHT = 100; 
const FOOTER_IMG = path.join(__dirname, 'public/images', 'photo2.png'); 

async function createPDF(data, outPath) {
    return new Promise((resolve, reject) => {
        try {
            // NOTE: Margins are 50 (left/right/top/bottom)
            const doc = new PDFDocument({ size: 'A4', margin: 50 });
            const stream = fs.createWriteStream(outPath);
            doc.pipe(stream);

            const imgFolder = path.join(__dirname, 'public/images');
            const headerIMG = path.join(imgFolder, 'photo1.png');

            const pageWidth = doc.page.width;
            let headerFound = false;

            // 1. Draw Header on Page 1 (photo1.png)
            try {
                if (fs.existsSync(headerIMG)) {
                    doc.image(headerIMG, 0, 0, { width: pageWidth });
                    headerFound = true;
                }
            } catch (e) {
                console.warn('Header image error. Check photo1.png integrity/size.', e.message);
            }

            // Set starting Y position
            let startY = headerFound ? HEADER_HEIGHT + 20 : 50; 
            doc.x = 50; 
            doc.y = startY; 
            
            // Company info (prefilled constants for Agrim)
            const orgName = 'Agrim Fincap Private Limited';
            const regNo = '07AAACC7658P123 (Example)';
            const regAddress = 'First Floor, 276, Gagan Vihar, Jagat Puri, East Delhi - 110051';
            const emailID = 'support@agrim.com';
            const signatory = data.authSignatory;
            const contactNo = data.contact;

            // Start writing content
            doc.font('Helvetica').fontSize(11);
            
            // Date with gap
            doc.text(`Date: ${moment().format('DD-MM-YYYY')}`, 50, startY);
            doc.moveDown(1.5); 

            // Subject & intro
            doc.text('To,', { continued: true });
            // doc.moveDown(0.2);
            doc.text('The Head,');
            doc.text('Cyber Crime Department,');
            doc.text('New Delhi.');
            doc.moveDown();

            doc.text('Subject: Cyber Fraud Complaint – Request for Action', { 
                width: pageWidth - 100, 
                underline: true 
            }); 
            doc.underline(50, doc.y - 15, pageWidth - 100, doc.y - 10, { color: 'black' });
            doc.moveDown(0.5);
            
            doc.text('Dear Sir/Madam,');
            doc.moveDown(0.5);

            doc.text(
                `I, ${signatory} (Authorized Signatory), representing ${orgName}, am writing to submit a cyber-fraud complaint involving suspected digital impersonation and fraudulent loan activity.`,
                { width: pageWidth - 100, align: 'justify' }
            );
            doc.moveDown();

            // SECTION 1: Complainant details
            doc.text('------------------------------------------------------------');
            doc.text('1. Complainant (NBFC/Lending Institution) Details');
            doc.text('------------------------------------------------------------');
            // doc.text(`• Name of Organization: ${orgName}`);
            // doc.text(`• RBI Certificate of Registration No.: ${regNo}`);
            // doc.text(`• Registered Office Address: ${regAddress}`);
            // doc.text(`• Official Email ID: ${emailID}`, { width: pageWidth - 100 });
            // doc.text(`• Contact Number: ${contactNo}`);
            // doc.text(`• Authorized Signatory: ${signatory}`);
            doc.text(`• Name of Organization: Agrim Fincap Private Limited`);
            doc.text(`• RBI Certificate of Registration No.:07AAACC7658P123 (Example)`);
            doc.text(`• Registered Office Address: First Floor, 276, Gagan Vihar, Jagat Puri, East Delhi - 110051`);
            doc.text(`• Official Email ID:  support@agrim.com`, { width: pageWidth - 100 });
            doc.text(`• Contact Number: +91-90000003`);
            doc.text(`• Authorized Signatory: Aman mishra`);
            doc.moveDown();

            // SECTION 2: Borrower details
            doc.text('------------------------------------------------------------');
            doc.text('2. Borrower / Suspected Fraudster Details');
            doc.text('------------------------------------------------------------');

            doc.text(`• Name (as per KYC): ${data.borrowerName || '-'}`);
            doc.text(`• Mobile Number Used: ${data.mobile || '-'}`);
            doc.text(`• Email ID (if any): ${data.borrowerEmail || '-'}`, { width: pageWidth - 120 });
            doc.text(`• Address Provided: ${data.address || '-'}`);
            doc.text(`• Loan Account ID: ${data.loanId || '-'}`);
            doc.text(`• Loan Amount: ${data.loanAmount || '-'}`);
            doc.text(`• Loan Disbursal Date: ${data.disbursalDate || '-'}`);
            doc.text(`• Bank Account / UPI used: ${data.bankDetails || '-'}`);
            doc.text(`• Aadhaar Last 4 Digits: ${data.aadhaarLast4 || '-'}`);
            doc.text(`• PAN Last 4 Digits: ${data.panLast4 || '-'}`);
            doc.moveDown();

            // SECTION 3: Facts
            doc.text('------------------------------------------------------------');
            doc.text('3. Facts of the Case');
            doc.text('------------------------------------------------------------');
            doc.text(data.facts || 'Details as reported.', { width: pageWidth - 100, align: 'justify' });
            doc.moveDown();

            doc.text('Based on our internal assessment, we strongly suspect:');
            doc.list(['Digital identity fraud', 'Impersonation', 'Misuse of personal details', 'Submission of forged or fabricated documents']);

            doc.moveDown();

            // SECTION 4: Key indicators
            doc.text('------------------------------------------------------------');
            doc.text('4. Key Indicators Observed');
            doc.text('------------------------------------------------------------');
            doc.list([
                `${data.forgedKyc ? '✅' : '☐'} Forged/fake KYC documents`,
                `${data.impersonation ? '✅' : '☐'} Impersonation / stolen identity`,
                `${data.fraudCred ? '✅' : '☐'} Fraudulent mobile/email credentials`,
                `${data.suspiciousIp ? '✅' : '☐'} Suspicious IP/device/location mismatch`,
                `${data.bankMismatch ? '✅' : '☐'} Bank account mismatch`,
                `${data.immediateWithdraw ? '✅' : '☐'} Immediate withdrawal after loan`,
                `${data.nonCooperation ? '✅' : '☐'} Non-cooperation during verification`
            ]);
            doc.moveDown();
            
            // SECTION 5: Legal provisions 
            doc.text('------------------------------------------------------------');
            doc.text('5. Applicable Legal Provisions – Under Bhartiya Nyaya Sanhita (BNS), 2023');
            doc.text('------------------------------------------------------------');
            doc.text('Section 318 – Cheating by personation (impersonation)');
            doc.text('Section 316 – Cheating and dishonestly inducing delivery of property');
            doc.text('Section 334 – Dishonest misappropriation of property');
            doc.text('Section 338 – Forgery of valuable security or electronic record');
            doc.text('Section 339 – Using forged documentation as genuine');
            doc.text('Section 111(3) – Conspiracy in commission of offense');
            doc.text('Section 356 – Criminal breach of trust');
            doc.moveDown(0.5);
            doc.text('IT Act 2000:');
            doc.list(['Section 66C – Identity theft', 'Section 66D – Cheating by impersonation using computer resources', 'Section 72 – Breach of privacy']);
            doc.moveDown(0.5);
            doc.text('RBI Guidelines:');
            doc.list(['Digital Lending Guidelines 2022', 'RBI KYC Master Directions', 'Data Privacy Framework']);
            doc.moveDown();

            // SECTION 6: Action requested
            doc.text('------------------------------------------------------------');
            doc.text('6. Action Requested');
            doc.text('------------------------------------------------------------');
            doc.text('We request the Cyber Crime Department to register a cyber-fraud case, conduct IP/device forensics, identify the fraudster, trace fund trail, prevent further misuse, and support legal recovery.', { width: pageWidth - 100, align: 'justify' });
            doc.moveDown();

            // SECTION 7: Documents
            doc.text('------------------------------------------------------------');
            doc.text('7. Documents Attached');
            doc.text('------------------------------------------------------------');
            doc.list(['KYC documents', 'Digital logs (IP/device/location)', 'Bank/UPI trail', 'Screenshots & evidence', 'Internal investigation report']);
            doc.moveDown();

            // SECTION 8: Declaration & signature
            doc.text('------------------------------------------------------------');
            doc.text('8. Declaration');
            doc.text('------------------------------------------------------------');
            doc.text('I hereby declare that the information provided is true and correct to the best of my knowledge and request urgent action.', { width: pageWidth - 100, align: 'justify' });
            doc.moveDown(2);

            doc.text('Regards,');
            doc.text('_________________________'); 
            doc.text(signatory);
            doc.text('(Authorized Signatory)');
            doc.text(orgName);
            doc.text(`Contact: ${contactNo}`);
            doc.text(`Date: ${moment().format('DD-MM-YYYY')}`);


            // FINAL FIX: Draw image at current text position (doc.y) to follow content
            if (fs.existsSync(FOOTER_IMG)) {
                // Add some space after the last signature line
                doc.moveDown(1); 
                
                // Draw the image: full width (pageWidth). 
                // By drawing it at doc.y, it flows right after the text. 
                // pdfkit will automatically calculate the height and add a new page if needed.
                doc.image(FOOTER_IMG, 0, doc.y, { width: pageWidth });
            }

            // Finish PDF
            doc.end();

            // Logging logic (after stream finishes)
            stream.on('finish', () => {
                try {
                    const logsFile = path.join(__dirname, 'logs.json');
                    let logs = [];
                    if (fs.existsSync(logsFile)) {
                        try { logs = JSON.parse(fs.readFileSync(logsFile, 'utf8')); } catch (e) { logs = []; }
                    }
                    logs.push({
                        loanId: data.loanId || '',
                        borrower: data.borrowerName || '',
                        generatedAt: moment().format('YYYY-MM-DD HH:mm:ss'),
                        status: 'PDF Created'
                    });
                    const safeLogs = logs.slice(-500);
                    fs.writeFileSync(logsFile, JSON.stringify(safeLogs, null, 2));
                } catch (e) {
                    console.warn('Failed to write generation log', e.message);
                }
                resolve();
            });

            stream.on('error', err => reject(err));
        } catch (err) {
            console.error('PDF Generation Error:', err.message);
            reject(err);
        }
    });
}


// -------------------------------
// START SERVER
// -------------------------------
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running at http://localhost:${PORT}`));