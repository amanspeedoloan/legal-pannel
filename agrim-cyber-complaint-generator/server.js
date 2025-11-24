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

// ---------------------------------
// Ensure required folders exist
// ---------------------------------
['uploads', 'generated', 'public/images', 'public/css', 'public/js'].forEach(dir => {
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
}); 

// Serve static assets from public
app.use(express.static('public'));

// Parse urlencoded forms (for form posts)
app.use(express.urlencoded({ extended: true }));

// -------------------------------
// Optional: Serve an uploaded logo (from your session)
// Developer note: file path taken from your session's upload.
// This makes testing easier — you can change the path or copy the file into public/images.
// -------------------------------
const UPLOADED_LOGO_PATH = '/mnt/data/ccbf460a-4240-4f0a-8fac-c4c116509143.png'; // provided in session
app.get('/uploaded-logo', (req, res) => {
  if (fs.existsSync(UPLOADED_LOGO_PATH)) return res.sendFile(UPLOADED_LOGO_PATH);
  return res.status(404).send('Uploaded logo not found on server (copy it to public/images to use it permanently).');
});

// -------------------------------
// Routes
// -------------------------------

// Dashboard (loads index.html from public)
app.get('/', (req, res) => {
  res.sendFile(path.resolve(__dirname, 'public/index.html'));
});

// Template Manager (Placeholder)
app.get('/templates', (req, res) => {
  let html = `<html><head><meta charset="utf-8"><title>Template Manager</title><link rel="stylesheet" href="/css/style.css"></head><body>`;
  html += `<div class="card" style="margin:20px;">`;
  // Back to Dashboard button/link at the top
  html += `<div style="margin-bottom:12px;"><a href="/">← Back to Dashboard</a></div>`;
  html += `<h2>Template Manager</h2>`;
  html += `<p>This page is currently a placeholder for managing PDF templates.</p>`;
  html += `<div style="margin-top:20px"><a href="/">← Return to Dashboard</a></div></div></body></html>`;
  res.send(html);
});

// Download Excel template
app.get('/download-template', async (req, res) => {
  try {
    const wb = new exceljs.Workbook();
    const ws = wb.addWorksheet('Template');

    const headers = [
      'Organization Name','RBI Reg No','Address','Email','Contact','Signatory',
      'Borrower Name','Mobile','Borrower Email','Borrower Address','Loan ID','Loan Amount',
      'Disbursal Date','Bank/UPI','Aadhaar Last 4','PAN Last 4','Facts',
      'Forged KYC (Y/N)', 'Impersonation (Y/N)', 'Fraud Cred (Y/N)', 'Suspicious IP (Y/N)',
      'Bank Mismatch (Y/N)', 'Immediate Withdraw (Y/N)', 'Non-Cooperation (Y/N)'
    ];
    ws.addRow(['', ...headers]);
    ws.addRow([
      '','Agrim Fincap Pvt Ltd','07AAACV...','Address line','info@agrim.com','+91-XXXXXX','Authorized Person',
      'Rahul Kumar','9876543210','rahul@mail.com','rahul@mail.com','Delhi','LN10203','15000',
      moment().format('YYYY-MM-DD'),'HDFC/UPI','1234','ABCD','Suspected fake KYC',
      'Y', 'N', 'Y', 'N', 'N', 'Y', 'N'
    ]);

    res.setHeader('Content-Disposition', 'attachment; filename=agrim_template.xlsx');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    await wb.xlsx.write(res);
    res.end();
  } catch (e) {
    console.error('Template generation error', e);
    res.status(500).send('Template generation failed');
  }
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
    const dataRows = rows.slice(1).filter(r => r && r.length > 1);
    
    // Build preview HTML (simple, server-rendered)
    let html = `<html><head><meta charset="utf-8"><title>Preview</title><link rel="stylesheet" href="/css/style.css"></head><body>`;
    html += `<div class="card" style="margin:20px;">`;
    // Added: Back to Dashboard button/link at the top
    html += `<div style="margin-bottom:12px;"><a href="/">← Back to Dashboard</a></div>`;
    html += `<h2>Excel Preview (${dataRows.length} rows)</h2>`;
    html += `<table class="preview-table"><thead><tr>`;
    for (let i = 1; i < headers.length; i++) html += `<th>${headers[i] || `Col ${i}`}</th>`;
    html += `</tr></thead><tbody>`;

    dataRows.forEach(r => {
      html += `<tr>`;
      for (let i = 1; i < headers.length; i++) html += `<td>${r[i] !== undefined && r[i] !== null ? String(r[i]) : '-'}</td>`;
      html += `</tr>`;
    });

    html += `</tbody></table>`;

    // Hidden form posts back file path to generate-all
    html += `<form method="POST" action="/generate-all" style="margin-top:20px;">
               <input type="hidden" name="filePath" value="${req.file.path}">
               <button class="btn primary" style="padding:10px 14px">Generate All PDFs (${dataRows.length})</button>
             </form>`;

    html += `<div style="margin-top:12px"><a href="/">← Back to Dashboard</a></div></div></body></html>`;
    res.send(html);
  } catch (err) {
    console.error('Upload-excel error:', err);
    if (req.file && fs.existsSync(req.file.path)) {
      try { fs.unlinkSync(req.file.path); } catch (e) {}
    }
    res.status(500).send(`Excel read error. Details: ${err.message}`);
  }
});

// Generate all PDFs from uploaded Excel
app.post('/generate-all', async (req, res) => {
  let successCount = 0;
  const filePath = req.body.filePath;
  let dataRows = [];

  try {
    if (!filePath || !fs.existsSync(filePath)) {
      console.error('ERROR: Missing or invalid filePath received for generation:', filePath);
      return res.status(400).send('Uploaded file not found. Please re-upload the file and try again.');
    }

    const workbook = new exceljs.Workbook();
    await workbook.xlsx.readFile(filePath);
    const sheet = workbook.worksheets[0];
    const rows = [];
    sheet.eachRow(r => rows.push(r.values));
    dataRows = rows.slice(1).filter(r => r && r.length > 1);

    console.log(`Starting PDF generation for ${dataRows.length} rows...`);

    for (const r of dataRows) {
      try {
        const data = extractExcelRow(r);
        const name = `Complaint_${sanitize(data.loanId || moment().format('YYYYMMDDHHmmss'))}.pdf`;
        await createPDF(data, path.join('generated', name));
        successCount++;
      } catch (innerErr) {
        const loanId = (r && r[11]) || 'N/A';
        console.error(`Failed to generate PDF for Loan ID: ${loanId}. Error: ${innerErr.message}`);
      }
    }

    // remove uploaded file
    try { fs.unlinkSync(filePath); } catch (e) { console.warn('Failed to delete uploaded file:', e.message); }

    console.log(`Finished PDF generation. Success count: ${successCount}`);

    return res.redirect(`/view-letters?generated=${successCount}`);
  } catch (err) {
    console.error('Bulk generation setup/file error:', err);
    if (filePath && fs.existsSync(filePath)) {
      try { fs.unlinkSync(filePath); } catch (e) {}
    }
    return res.status(500).send(`Generation error (processed ${successCount} files). Details: ${err.message}`);
  }
});

// View generated letters (list) with optional simple server-side search (q)
app.get('/view-letters', (req, res) => {
  const generatedCount = parseInt(req.query.generated) || 0;
  const q = (req.query.q || '').toLowerCase().trim();

  let allFiles = [];
  if (fs.existsSync('generated')) {
    allFiles = fs.readdirSync('generated')
      .filter(f => f.endsWith('.pdf'))
      .map(f => ({
        name: f,
        path: path.join('generated', f),
        mtime: fs.statSync(path.join('generated', f)).mtime.getTime()
      }))
      .sort((a, b) => b.mtime - a.mtime);
  }

  // If query present, filter server-side
  if (q) {
    allFiles = allFiles.filter(f => f.name.toLowerCase().includes(q));
  }

  const totalFiles = allFiles.length;

  let html = `<html><head><meta charset="utf-8"><title>Generated</title><link rel="stylesheet" href="/css/style.css"></head><body>`;
  html += `<div class="card" style="margin:20px;">`;
  // Added: Back to Dashboard button/link at the top
  html += `<div style="margin-bottom:12px;"><a href="/">← Back to Dashboard</a></div>`;
  html += `<h2>Generated Letters (${totalFiles})</h2>`;

  if (generatedCount > 0) {
    html += `<div class="alert success">✅ Successfully generated ${generatedCount} new Complaint letter(s).</div>`;
  } else if (generatedCount === 0 && req.query.generated !== undefined) {
    html += `<div class="alert error">❌ 0 letters generated. Check server logs or source Excel data.</div>`;
  }

  html += `
    <div style="margin:12px 0">
      <form method="GET" action="/view-letters" style="display:flex;gap:8px;align-items:center">
        <input name="q" placeholder="Search PDF name..." value="${escapeHtml(req.query.q || '')}" style="padding:8px;border-radius:6px;border:1px solid #ddd;width:320px">
        <button class="btn outline" style="padding:8px 12px">Search</button>
        <a href="/view-letters" class="btn outline" style="padding:8px 12px">Reset</a>
      </form>
    </div>
  `;

  if (!allFiles.length) {
    html += `<p class="muted">No letters generated yet.</p>`;
  } else {
    html += `<table class="preview-table" style="width:100%;margin-top:12px"><thead><tr><th>PDF Name</th><th>Status</th><th>View</th><th>Download</th><th>Delete</th></tr></thead><tbody>`;

    allFiles.forEach((file, index) => {
      const isNew = index < generatedCount;
      const rowClass = isNew ? 'highlight-row' : '';
      html += `<tr class="${rowClass}">
        <td>${escapeHtml(file.name)}</td>
        <td>${isNew ? '✨ NEWLY GENERATED' : 'Older File'}</td>
        <td><a href="/generated/${encodeURIComponent(file.name)}" target="_blank" class="btn small view">View</a></td>
        <td><a href="/download/${encodeURIComponent(file.name)}" class="btn small download">Download</a></td>
        <td>
          <form method="POST" action="/delete-generated/${encodeURIComponent(file.name)}" style="display:inline">
            <button class="btn outline small" onclick="return confirm('Delete ${escapeHtml(file.name)}?');">Delete</button>
          </form>
        </td>
      </tr>`;
    });

    html += `</tbody></table>`;
  }

  html += `<div style="margin-top:12px"><a href=\"/\">← Back to Dashboard</a> | <a href=\"/download-logs\">View Download Logs</a> | <a href=\"/generation-logs\">View Generation Logs</a></div></div></body></html>`;
  res.send(html);
});

// Serve generated PDFs for viewing (static)
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

// Download logs viewer
app.get('/download-logs', (req, res) => {
  const logFile = path.join(__dirname, "download_logs.json");
  let logs = [];
  if (fs.existsSync(logFile)) {
    try { logs = JSON.parse(fs.readFileSync(logFile, "utf8")).reverse(); } catch (e) { logs = []; }
  }

  let html = `<html><head><meta charset="utf-8"><title>Download Logs</title><link rel="stylesheet" href="/css/style.css"></head><body>`;
  html += `<div class="card" style="margin:20px;">`;
  // Added: Back to Dashboard button/link at the top
  html += `<div style="margin-bottom:12px;"><a href="/">← Back to Dashboard</a></div>`;
  html += `<h2>Download Logs (${logs.length} entries)</h2>`;
  if (!logs.length) html += `<p>No downloads yet.</p>`;
  else {
    html += `<table class="preview-table"><thead><tr><th>PDF Name</th><th>Downloaded At</th><th>View</th></tr></thead><tbody>`;
    logs.forEach(l => {
      html += `<tr><td>${escapeHtml(l.filename)}</td><td>${escapeHtml(l.downloadedAt)}</td><td><a href="/generated/${encodeURIComponent(l.filename)}" target="_blank">View</a></td></tr>`;
    });
    // FIX: Close table and wrap navigation link properly
    html += `</tbody></table>`;
    html += `<div style="margin-top:12px"><a href="/">← Back to Dashboard</a></div>`; 
  }
  html += `</div></body></html>`;
  res.send(html);
});

// Generation logs viewer
app.get('/generation-logs', (req, res) => {
  const logFile = path.join(__dirname, "logs.json");
  let logs = [];
  if (fs.existsSync(logFile)) {
    try { logs = JSON.parse(fs.readFileSync(logFile, "utf8")).reverse(); } catch (e) { logs = []; }
  }

  let html = `<html><head><meta charset="utf-8"><title>Generation Logs</title><link rel="stylesheet" href="/css/style.css"></head><body>`;
  html += `<div class="card" style="margin:20px;">`;
  // Added: Back to Dashboard button/link at the top
  html += `<div style="margin-bottom:12px;"><a href="/">← Back to Dashboard</a></div>`;
  html += `<h2>Generation Logs (${logs.length} entries)</h2>`;
  if (!logs.length) html += `<p>No generations yet.</p>`;
  else {
    html += `<table class="preview-table"><thead><tr><th>Loan ID</th><th>Borrower</th><th>Generated At</th><th>Status</th></tr></thead><tbody>`;
    logs.forEach(l => {
      html += `<tr><td>${escapeHtml(l.loanId)}</td><td>${escapeHtml(l.borrower)}</td><td>${escapeHtml(l.generatedAt)}</td><td>${escapeHtml(l.status)}</td></tr>`;
    });
    // FIX: Close table and wrap navigation link properly
    html += `</tbody></table>`;
    html += `<div style="margin-top:12px"><a href="/">← Back to Dashboard</a></div>`;
  }
  html += `</div></body></html>`;
  res.send(html);
});

// -------------------------------
// Delete endpoints (file-based)
// -------------------------------

// Delete a single generated PDF (file-based)
app.post('/delete-generated/:filename', (req, res) => {
  try {
    const decoded = decodeURIComponent(req.params.filename);
    const p = path.join(__dirname, 'generated', decoded);
    if (!fs.existsSync(p)) return res.status(404).send('File not found');
    fs.unlinkSync(p);
    return res.redirect('/view-letters');
  } catch (e) {
    console.error('Delete single generated error', e);
    return res.status(500).send('Delete failed');
  }
});

// Delete all generated PDFs (file-based)
app.post('/delete-all-generated', (req, res) => {
  try {
    const dir = path.join(__dirname, 'generated');
    if (!fs.existsSync(dir)) return res.sendStatus(200);
    const files = fs.readdirSync(dir).filter(f => f.endsWith('.pdf'));
    files.forEach(f => {
      try { fs.unlinkSync(path.join(dir, f)); } catch (e) {}
    });
    return res.sendStatus(200);
  } catch (e) {
    console.error('Delete all generated error', e);
    return res.status(500).send('Clear failed');
  }
});

// -------------------------------
// Helpers & PDF creation
// -------------------------------
function sanitize(s) {
  return String(s || '').replace(/[\\/:*?"<>|]/g, '_');
}

function escapeHtml(text) {
  if (!text) return '';
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function extractExcelRow(r) {
  const formatDisbursalDate = (val) => {
    if (!val) return '';
    if (moment.isDate(val)) {
      return moment(val).format('DD-MM-YYYY');
    }
    let dateVal = String(val);
    if (!isNaN(dateVal) && Number(dateVal) > 40000 && Number(dateVal) < 70000) {
      return moment(new Date(1899, 11, 30)).add(Number(dateVal), 'days').format('DD-MM-YYYY');
    }
    return dateVal;
  };

  const checkFlag = (val) => String(val || '').toUpperCase() === 'Y' ? 'Y' : '';

  return {
    email: r[4] || 'support@agrim.com',
    contact: r[5] || '9876543210',
    authSignatory: r[6] || 'Aman Mishra',
    borrowerName: r[7] || '',
    mobile: r[8] || '',
    borrowerEmail: r[9] || '',
    address: r[10] || '',
    loanId: r[11] || '',
    loanAmount: r[12] || '',
    disbursalDate: formatDisbursalDate(r[13]),
    bankDetails: r[14] || '',
    aadhaarLast4: r[15] || '',
    panLast4: r[16] || '',
    facts: r[17] || 'Suspected cyber fraud with details below.',
    forgedKyc: checkFlag(r[18]),
    impersonation: checkFlag(r[19]),
    fraudCred: checkFlag(r[20]),
    suspiciousIp: checkFlag(r[21]),
    bankMismatch: checkFlag(r[22]),
    immediateWithdraw: checkFlag(r[23]),
    nonCooperation: checkFlag(r[24]),
  };
}

// PDF Creation logic (returns Promise)
const HEADER_HEIGHT = 110;
const FOOTER_IMG = path.join(__dirname, 'public/images', 'photo2.png'); // default footer from public images

async function createPDF(data, outPath) {
  return new Promise((resolve, reject) => {
    try {
      // Removed erroneous line

      const doc = new PDFDocument({ size: 'A4', margin: 50 });
      const stream = fs.createWriteStream(outPath);
      doc.pipe(stream);

      const imgFolder = path.join(__dirname, 'public/images');
      const headerIMG = path.join(imgFolder, 'photo1.png');

      const pageWidth = doc.page.width;
      let headerFound = false;

      // try header image
      try {
        if (fs.existsSync(headerIMG)) {
          doc.image(headerIMG, 0, 0, { width: pageWidth });
          headerFound = true;
        }
      } catch (e) {
        console.warn('Header image error. Check photo1.png integrity/size.', e.message);
      }

      let startY = headerFound ? HEADER_HEIGHT + 20 : 50;
      doc.x = 50;
      doc.y = startY;

      const orgName = 'Agrim Fincap Private Limited';
      const regNo = '07AAACC7658P123 (Example)';
      const regAddress = 'First Floor, 276, Gagan Vihar, Jagat Puri, East Delhi - 110051';
      const emailID = 'support@agrim.com';
      const signatory = data.authSignatory;
      const contactNo = data.contact;

      doc.font('Helvetica').fontSize(11);
      doc.text(`Date: ${moment().format('DD-MM-YYYY')}`, 50, startY);
      doc.moveDown(1.5);

      doc.text('To,', { continued: true });
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
        `I, ${signatory} (Authorized Signitory), representing ${orgName}, am writing to submit a cyber-fraud complaint involving suspected digital impersonation and fraudulent loan activity.`,
        { width: pageWidth - 100, align: 'justify' }
      );
      doc.moveDown();

      // Complainant details
      doc.text('------------------------------------------------------------');
      doc.text('1. Complainant (NBFC/Lending Institution) Details');
      doc.text('------------------------------------------------------------');
      doc.text(`• Name of Organization: ${orgName}`);
      doc.text(`• RBI Certificate of Registration No.: ${regNo}`);
      doc.text(`• Registered Office Address: ${regAddress}`);
      doc.text(`• Official Email ID: ${emailID}`, { width: pageWidth - 100 });
      doc.text(`• Contact Number: ${contactNo}`);
      doc.text(`• Authorized Signitory: ${signatory}`);
      doc.moveDown();

      // Borrower details
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

      // Facts
      doc.text('------------------------------------------------------------');
      doc.text('3. Facts of the Case');
      doc.text('------------------------------------------------------------');
      doc.text(data.facts || 'Details as reported.', { width: pageWidth - 100, align: 'justify' });
      doc.moveDown();

      doc.text('Based on our internal assessment, we strongly suspect:');
      doc.list(['Digital identity fraud', 'Impersonation', 'Misuse of personal details', 'Submission of forged or fabricated documents']);
      doc.moveDown();

      // Key indicators
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

      // Legal provisions
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

      // Action requested
      doc.text('------------------------------------------------------------');
      doc.text('6. Action Requested');
      doc.text('------------------------------------------------------------');
      doc.text('We request the Cyber Crime Department to register a cyber-fraud case, conduct IP/device forensics, identify the fraudster, trace fund trail, prevent further misuse, and support legal recovery.', { width: pageWidth - 100, align: 'justify' });
      doc.moveDown();

      // Documents
      doc.text('------------------------------------------------------------');
      doc.text('7. Documents Attached');
      doc.text('------------------------------------------------------------');
      doc.list(['KYC documents', 'Digital logs (IP/device/location)', 'Bank/UPI trail', 'Screenshots & evidence', 'Internal investigation report']);
      doc.moveDown();

      // Declaration & signature
      doc.text('------------------------------------------------------------');
      doc.text('8. Declaration');
      doc.text('------------------------------------------------------------');
      doc.text('I hereby declare that the information provided is true and correct to the best of my knowledge and request urgent action.', { width: pageWidth - 100, align: 'justify' });
      doc.moveDown(2);

      doc.text('Regards,');
      doc.text('_________________________');
      doc.text(signatory);
      doc.text('(Authorized Signitory)');
      doc.text(orgName);
      doc.text(`Contact: ${contactNo}`);
      doc.text(`Date: ${moment().format('DD-MM-YYYY')}`);

      // Footer image
      if (fs.existsSync(FOOTER_IMG)) {
        doc.moveDown(1);
        try { doc.image(FOOTER_IMG, 0, doc.y, { width: pageWidth }); } catch (e) { /* ignore image errors */ }
      }

      doc.end();

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
// Start server
// -------------------------------
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running at http://localhost:${PORT}`));