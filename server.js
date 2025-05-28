// server.js
import express from 'express';
import multer  from 'multer';
import { run } from './invite.js';
import fs from 'fs/promises';
import path from 'path';

const app   = express();
const store = multer({ dest: 'uploads/' });
const PORT  = 3080;

// Store uploaded rows in memory for session (simple approach)
const uploads = {};

app.use(express.urlencoded({ extended: true }));

app.post('/upload', store.single('sheet'), async (req, res) => {
  const name = path.basename(req.file.path); // safe id
  try {
    // Validation phase: read and analyze the file, but do not send invites yet
    const ExcelJS = (await import('exceljs')).default;
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(req.file.path);
    const ws = wb.getWorksheet('Bookings');
    const buildHeaderMap = row => {
      const map = {};
      row.values.forEach((v, i) => {
        if (v === null || v === undefined) return;
        const key = v.toString().trim().toLowerCase();
        if (key) map[key] = i;
      });
      return map;
    };
    const col = (header, name) => header[name.toLowerCase()];
    // Safe cell access to avoid "A Cell needs a Row" errors
    const getCellText = (name, row) => {
      const idx = col(header, name);
      return idx ? row.getCell(idx).text : '';
    };
    const getCellValue = (name, row) => {
      const idx = col(header, name);
      return idx ? row.getCell(idx).value : undefined;
    };
    const parseEmails = text => (text || '').split(/[,;]/).map(e => e.trim()).filter(Boolean);
    const header = buildHeaderMap(ws.getRow(1));
    let rows = [];
    let missingFields = false;
    let filledRows = 0;
    for (let i = 2; i <= ws.rowCount; i++) {
      if (i === 2) continue;
      const r = ws.getRow(i);
      if (!r || !r.hasValues) break;
      // If the row is completely empty, stop processing further rows
      const isEmpty = r.values.every(v => v === null || v === undefined || v === '');
      if (isEmpty) break;
      // Only count as filled if at least one essential field is present
      const topic = getCellText('meeting name', r);
      const teamName= getCellText('Team Name', r);
      const dateChosen = getCellText('date chosen', r);
      const team = getCellText('team invitees', r);
      const zoomAccount = getCellText('Zoom Account to book from', r) || process.env.OUTLOOK_HOST;
      const durationMin = getCellValue('length of meeting(minutes)', r);
      const meetingType = getCellText('meeting type', r).toLowerCase() || 'zoom';
      const participants = parseEmails([
        getCellText('team invitees', r),
        getCellText('added invitee 1', r),
        getCellText('added invitee 2', r),
        getCellText('added invitee 3', r),
        getCellText('added invitee 4', r),
        getCellText('email from form', r)
      ].join(','));
      if (topic || dateChosen || team || zoomAccount || participants.length > 0) filledRows++;
      if ((getCellText('meeting booked', r).toLowerCase()) === 'yes') continue;
      // Check for essential fields
      const essentials = {
        'Meeting Name': topic,
        'Date Chosen': dateChosen,
        'Team Invitees': team,
        'Zoom Account to book from': zoomAccount,
        'Invitees': participants.length > 0 ? participants.join(', ') : ''
      };
      let missing = Object.entries(essentials).filter(([k, v]) => !v);
      if (missing.length) missingFields = true;
      rows.push({
        team,
        teamName,
        topic,
        dateChosen,
        durationMin,
        zoomAccount,
        participants: participants.join(', '),
        meetingType,
        missing: missing.map(([k]) => k)
      });
    }
    // Save rows for later processing
    uploads[name] = rows;
    // Render validation summary as modern HTML table and buttons
    let html = `
    <style>
      body { font-family: 'Segoe UI', Arial, sans-serif; background: #f7f8fa; margin: 0; padding: 0; }
      .container { max-width: 1000px; margin: 40px auto; background: #fff; border-radius: 18px; box-shadow: 0 4px 32px #0002; padding: 40px 18px; }
      h2 { margin-top: 0; font-size: 2.4em; letter-spacing: -1px; font-weight: 800; }
      .summary { font-size: 1.15em; margin-bottom: 22px; color: #222b; }
      .table-wrap { width: 100%; overflow-x: auto; }
      table { width: 100%; border-collapse: collapse; margin-bottom: 22px; background: #f9fafb; border-radius: 10px; overflow: hidden; min-width: 700px; }
      th, td { padding: 12px 10px; text-align: left; }
      th { background: #f1f5f9; font-weight: 700; border-bottom: 2px solid #e2e8f0; }
      tr { transition: background 0.2s; }
      tr:hover { background: #f1f5ff; }
      tr.missing { background: #ffe0e0; }
      td { border-bottom: 1px solid #e2e8f0; }
      .actions { display: flex; gap: 16px; margin-top: 18px; flex-wrap: wrap; }
      button, .btn { background: #2563eb; color: #fff; border: none; border-radius: 8px; padding: 12px 28px; font-size: 1.08em; font-weight: 600; cursor: pointer; transition: background 0.2s, box-shadow 0.2s; box-shadow: 0 2px 8px #2563eb22; }
      button[disabled], .btn[disabled] { background: #b0b8c9; cursor: not-allowed; box-shadow: none; }
      button.cancel, .btn.cancel { background: #e53e3e; }
      .note { color: #e53e3e; margin-top: 12px; font-size: 1.13em; }
      .meeting-type-select { font-size: 1em; padding: 6px 10px; border-radius: 6px; border: 1px solid #cbd5e1; margin-right: 10px; }
      .location-input { font-size: 1em; padding: 6px 10px; border-radius: 6px; border: 1px solid #cbd5e1; margin-left: 10px; width: 220px; }
      .global-controls { margin-bottom: 18px; display: flex; align-items: center; gap: 16px; }
      @media (max-width: 900px) {
        .container { padding: 8px; }
        .table-wrap { padding-bottom: 8px; }
        table, th, td { font-size: 0.97em; }
        .actions { flex-direction: column; gap: 8px; }
        .global-controls { flex-direction: column; align-items: flex-start; gap: 8px; }
      }
      @media (max-width: 600px) {
        .container { padding: 2px; border-radius: 0; box-shadow: none; }
        h2 { font-size: 1.5em; }
        .summary { font-size: 1em; }
        table, th, td { font-size: 0.93em; }
        th, td { padding: 7px 4px; }
        .location-input { width: 100%; margin-left: 0; }
      }
    </style>
    <div class="container">
      <h2>Validation Summary</h2>
      <div class="summary">Total invites to send: <b>${rows.length}</b> (Rows filled: <b>${filledRows}</b>)</div>
      <form action="/process" method="POST" id="validationForm">
        <input type="hidden" name="id" value="${name}">
        <input type="hidden" name="meetingType" value="${req.body?.meetingType || 'zoom'}">
        <input type="hidden" name="location" value="${req.body?.location || ''}">
        <div class="table-wrap">
        <table>
          <tr><th>Team</th><th>Meeting Name</th><th>Date</th><th>Duration</th><th>Invitees</th><th>Account</th><th>Missing</th></tr>`;
    rows.forEach((row) => {
      html += `<tr class="${row.missing.length ? 'missing' : ''}">
        <td>${row.team || ''}</td>
        <td>${row.topic || ''}</td>
        <td>${row.dateChosen || ''}</td>
        <td>${row.durationMin || ''}</td>
        <td style="max-width:220px;overflow-wrap:break-word;">${row.participants || ''}</td>
        <td>${row.zoomAccount || ''}</td>
        <td>${row.missing.join(', ')}</td>
      </tr>`;
    });
    html += `</table></div>
        <div class="actions">
          <button type="submit" id="submitBtn"${missingFields ? ' disabled' : ''}>Send Invites</button>
          <button type="button" class="cancel" onclick="window.location='/'">Cancel</button>
        </div>
      </form>
    </div>`;
    res.send(html);
  } catch (err) {
    console.error(err);
    res.status(500).send('<p>❌ Error – check server log.</p>');
  }
});

// Modern Home Page
app.get('/', (_, res) => res.send(`
  <style>
    body { font-family: 'Segoe UI', Arial, sans-serif; background: #f7f8fa; margin: 0; padding: 0; }
    .container { max-width: 500px; margin: 80px auto; background: #fff; border-radius: 16px; box-shadow: 0 4px 32px #0002; padding: 48px 32px; text-align: center; }
    h1 { font-size: 2.5em; font-weight: 800; margin-bottom: 0.2em; letter-spacing: -1px; }
    p { color: #444b; font-size: 1.15em; margin-bottom: 2em; }
    .upload-form { margin-top: 2em; }
    input[type=file] { display: block; margin: 0 auto 1.5em auto; font-size: 1.1em; }
    button { background: #2563eb; color: #fff; border: none; border-radius: 8px; padding: 14px 36px; font-size: 1.1em; font-weight: 600; cursor: pointer; transition: background 0.2s, box-shadow 0.2s; box-shadow: 0 2px 8px #2563eb22; }
    button:hover { background: #1749b1; }
    .radio-group { margin: 1.5em 0 1em 0; display: flex; justify-content: center; gap: 2em; }
    .location-wrap { margin: 1em 0; display: none; }
    .location-wrap.active { display: block; }
    @media (max-width: 600px) {
      .container { padding: 12px; }
      h1 { font-size: 2em; }
    }
  </style>
  <div class="container">
    <h1>Zoom/In-Person Invite Uploader</h1>
    <p>Upload your Doodle Excel sheet to validate and send calendar invites.<br>Choose meeting type below.</p>
    <form class="upload-form" action="/upload" method="POST" enctype="multipart/form-data" id="mainForm">
      <div class="radio-group">
        <label><input type="radio" name="meetingType" value="zoom" checked> Zoom meeting</label>
        <label><input type="radio" name="meetingType" value="in-person"> In person meeting</label>
      </div>
      <div class="location-wrap" id="locationWrap">
        <input type="text" name="location" id="locationInput" placeholder="Enter location name" style="width: 90%; padding: 10px; border-radius: 6px; border: 1px solid #cbd5e1; font-size: 1.1em;">
      </div>
      <input type="file" name="sheet" accept=".xlsx" required>
      <button>Upload & Validate</button>
    </form>
    <script>
      const radios = document.querySelectorAll('input[name="meetingType"]');
      const locationWrap = document.getElementById('locationWrap');
      radios.forEach(r => r.addEventListener('change', () => {
        if (document.querySelector('input[name="meetingType"]:checked').value === 'in-person') {
          locationWrap.classList.add('active');
        } else {
          locationWrap.classList.remove('active');
        }
      }));
    </script>
  </div>
`));

// Modern Completed Page
app.post('/process', async (req, res) => {
  const { id, meetingType, location } = req.body;
  const rows = uploads[id];
  if (!rows) return res.status(400).send('Session expired. Please re-upload.');
  let sent = 0, failed = 0, errors = [];
  for (const row of rows) {
    if (row.missing && row.missing.length) continue;
    try {
      // Pass meetingType and location to invite logic
      await run({ ...row, teamName: row.teamName, meetingType, location });
      sent++;
    } catch (err) {
      failed++;
      errors.push({ row, error: err.message });
    }
  }
  res.send(`<div class='container'><h2>Invites sent: ${sent}, Failed: ${failed}</h2>${errors.length ? '<pre>' + errors.map(e => e.error).join('\n') + '</pre>' : ''}<br><a href='/'>Back to Home</a></div>`);
});

app.listen(PORT, () => console.log(`Open http://localhost:${PORT}`));