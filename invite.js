import ExcelJS from 'exceljs';
import axios from 'axios';
import dotenv from 'dotenv';
import fs from 'fs';
import { graphClient, sendInvite } from './graph.js';
import dayjs from './dayjs-setup.js';

dotenv.config();
const TZ = process.env.TIMEZONE || 'UTC';
const ZOOM = { account: process.env.ACCOUNT_ID, id: process.env.CLIENT_ID, secret: process.env.CLIENT_SECRET };

/**
 * Expected headers in the Doodle sheet:
 *   Meeting Name | Date Chosen | Length of Meeting(minutes) | Meeting Name For Invite
 *   Team invitees | Added Invitee 1‑4 | Email from form | Meeting Booked | JoinURL
 */

function buildHeaderMap(row) {
  const map = {};
  row.values.forEach((v, i) => {
    if (v === null || v === undefined) return;    // skip blanks
    const key = v.toString().trim().toLowerCase();
    if (key) map[key] = i;
  });
  return map;
}

function col(header, name) {
  return header[name.toLowerCase()];
}

function parseEmails(text) {
  return (text || '').split(/[,;]/).map(e => e.trim()).filter(Boolean);
}

async function zToken() {
  const { data } = await axios.post(
    'https://zoom.us/oauth/token',
    null,
    {
      params: { grant_type: 'account_credentials', account_id: ZOOM.account },
      auth: { username: ZOOM.id, password: ZOOM.secret }
    });
  return data.access_token;
}

async function zCreate(token, topic, startISO, duration, agenda, zoomAccount) {
  const { data } = await axios.post(
    `https://api.zoom.us/v2/users/${zoomAccount}/meetings`,
    { topic, type: 2, start_time: startISO, duration, timezone: 'Asia/Beirut', agenda },
    { headers: { Authorization: `Bearer ${token}` } });
  return data;
}

(async () => {
  if (process.argv.length <= 1 || !process.argv[2]) return; // Only run if a file is provided as argument
  const file = process.argv[2];
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(file);
  const ws = wb.getWorksheet('Bookings');

  const header = buildHeaderMap(ws.getRow(1));
  const token = await zToken();
  const graph = await graphClient();

  if (!header['joinurl']) {
    header['joinurl'] = ws.columnCount + 1;
    ws.getRow(1).getCell(header['joinurl']).value = 'JoinURL';
  }

  for (let i = 2; i <= ws.rowCount; i++) {
    const r = ws.getRow(i);
    if (i === 2) continue; // skip Doodle dummy row
    if ((r.getCell(col(header, 'meeting booked')).text || '').toLowerCase() === 'yes') continue;

    const topic = r.getCell(col(header, 'meeting name')).text;
    const dateChosen = r.getCell(col(header, 'date chosen')).text;
    const dateObj = dayjs.tz(dateChosen, 'Asia/Beirut');
    if (!dateObj.isValid()) {
      console.warn(`Skipping row ${i}: Invalid or missing date:`, dateChosen);
      continue;
    }
    const startISO = dateObj.format('YYYY-MM-DDTHH:mm:ss');
    const durationMin = +r.getCell(col(header, 'length of meeting(minutes)')).value || 60;
    const agenda = r.getCell(col(header, 'meeting name for invite')).text;
    const zoomAccount = r.getCell(col(header, 'Zoom Account to book from')).text || process.env.OUTLOOK_HOST;

    const participants = parseEmails(
      [
        r.getCell(col(header,'team invitees')).text,
        r.getCell(col(header,'added invitee 1')).text,
        r.getCell(col(header,'added invitee 2')).text,
        r.getCell(col(header,'added invitee 3')).text,
        r.getCell(col(header,'added invitee 4')).text,
        r.getCell(col(header,'email from form')).text
      ].join(',')
    );

    const zoom = await zCreate(token, topic, startISO, durationMin, agenda, zoomAccount);
    r.getCell(header['joinurl']).value = zoom.join_url;
    r.getCell(col(header, 'meeting booked')).value = 'yes';
    r.commit();

    // Always use OUTLOOK_HOST from .env for calendar invites
    await sendInvite(graph, {
      Topic: topic,
      StartUTC: startISO,
      DurationMin: durationMin,
      Agenda: agenda,
      Participants: participants.join(', '),
      TeamName: r.getCell(col(header, 'Team Name')).text,
      Team: r.getCell(col(header, 'team invitees')).text
    }, zoom.join_url, participants, process.env.OUTLOOK_HOST);
    console.log(`✔ ${topic}`);
  }

  await wb.xlsx.writeFile(file);
})();

export async function run(opts = 'Doodle Bookings.xlsx') {
  if (typeof opts === 'string') {
    // Legacy: file path mode
    const filePath = opts;
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(filePath);
    const ws    = wb.getWorksheet('Bookings');

    const header = buildHeaderMap(ws.getRow(1));
    const token  = await zToken();
    const graph  = await graphClient();

    if (!header['joinurl']) {
      header['joinurl'] = ws.columnCount + 1;
      ws.getRow(1).getCell(header['joinurl']).value = 'JoinURL';
    }

    for (let i = 2; i <= ws.rowCount; i++) {
      const r = ws.getRow(i);
      if (i === 2) continue; // skip Doodle dummy row
      if ((r.getCell(col(header, 'meeting booked')).text || '').toLowerCase() === 'yes') continue;

      const topic = r.getCell(col(header, 'meeting name')).text;
      const dateChosen = r.getCell(col(header, 'date chosen')).text;
      const dateObj = dayjs.tz(dateChosen, 'Asia/Beirut');
      if (!dateObj.isValid()) {
        console.warn(`Skipping row ${i}: Invalid or missing date:`, dateChosen);
        continue;
      }
      const startISO = dateObj.format('YYYY-MM-DDTHH:mm:ss');
      const durationMin = +r.getCell(col(header, 'length of meeting(minutes)')).value || 60;
      const agenda = r.getCell(col(header, 'meeting name for invite')).text;
      const zoomAccount = r.getCell(col(header, 'Zoom Account to book from')).text || process.env.OUTLOOK_HOST;

      const participants = parseEmails(
        [
          r.getCell(col(header,'team invitees')).text,
          r.getCell(col(header,'added invitee 1')).text,
          r.getCell(col(header,'added invitee 2')).text,
          r.getCell(col(header,'added invitee 3')).text,
          r.getCell(col(header,'added invitee 4')).text,
          r.getCell(col(header,'email from form')).text
        ].join(',')
      );

      const meetingType = r.getCell(col(header, 'meeting type')).text?.toLowerCase() || 'zoom';
      const location = r.getCell(col(header, 'location')).text || '';

      const zoom = await zCreate(token, topic, startISO, durationMin, agenda, zoomAccount);
      r.getCell(header['joinurl']).value = zoom.join_url;
      r.getCell(col(header, 'meeting booked')).value = 'yes';
      r.commit();

      // Pass global meetingType and location to sendInvite
      await sendInvite({
        team: r.getCell(col(header, 'team')).text,
        topic,
        dateChosen,
        durationMin,
        participants,
        zoomAccount,
        meetingType, // from global selection
        location    // from global selection
      });
      console.log('✔', topic);
    }

    await wb.xlsx.writeFile(filePath);
    return;
  }
  // New: object mode (from web form)
  const { team, teamName, topic, dateChosen, durationMin, participants, zoomAccount, meetingType, location } = opts;
  const startISO = dayjs.tz(dateChosen, 'Asia/Beirut').format('YYYY-MM-DDTHH:mm:ss');
  const token = await zToken();
  const graph = await graphClient();
  let joinUrl = '';
  if (meetingType !== 'in-person') {
    // Only create Zoom meeting if not in-person
    const zoom = await zCreate(token, topic, startISO, durationMin, topic, zoomAccount);
    joinUrl = zoom.join_url;
  }
  // Always send Outlook invite, with location set appropriately
  await sendInvite(
    graph,
    {
      Topic: topic,
      StartUTC: startISO,
      DurationMin: durationMin,
      TeamName: teamName, // Always use the correct team name
      Location: location || (meetingType === 'in-person' ? 'In person' : 'Zoom'),
    },
    joinUrl,
    participants.split(',').map(e => e.trim()).filter(Boolean),
    process.env.OUTLOOK_HOST
  );
}

/* still allow CLI use */
if (import.meta.url === `file://${process.argv[1]}`) {
  run(process.argv[2]);
}