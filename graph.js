import { Client } from '@microsoft/microsoft-graph-client';
import { ClientSecretCredential } from '@azure/identity';
import dayjs from './dayjs-setup.js';
import 'dotenv/config';

const cred = new ClientSecretCredential(
  process.env.AAD_TENANT_ID,
  process.env.AAD_CLIENT_ID,
  process.env.AAD_CLIENT_SECRET
);

export async function graphClient() {
  const tok = await cred.getToken('https://graph.microsoft.com/.default');
  return Client.init({
    defaultVersion: 'v1.0',
    authProvider: d => d(null, tok.token)
  });
}

export async function sendInvite(graph, row, joinUrl, attendees, outlookHost) {
  // Parse as Beirut local time
  const start = dayjs.tz(row.StartUTC, 'Asia/Beirut');
  const isZoom = row.Location === 'Zoom' || row.meetingType === 'zoom';
  // Use the team name from the 'Team Name' column only
  const teamName = row.TeamName || '';
  let locationLine = '';
  if (isZoom && joinUrl) {
    locationLine = `<br>Zoom Link: <a href="${joinUrl}">${joinUrl}</a><br>`;
  } else if (!isZoom && row.Location) {
    locationLine = `<br>Location: ${row.Location}<br>`;
  }
  const inviteBody = `Dear ${teamName},<br>You are invited to a meeting for your team.<br>Meeting Day & Time: ${start.format('YYYY-MM-DD HH:mm')}${locationLine}<br>Best regards,<br>Berytech Team`;
  const eventBody = {
    subject: row.Topic,
    body: {
      contentType: 'HTML',
      content: inviteBody
    },
    start: { dateTime: start.format('YYYY-MM-DDTHH:mm:ss'), timeZone: 'Asia/Beirut' },
    end:   { dateTime: start.add(+row.DurationMin, 'minute').format('YYYY-MM-DDTHH:mm:ss'), timeZone: 'Asia/Beirut' },
    location: { displayName: row.Location || 'Zoom' },
    attendees: attendees.map(a => ({ type: 'required', emailAddress: { address: a } }))
  };
  await graph.api(`/users/${outlookHost}/events`).post(eventBody);
}