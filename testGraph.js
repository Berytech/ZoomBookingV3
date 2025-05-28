// testGraph.js
import dotenv from 'dotenv';
import { Client } from '@microsoft/microsoft-graph-client';
import { ClientSecretCredential } from '@azure/identity';

dotenv.config();

const cred = new ClientSecretCredential(
  process.env.AAD_TENANT_ID,
  process.env.AAD_CLIENT_ID,
  process.env.AAD_CLIENT_SECRET
);

const token = await cred.getToken('https://graph.microsoft.com/.default');
console.log('Token OK, expires', token.expiresOnTimestamp);

const graph = Client.init({
  defaultVersion: 'v1.0',
  authProvider: done => done(null, token.token)
});

const me = await graph.api('/organization?$select=displayName').get();
console.log('Tenant:', me.value[0].displayName);