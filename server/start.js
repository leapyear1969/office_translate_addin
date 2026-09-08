const fs = require('fs');
const path = require('path');
const https = require('https');
const express = require('express');
const { readConfig } = require('./config');
const { createApp } = require('./app');

async function start() {
  const config = readConfig();
  const app = createApp(config);
  app.use(express.static(path.resolve(__dirname, '../dist')));
  app.get('/', (_req, res) => res.redirect('/taskpane.html'));
  if (process.env.NODE_ENV === 'production' && !process.env.SSL_CERT_PATH) {
    app.listen(config.port, '0.0.0.0', () => console.log(`Backend listening on 0.0.0.0:${config.port}; HTTPS reverse proxy required.`));
    return;
  }
  const certificate = process.env.SSL_CERT_PATH ? {
    cert: fs.readFileSync(process.env.SSL_CERT_PATH), key: fs.readFileSync(process.env.SSL_KEY_PATH),
  } : await require('office-addin-dev-certs').getHttpsServerOptions();
  https.createServer(certificate, app)
    .on('error', error => { console.error(error.code === 'EADDRINUSE' ? `Port ${config.port} is already in use. Stop the other local instance first.` : 'HTTPS server failed to start.'); process.exitCode = 1; })
    .listen(config.port, '0.0.0.0', () => console.log(`Office translation available at ${config.origin}`));
}
start().catch(() => { console.error('Server failed to start. Check port and TLS certificate configuration.'); process.exitCode = 1; });
