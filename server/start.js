const fs = require('fs');
const path = require('path');
const https = require('https');
const express = require('express');
const { readConfig } = require('./config');
const { createApp } = require('./app');
let analytics;
async function failStart(message) {
  console.error(message);
  process.exitCode = 1;
  await analytics?.close().catch(() => {});
}

async function start() {
  const config = readConfig();
  const app = createApp(config);
  analytics = app.locals.analytics;
  let server;
  let stopping = false;
  const stop = () => {
    if (stopping) return;
    stopping = true;
    // Stop accepting work, then drain request completions before closing the writer.
    const deadline = setTimeout(() => { console.error('Shutdown timed out; statistics may be incomplete.'); process.exit(1); }, 120000);
    server?.close(async () => {
      try { await app.locals.analytics.close(); clearTimeout(deadline); process.exit(0); }
      catch { clearTimeout(deadline); process.exit(1); }
    });
  };
  process.on('SIGINT', stop); process.on('SIGTERM', stop);
  app.use(express.static(path.resolve(__dirname, '../dist')));
  app.get('/', (_req, res) => res.redirect('/taskpane.html'));
  if (process.env.NODE_ENV === 'production' && !process.env.SSL_CERT_PATH) {
    server = app.listen(config.port, '0.0.0.0', () => console.log(`Backend listening on 0.0.0.0:${config.port}; HTTPS reverse proxy required.`));
    server.on('error', () => { void failStart('Server failed to start. Check port configuration.'); });
    return;
  }
  const certificate = process.env.SSL_CERT_PATH ? {
    cert: fs.readFileSync(process.env.SSL_CERT_PATH), key: fs.readFileSync(process.env.SSL_KEY_PATH),
  } : await require('office-addin-dev-certs').getHttpsServerOptions();
  server = https.createServer(certificate, app)
    .on('error', error => { void failStart(error.code === 'EADDRINUSE' ? `Port ${config.port} is already in use. Stop the other local instance first.` : 'HTTPS server failed to start.'); })
    .listen(config.port, '0.0.0.0', () => console.log(`Office translation available at ${config.origin}`));
}
start().catch(() => failStart('Server failed to start. Check port and TLS certificate configuration.'));
