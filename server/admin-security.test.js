const express = require('express');
const { mkdtempSync, mkdirSync, writeFileSync, rmSync } = require('fs');
const { tmpdir } = require('os');
const { join } = require('path');
const { createApp } = require('./app');

let server, root, origin;
beforeAll(async () => {
  root = mkdtempSync(join(tmpdir(), 'admin-security-'));
  mkdirSync(join(root, 'word'));
  for (const file of ['admin.html', 'taskpane.html', 'commands.html', 'consent.html', 'word/taskpane.html']) {
    writeFileSync(join(root, file), '<!doctype html><title>fixture</title>');
  }
  const app = createApp({ origin: 'https://example.test', authority: 'https://login.partner.microsoftonline.cn' }, {});
  // Use the same ordering as production: app routes, then express.static.
  app.use(express.static(root));
  server = await new Promise(resolve => { const s = app.listen(0, '127.0.0.1', () => resolve(s)); });
  origin = `http://127.0.0.1:${server.address().port}`;
});
afterAll(async () => {
  server?.closeAllConnections();
  if (server) await new Promise(resolve => server.close(resolve));
  if (root) rmSync(root, { recursive: true, force: true });
});

test.each(['/admin/usage', '/admin/api', '/admin/settings', '/admin.html', '/%61dmin.html', '/admin%2ehtml', '/admin/usage/?x=1'])('protects GET and HEAD for %s', async path => {
  for (const method of ['GET', 'HEAD']) {
    const response = await fetch(`${origin}${path}`, { method });
    expect(response.status).toBe(200);
    const csp = response.headers.get('content-security-policy');
    expect(csp).toContain("script-src 'self'");
    expect(csp).toContain("frame-ancestors 'none'");
    expect(csp).toContain("frame-src 'none'");
    expect(csp).toContain("base-uri 'none'");
    expect(csp).toContain('https://login.chinacloudapi.cn');
    expect(csp).toContain('https://login.partner.microsoftonline.cn');
    expect(csp).toContain('https://login.microsoftonline.com');
    expect(csp).not.toMatch(/unsafe-inline|unsafe-eval|\*/);
    expect(response.headers.get('x-frame-options')).toBe('DENY');
    expect(response.headers.get('x-content-type-options')).toBe('nosniff');
    expect(response.headers.get('cache-control')).toBe('no-store');
    if (method === 'GET') expect(await response.text()).toContain('<title>fixture</title>');
  }
});

test.each(['/taskpane.html', '/commands.html', '/consent.html', '/word/taskpane.html'])('does not restrict Office framing for %s', async path => {
  const response = await fetch(`${origin}${path}`);
  expect(response.status).toBe(200);
  expect(response.headers.get('content-security-policy')).toBeNull();
  expect(response.headers.get('x-frame-options')).toBeNull();
});
