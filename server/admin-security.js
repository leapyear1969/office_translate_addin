const { posix } = require('path');

function adminPageSecurity(config) {
  const authority = new URL(config.authority || 'https://login.partner.microsoftonline.cn').origin;
  // MSAL normalizes cloud aliases and uses the global instance discovery endpoint.
  const connections = [...new Set([authority, 'https://login.microsoftonline.com',
    ...(['login.partner.microsoftonline.cn', 'login.chinacloudapi.cn'].includes(new URL(authority).hostname)
      ? ['https://login.chinacloudapi.cn', 'https://login.partner.microsoftonline.cn'] : [])])];
  const policy = ["default-src 'none'", "script-src 'self'", "style-src 'self'",
    "img-src 'self'", `connect-src 'self' ${connections.join(' ')}`,
    "frame-src 'none'", "frame-ancestors 'none'", "object-src 'none'", "base-uri 'none'", "form-action 'none'"].join('; ');
  return (req, res, next) => {
    let pathname;
    try { pathname = posix.normalize(decodeURIComponent(req.path).replace(/\\/g, '/')).toLowerCase(); }
    catch { return next(); }
    // Include direct/encoded static aliases without preventing Office from embedding task panes.
    if (pathname === '/admin.html' || /^\/admin(?:\/|$)/.test(pathname)) {
      res.set({ 'Content-Security-Policy': policy, 'X-Frame-Options': 'DENY',
        'X-Content-Type-Options': 'nosniff', 'Cache-Control': 'no-store', 'Referrer-Policy': 'no-referrer' });
    }
    next();
  };
}

module.exports = { adminPageSecurity };
