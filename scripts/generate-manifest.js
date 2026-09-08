const fs = require('fs');
const path = require('path');
const { readConfig } = require('../server/config');
const { selectedHosts } = require('./hosts');
function validateSsoConfig(config) {
  let origin;
  let resource;
  try {
    origin = new URL(config.origin);
    resource = new URL(config.resource);
  } catch {
    throw new Error('APP_BASE_URL 和 SSO_RESOURCE 必须是有效 URL；SSO_RESOURCE 应与 Entra Application ID URI 一致。');
  }
  // Convert to HTTPS before comparing so default ports normalize equally.
  const resourceOrigin = new URL('https://' + resource.host).origin;
  if (origin.protocol !== 'https:' || origin.username || origin.password ||
      origin.pathname !== '/' || origin.search || origin.hash ||
      resource.protocol !== 'api:' || resource.username || resource.password ||
      resource.search || resource.hash || resource.pathname.length <= 1 ||
      origin.origin !== resourceOrigin) {
    throw new Error('Office SSO 13004：APP_BASE_URL 必须是 HTTPS 源地址，且域名和端口必须与 SSO_RESOURCE（api://域名:端口/应用标识）一致。请同时核对 Entra Application ID URI 后重新生成清单。');
  }
}
function generate(host = 'all', config = readConfig(), directory = path.resolve(__dirname, '..')) {
  const targets = selectedHosts(host);
  validateSsoConfig(config);
  for (const target of targets) {
    const manifest = require('./manifests/' + target)(config);
    fs.writeFileSync(path.join(directory, 'manifest.' + target + '.xml'), manifest);
    // Preserve the existing Outlook sideload path and identity.
    if (target === 'outlook') fs.writeFileSync(path.join(directory, 'manifest.xml'), manifest);
    console.log('Generated manifest.' + target + '.xml');
  }
}
if (require.main === module) generate(process.argv[2]);
module.exports = { generate };
