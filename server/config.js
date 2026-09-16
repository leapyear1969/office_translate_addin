const path = require('path');
const { parseAdmins } = require('./admin');
const { parseRetentionMonths } = require('./analytics-retention');
require('dotenv').config({ path: path.resolve(__dirname, '../.env') });
function readConfig(env = process.env) {
  return {
    origin: (env.APP_BASE_URL || 'https://localhost:3000').replace(/\/$/, ''),
    port: Number(env.PORT || 3000),
    analyticsDatabaseUrl: env.ANALYTICS_DATABASE_URL || '',
    analyticsRetentionMonths: parseRetentionMonths(env.ANALYTICS_RETENTION_MONTHS),
    analyticsAdmins: parseAdmins(env.ANALYTICS_ADMINS || '[]'),
    adminClientId: env.ADMIN_CLIENT_ID || env.CLIENT_ID || '',
    adminTenant: env.ADMIN_LOGIN_TENANT || 'organizations',
    tenantId: env.TENANT_ID || '', clientId: env.CLIENT_ID || '',
    clientSecret: env.CLIENT_SECRET || '', resource: env.SSO_RESOURCE || '',
    authority: (env.AUTHORITY || 'https://login.partner.microsoftonline.cn').replace(/\/$/, ''),
    graphBase: (env.GRAPH_BASE || 'https://microsoftgraph.chinacloudapi.cn').replace(/\/$/, ''),
    translatorEndpoint: env.TRANSLATOR_ENDPOINT || 'https://api.translator.azure.cn/',
    localTranslatorEndpoint: env.LOCAL_TRANSLATOR_ENDPOINT === '' ? '' : (env.LOCAL_TRANSLATOR_ENDPOINT || 'http://192.168.3.101:30261'),
    localTranslatorFallback: env.LOCAL_TRANSLATOR_FALLBACK !== 'false',
    translatorKey: env.TRANSLATOR_KEY || '', translatorRegion: env.TRANSLATOR_REGION || '',
  };
}
module.exports = { readConfig };
