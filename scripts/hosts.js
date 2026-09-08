function selectedHosts(host = 'all') {
  if (host === 'all') return ['outlook', 'word'];
  if (host === 'outlook' || host === 'word') return [host];
  throw new Error('Unknown add-in host: ' + host + '. Expected outlook, word or all.');
}
module.exports = { selectedHosts };
