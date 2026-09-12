const DEFAULT_RETENTION_MONTHS = 12;

function parseRetentionMonths(value) {
  if (value === undefined || value === '') return DEFAULT_RETENTION_MONTHS;
  const months = Number(value);
  if ((typeof value === 'string' && !value.trim()) || !Number.isSafeInteger(months) || months < 0 || months > 1200) {
    throw new Error('ANALYTICS_RETENTION_MONTHS must be an integer between 0 and 1200 (0 keeps data forever)');
  }
  return months;
}

module.exports = { DEFAULT_RETENTION_MONTHS, parseRetentionMonths };
