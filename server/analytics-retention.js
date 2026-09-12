const DEFAULT_RETENTION_MONTHS = 12;

function parseRetentionMonths(value) {
  if (value === undefined || value === '') return DEFAULT_RETENTION_MONTHS;
  const months = Number(value);
  if (!Number.isSafeInteger(months) || months < 1 || months > 1200) {
    throw new Error('ANALYTICS_RETENTION_MONTHS must be an integer between 1 and 1200');
  }
  return months;
}

module.exports = { DEFAULT_RETENTION_MONTHS, parseRetentionMonths };
