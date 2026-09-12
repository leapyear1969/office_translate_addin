const { readConfig } = require('./config');
const { cutoff } = require('./analytics-store');
const { filters } = require('./admin');
const { createAnalytics } = require('./analytics');

test('retention environment defaults and invalid values', () => {
  expect(readConfig({}).analyticsRetentionMonths).toBe(12);
  expect(readConfig({ANALYTICS_RETENTION_MONTHS:''}).analyticsRetentionMonths).toBe(12);
  expect(readConfig({ANALYTICS_RETENTION_MONTHS:'24'}).analyticsRetentionMonths).toBe(24);
  for (const value of ['0', '-1', '1.5', 'abc', 'Infinity', ' ', '1201']) {
    expect(() => readConfig({ANALYTICS_RETENTION_MONTHS:value})).toThrow('ANALYTICS_RETENTION_MONTHS');
  }
});

test('custom retention handles month ends, leap years and Shanghai date boundaries', () => {
  expect(cutoff(Date.parse('2024-03-31T03:00:00Z'), 1)).toBe('2024-02-29');
  expect(cutoff(Date.parse('2025-03-31T03:00:00Z'), 1)).toBe('2025-02-28');
  expect(cutoff(Date.parse('2026-09-30T16:00:00Z'), 24)).toBe('2024-10-01');
});

test('filters use custom retention and clamp default dates for short months', () => {
  const now = Date.parse('2025-03-01T03:00:00Z');
  expect(filters({}, ['*'], now, 1).from).toBe('2025-02-01');
  expect(() => filters({from:'2025-01-31'}, ['*'], now, 1)).toThrow('保留期');
  expect(filters({from:'2023-03-01'}, ['*'], now, 24).from).toBe('2023-03-01');
});

test('analytics passes retention to the store and performs startup cleanup', async () => {
  const store = {meta:async()=>({}), cleanup:jest.fn(), close:async()=>{}};
  const factory = jest.fn(async()=>store);
  const analytics = createAnalytics('test', factory, readConfig({ANALYTICS_RETENTION_MONTHS:'6'}).analyticsRetentionMonths);
  try {
    await analytics.flush();
    expect(factory).toHaveBeenCalledWith('test', undefined, undefined, 6);
    expect(store.cleanup).toHaveBeenCalledTimes(1);
  } finally { await analytics.close(); }
});
