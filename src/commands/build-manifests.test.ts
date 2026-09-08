import { mkdtempSync, readFileSync, readdirSync, rmSync } from 'fs';
import { tmpdir } from 'os';
import { join } from 'path';
import { load } from 'cheerio';

const { generate } = require('../../scripts/generate-manifest');
const config = { origin: 'https://example.com', clientId: 'shared-client', resource: 'api://example.com/shared&client' };
let directory: string;
beforeEach(() => { directory = mkdtempSync(join(tmpdir(), 'office-manifests-')); });
afterEach(() => { rmSync(directory, { recursive: true, force: true }); });

test('generates independent host identities with shared SSO and a compatible Outlook alias', () => {
  generate('all', config, directory);
  const outlook = readFileSync(join(directory, 'manifest.outlook.xml'), 'utf8');
  const word = readFileSync(join(directory, 'manifest.word.xml'), 'utf8');
  const mail = load(outlook, { xmlMode: true });
  const doc = load(word, { xmlMode: true });
  expect(readFileSync(join(directory, 'manifest.xml'), 'utf8')).toBe(outlook);
  expect(mail('OfficeApp').attr('xsi:type')).toBe('MailApp');
  expect(doc('OfficeApp').attr('xsi:type')).toBe('TaskPaneApp');
  expect(doc('OfficeApp > Id').text()).not.toBe(mail('OfficeApp > Id').text());
  expect(doc('OfficeApp > Hosts > Host').attr('Name')).toBe('Document');
  expect(doc('Permissions').text()).toBe('ReadWriteDocument');
  expect(doc('DefaultSettings SourceLocation').attr('DefaultValue')).toBe('https://example.com/word/taskpane.html');
  for (const $ of [mail, doc]) {
    expect($('WebApplicationInfo > Id').text()).toBe(config.clientId);
    expect($('WebApplicationInfo > Resource').text()).toBe(config.resource);
  }
  expect(word).not.toMatch(/Mailbox|MailHost|ReadWriteItem|SupportsPinning|translateMessage/);
});

test.each(['outlook', 'word'])('generating %s preserves the other host manifest', host => {
  generate('all', config, directory);
  const other = host === 'outlook' ? 'word' : 'outlook';
  const previous = readFileSync(join(directory, `manifest.${other}.xml`), 'utf8');
  const alias = readFileSync(join(directory, 'manifest.xml'), 'utf8');
  generate(host, { ...config, origin: 'https://new.example.com', resource: 'api://new.example.com/shared&client' }, directory);
  expect(readFileSync(join(directory, `manifest.${other}.xml`), 'utf8')).toBe(previous);
  expect(readFileSync(join(directory, `manifest.${host}.xml`), 'utf8')).toContain('https://new.example.com');
  if (host === 'word') expect(readFileSync(join(directory, 'manifest.xml'), 'utf8')).toBe(alias);
});

test('rejects an unknown host before writing files', () => {
  expect(() => generate('excel', config, directory)).toThrow('Unknown add-in host');
  expect(readdirSync(directory)).toEqual([]);
});

test.each([
  'api://other.example.com/shared-client',
  'api://example.com:3000/shared-client',
  'api://shared-client',
  'https://example.com/shared-client',
])('rejects invalid SSO resource %s without overwriting manifests', resource => {
  generate('all', config, directory);
  const before = readdirSync(directory).map(file => readFileSync(join(directory, file), 'utf8'));
  expect(() => generate('all', { ...config, resource }, directory)).toThrow('13004');
  expect(readdirSync(directory).map(file => readFileSync(join(directory, file), 'utf8'))).toEqual(before);
});

test('accepts matching local development ports', () => {
  expect(() => generate('word', { ...config, origin: 'https://localhost:3000', resource: 'api://localhost:3000/shared-client' }, directory)).not.toThrow();
});

test('Word context menu routes translation and settings into the same shared runtime', () => {
  generate('word', config, directory);
  const $ = load(readFileSync(join(directory, 'manifest.word.xml'), 'utf8'), { xmlMode: true });
  const controls = $('OfficeMenu[id="ContextMenuText"] > Control');
  expect(controls.length).toBe(2);
  expect(controls.eq(0).find('FunctionName').text()).toBe('translateSelectionChinese');
  expect(controls.eq(1).find('Action').attr('xsi:type')).toBe('ShowTaskpane');
  expect($('Runtime').attr('resid')).toBe($('FunctionFile').attr('resid'));
  expect($('[id="Open.Label"]').attr('DefaultValue')).toBe('翻译选项');
});
