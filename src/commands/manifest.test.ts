import { readFileSync } from 'fs';
import { join } from 'path';
import { load } from 'cheerio';

const root = join(__dirname, '../..');
function generatedManifest(): string {
  return require('../../scripts/manifests/outlook')({ origin: 'https://example.com', clientId: 'id', resource: 'resource' });
}

test.each(['checked-in', 'generated'])('%s manifest exposes only translation and options menu items', source => {
  const xml = source === 'generated' ? generatedManifest() : readFileSync(join(root, 'manifest.outlook.xml'), 'utf8');
  const $ = load(xml, { xmlMode: true });
  const menus = $('Control[id="Translation.Menu"]');
  expect(menus.length).toBe(2);
  menus.each((_index, menu) => {
    expect($(menu).find('Item').map((_i, item) => $(item).attr('id')).get()).toEqual(['Translation.Message', 'Translation.Options']);
    expect($(menu).find('Item[id="Translation.Options"] SourceLocation').attr('resid')).toBe('Taskpane.Url');
  });
  expect(xml).not.toContain('Original.Url');
});
