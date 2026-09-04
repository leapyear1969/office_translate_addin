const { translateHtml, detectHtml } = require('./translator');

describe('mail HTML translation', () => {
  test('preserves spaces around inline formatting even if Translator trims them', async () => {
    const result = await translateHtml('<p>Hello <b>world</b> !</p>', 'fr', async texts => texts.map(t => t.trim().toUpperCase()));
    expect(result).toBe('<p>HELLO <b>WORLD</b> !</p>');
  });
  test('preserves formatting, images and links and escapes translated text', async () => {
    const translate = jest.fn(async texts => texts.map(() => '译文 <安全>'));
    const html = '<table><tr><td style="color:red">Hello</td></tr></table><img src="cid:photo"><a href="https://example.com">Link</a><style>p{color:red}</style><span class="notranslate">Keep</span>';
    const result = await translateHtml(html, 'zh-Hans', translate);
    expect(translate.mock.calls.flatMap(c => c[0])).toEqual(['Hello', 'Link']);
    expect(result).toContain('style="color:red"');
    expect(result).toContain('src="cid:photo"');
    expect(result).toContain('href="https://example.com"');
    expect(result).toContain('译文 &lt;安全&gt;');
    expect(result).toContain('p{color:red}');
    expect(result).toContain('>Keep</span>');
  });
  test('batches long emails and preserves all characters', async () => {
    const translate = jest.fn(async texts => texts);
    const html = `<p>${'中😀'.repeat(24000)}</p>`;
    const result = await translateHtml(html, 'en', translate);
    expect(result).toContain('中😀'.repeat(24000));
    for (const [texts] of translate.mock.calls) {
      expect(texts.join('').length).toBeLessThanOrEqual(45000);
      expect(texts.length).toBeLessThanOrEqual(100);
      expect(texts.every(t => t.length <= 4500)).toBe(true);
      expect(texts.some(t => /[\uD800-\uDBFF]$/.test(t))).toBe(false);
    }
  });
  test('fails atomically on partial translator responses', async () => {
    await expect(translateHtml('<p>Hello</p><p>World</p>', 'en', async () => ['one']))
      .rejects.toThrow('翻译结果不完整');
  });
  test('does not call translator for empty visible content', async () => {
    const translate = jest.fn();
    await translateHtml('<img src="cid:x"><style>abc</style>', 'en', translate);
    expect(translate).not.toHaveBeenCalled();
  });
  test('detects visible text only and skips empty messages', async () => {
    const detect = jest.fn(async text => ({ language: 'en', score: 0.99 }));
    expect(await detectHtml('<style>hidden</style><p>Hello</p>', detect)).toEqual({ language: 'en', score: 0.99 });
    expect(detect).toHaveBeenCalledWith('Hello');
    expect(await detectHtml('<img src="cid:x">', detect)).toBeNull();
  });
});
