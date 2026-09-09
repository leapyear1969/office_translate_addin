const { translateHtml, detectHtml, createTranslator } = require('./translator');

describe('Word paragraph translation', () => {
  const originalFetch = global.fetch;
  afterEach(() => { global.fetch = originalFetch; });
  const service = () => createTranslator({ translatorKey: 'test', translatorEndpoint: 'https://translator.example/' });
  test('sends entire mixed-format paragraphs in HTML mode and batches within limits', async () => {
    const paragraph = '<p><span id="r0">Hello </span><span id="r1">world</span></p>';
    const paragraphs = Array(102).fill(paragraph).concat(['<p><span id="r0">' + 'x'.repeat(39000) + '</span></p>', '<p><span id="r0">' + 'y'.repeat(39000) + '</span></p>']);
    global.fetch = jest.fn(async (_url, options) => ({ ok: true,
      json: async () => JSON.parse(options.body).map(({ Text }) => ({ translations: [{ text: Text }] })),
    }));
    expect(await service().translateWord(paragraphs, 'zh-Hans')).toEqual(paragraphs);
    for (const [url, options] of global.fetch.mock.calls) {
      expect(url.searchParams.get('textType')).toBe('html');
      expect(url.searchParams.get('to')).toBe('zh-Hans');
      const body = JSON.parse(options.body);
      expect(body.length).toBeLessThanOrEqual(100);
      expect(body.reduce((n, item) => n + item.Text.length, 0)).toBeLessThanOrEqual(45000);
    }
    expect(JSON.parse(global.fetch.mock.calls[0][1].body)[0].Text).toBe(paragraph);
  });
  test('rejects incomplete provider output rather than losing paragraphs', async () => {
    global.fetch = jest.fn(async () => ({ ok: true, json: async () => [] }));
    await expect(service().translateWord(['<p>Hello</p>'], 'en')).rejects.toThrow('不完整');
  });
  test.each([
    '<p><span id="r0">整段译文</span><span id="r1"></span></p>',
    '<p><span id="r1">链接</span><span id="r0">正文</span></p>',
    '<p><span id="r0">正文</span>未标记文字<span id="r1">链接</span></p>',
    '<p><span id="r0">正文</span></p>',
  ])('retries broken markers as plain fragments and keeps valid paragraphs: %s', broken => {
    const valid = '<p><span id="r0">Valid</span></p>';
    const source = '<p><span id="r0">Body </span><span id="r1">Link</span></p>';
    global.fetch = jest.fn(async (url, options) => ({ ok: true, json: async () =>
      url.searchParams.get('textType') === 'html'
        ? [valid, broken].map(text => ({ translations: [{ text }] }))
        : JSON.parse(options.body).map(({ Text }) => ({ translations: [{ text: Text.trim() === 'Body' ? '正文' : '链接 <安全>' }] })),
    }));
    return service().translateWord([valid, source], 'zh-Hans').then(result => {
      expect(result).toEqual([valid, '<p><span id="r0">正文 </span><span id="r1">链接 &lt;安全&gt;</span></p>']);
      expect(global.fetch).toHaveBeenCalledTimes(2);
      expect(JSON.parse(global.fetch.mock.calls[1][1].body)).toEqual([{ Text: 'Body ' }, { Text: 'Link' }]);
      expect(global.fetch.mock.calls[1][0].searchParams.has('textType')).toBe(false);
    });
  });
  test('rejects empty fallback translations instead of returning a partially untranslated document', async () => {
    global.fetch = jest.fn(async url => ({ ok: true, json: async () => [{ translations: [{
      text: url.searchParams.get('textType') === 'html' ? '<p><span id="r0"></span></p>' : ' ',
    }] }] }));
    await expect(service().translateWord(['<p><span id="r0">Original</span></p>'], 'zh-Hans')).rejects.toThrow('格式标记');
    expect(global.fetch).toHaveBeenCalledTimes(2);
  });
});

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
