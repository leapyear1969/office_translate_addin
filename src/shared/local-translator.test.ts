/** @jest-environment jsdom */
import { tryLocalRequest, setupLocalModels } from './local-translator';
import { api } from './api';

const env = globalThis as any;
let session: { translate: jest.Mock; destroy: jest.Mock };
let detector: { detect: jest.Mock; destroy: jest.Mock };
beforeEach(() => {
  session = { translate: jest.fn(async () => '译文 <安全>'), destroy: jest.fn() };
  detector = { detect: jest.fn(async () => [{ detectedLanguage: 'en', confidence: 0.99 }]), destroy: jest.fn() };
  env.Translator = { availability: jest.fn(async () => 'available'), create: jest.fn(async () => session) };
  env.LanguageDetector = { availability: jest.fn(async () => 'available'), create: jest.fn(async () => detector) };
  env.fetch = jest.fn(async () => ({ ok: true, json: async () => ({ html: 'cloud' }) }));
});
afterEach(() => { delete env.Translator; delete env.LanguageDetector; delete env.fetch; document.body.innerHTML = ''; jest.useRealTimers(); });

test('local mail preserves markup, whitespace and exclusions without sending content to the backend', async () => {
  const html = '<p>Hello <b>world</b></p><img src="cid:photo"><a href="https://example.com">Link</a><style>hidden</style><span translate="no">Keep</span>';
  const result = await api<any>('/api/translate', 'token', { html, subject: 'Meeting', to: 'zh-Hans' });
  expect(result.html).toContain('<p>译文 &lt;安全&gt; <b>译文 &lt;安全&gt;</b></p>');
  expect(result.html).toContain('src="cid:photo"');
  expect(result.html).toContain('href="https://example.com"');
  expect(result.html).toContain('<span translate="no">Keep</span>');
  expect(result.subject).toBe('译文 <安全>');
  expect(session.translate.mock.calls.map(call => call[0])).toEqual(['Hello ', 'world', 'Link', 'Meeting']);
  expect(env.fetch).not.toHaveBeenCalled();
  expect(session.destroy).toHaveBeenCalled();
  expect(detector.destroy).toHaveBeenCalled();
});
test('preserves Word markers and paragraph ordering', async () => {
  const result = await tryLocalRequest('/api/translate/word', {
    paragraphs: ['<p><span id="r0">Hello</span><span id="r1"> world</span></p>', '<p><span id="r0">Again</span></p>'], to: 'zh-Hans',
  });
  expect(result).toEqual({ paragraphs: ['<p><span id="r0">译文 &lt;安全&gt;</span><span id="r1"> 译文 &lt;安全&gt;</span></p>', '<p><span id="r0">译文 &lt;安全&gt;</span></p>'] });
});
test.each(['missing', 'unavailable', 'downloadable', 'denied', 'uncertain', 'partial'])(
  'falls back atomically when local model is %s', async mode => {
    if (mode === 'missing') delete env.Translator;
    if (mode === 'unavailable' || mode === 'downloadable') env.Translator.availability.mockResolvedValue(mode);
    if (mode === 'denied') env.Translator.create.mockRejectedValue(new Error('Permissions policy'));
    if (mode === 'uncertain') detector.detect.mockResolvedValue([{ detectedLanguage: 'und', confidence: 0.1 }]);
    if (mode === 'partial') session.translate.mockResolvedValueOnce('first').mockRejectedValueOnce(new Error('failed'));
    const body = { html: '<p>Hello</p><p>World</p>', to: 'zh-Hans' };
    expect(await api('/api/translate', 'token', body)).toEqual({ html: 'cloud' });
    expect(env.fetch).toHaveBeenCalledTimes(1);
    expect(JSON.parse(env.fetch.mock.calls[0][1].body).html).toBe(body.html);
  });
test('detects language locally without translating', async () => {
  expect(await api('/api/detect', 'token', { html: '<p>Hello world</p><style>hidden</style>' })).toEqual({ language: 'en', score: 0.99 });
  expect(detector.detect.mock.calls[0][0]).toBe('Hello world');
  expect(env.Translator.create).not.toHaveBeenCalled();
  expect(env.fetch).not.toHaveBeenCalled();
});
test('splits long nodes without splitting surrogate pairs', async () => {
  session.translate.mockImplementation(async text => text);
  const text = 'a'.repeat(4499) + '😀hello';
  expect(await tryLocalRequest('/api/translate', { html: `<p>${text}</p>`, to: 'zh-Hans' })).toEqual({ html: `<p>${text}</p>` });
  expect(session.translate.mock.calls[0][0]).toHaveLength(4499);
  expect(session.translate.mock.calls[1][0]).toBe('😀hello');
});
test('times out a stuck local API and continues through cloud', async () => {
  jest.useFakeTimers();
  env.Translator.create.mockReturnValue(new Promise(() => {}));
  const pending = api('/api/translate', 'token', { html: '<p>Hello</p>', to: 'zh-Hans' });
  await jest.advanceTimersByTimeAsync(75001);
  expect(await pending).toEqual({ html: 'cloud' });
  expect(detector.destroy).toHaveBeenCalled();
});
test('local translation remains available after a cloud quota error', async () => {
  env.Translator.availability.mockResolvedValueOnce('unavailable');
  env.fetch.mockResolvedValue({ ok: false, json: async () => ({ code: 'cloud_quota_exhausted', error: '本月云翻译额度已用完' }) });
  const body = { html: '<p>Hello</p>', to: 'zh-Hans' };
  await expect(api('/api/translate', 'token', body)).rejects.toMatchObject({ code: 'cloud_quota_exhausted' });
  expect(await api('/api/translate', 'token', body)).toEqual({ html: '<p>译文 &lt;安全&gt;</p>' });
  expect(env.fetch).toHaveBeenCalledTimes(1);
});
test('explicit preparation starts both model downloads on click', async () => {
  document.body.innerHTML = '<select id="local-source"></select><select id="local-target"></select><button id="prepare-local-models"></button><p id="local-model-status"></p>';
  setupLocalModels();
  document.getElementById('prepare-local-models')!.click();
  expect(env.LanguageDetector.create).toHaveBeenCalledTimes(1);
  expect(env.Translator.create).toHaveBeenCalledWith(expect.objectContaining({ sourceLanguage: 'en', targetLanguage: 'zh-Hans' }));
  for (let i = 0; i < 10; i++) await Promise.resolve();
  expect(document.getElementById('local-model-status')!.textContent).toContain('已就绪');
});

test('model preparation shows independent real progress and waits for initialization', async () => {
  document.body.innerHTML = '<select id="local-source"></select><select id="local-target"></select><button id="prepare-local-models"></button><p id="local-model-status"></p>';
  let report!: (event: { loaded: number; total?: number }) => void;
  let finish!: (value: typeof session) => void;
  env.Translator.create.mockImplementation((options: any) => {
    options.monitor({ addEventListener: (type: string, listener: typeof report) => { expect(type).toBe('downloadprogress'); report = listener; } });
    return new Promise(resolve => { finish = resolve; });
  });
  setupLocalModels();
  const button = document.getElementById('prepare-local-models') as HTMLButtonElement;
  button.click();
  const bars = document.querySelectorAll('progress');
  expect(bars).toHaveLength(2);
  expect(bars[1].hasAttribute('value')).toBe(false);
  for (let i = 0; i < 10; i++) await Promise.resolve();
  expect(bars[0].value).toBe(1);
  report({ loaded: 25, total: 100 });
  expect(bars[1].value).toBe(0.25);
  expect(bars[1].parentElement!.textContent).toContain('25%');
  report({ loaded: 1, total: 1 });
  expect(bars[1].parentElement!.textContent).toContain('正在初始化');
  expect(button.disabled).toBe(true);
  finish(session);
  for (let i = 0; i < 10; i++) await Promise.resolve();
  expect(bars[1].parentElement!.textContent).toContain('已就绪');
  expect(button.disabled).toBe(false);
  expect(session.destroy).toHaveBeenCalledTimes(1);
  report({ loaded: 0.5 });
  expect(bars[1].parentElement!.textContent).toContain('已就绪');
});

test('failed preparation stops its progress indicator and resets on retry', async () => {
  document.body.innerHTML = '<select id="local-source"></select><select id="local-target"></select><button id="prepare-local-models"></button><p id="local-model-status"></p>';
  env.Translator.create.mockRejectedValueOnce(new Error('network'));
  setupLocalModels();
  const button = document.getElementById('prepare-local-models') as HTMLButtonElement;
  button.click();
  for (let i = 0; i < 10; i++) await Promise.resolve();
  expect(document.querySelectorAll('progress')[1].hidden).toBe(true);
  expect(document.getElementById('local-model-status')!.textContent).toContain('失败');
  expect(button.disabled).toBe(false);
  button.click();
  expect(document.querySelectorAll('progress')).toHaveLength(2);
  for (let i = 0; i < 10; i++) await Promise.resolve();
  expect(document.querySelectorAll('progress')[1].hidden).toBe(false);
  expect(document.getElementById('local-model-status')!.textContent).toContain('已就绪');
});

test('an immediate failure cancels a stuck peer and preserves the actual error', async () => {
  document.body.innerHTML = '<select id="local-source"></select><select id="local-target"></select><button id="prepare-local-models"></button><p id="local-model-status"></p>';
  let finish!: (value: typeof detector) => void;
  env.LanguageDetector.create.mockReturnValue(new Promise(resolve => { finish = resolve; }));
  env.Translator.create.mockRejectedValue(new DOMException('Permissions policy blocked translator', 'NotAllowedError'));
  setupLocalModels();
  const button = document.getElementById('prepare-local-models') as HTMLButtonElement;
  button.click();
  for (let i = 0; i < 20; i++) await Promise.resolve();
  expect(button.disabled).toBe(false);
  expect(document.getElementById('local-model-status')!.textContent).toContain('未授予本地模型权限');
  expect(document.getElementById('local-model-status')!.textContent).toContain('Permissions policy blocked translator');
  expect(document.querySelectorAll('progress')[0].hidden).toBe(true);
  expect(document.querySelectorAll('progress')[0].parentElement!.textContent).toContain('已停止');
  finish(detector);
  for (let i = 0; i < 10; i++) await Promise.resolve();
  expect(detector.destroy).toHaveBeenCalledTimes(1);
  expect(document.querySelectorAll('progress')[0].parentElement!.textContent).toContain('已停止');
});

test('preparation timeout restores retry even if native APIs ignore abort', async () => {
  jest.useFakeTimers();
  document.body.innerHTML = '<select id="local-source"></select><select id="local-target"></select><button id="prepare-local-models"></button><p id="local-model-status"></p>';
  env.LanguageDetector.create.mockReturnValue(new Promise(() => {}));
  env.Translator.create.mockReturnValue(new Promise(() => {}));
  setupLocalModels();
  const button = document.getElementById('prepare-local-models') as HTMLButtonElement;
  button.click();
  await jest.advanceTimersByTimeAsync(300001);
  expect(button.disabled).toBe(false);
  expect(document.getElementById('local-model-status')!.textContent).toContain('超时');
  expect(Array.from(document.querySelectorAll('progress')).every(bar => bar.hidden)).toBe(true);
});
