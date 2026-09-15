import { LANGUAGES } from './settings';
type Availability = 'available' | 'downloadable' | 'downloading' | 'unavailable';
interface Detector {
  detect(text: string, options: { signal: AbortSignal }): Promise<{ detectedLanguage: string; confidence: number }[]>;
  destroy(): void;
}
interface Session {
  translate(text: string, options: { signal: AbortSignal }): Promise<string>;
  destroy(): void;
}
interface Pair { sourceLanguage: string; targetLanguage: string }
interface Models {
  LanguageDetector?: { availability(): Promise<Availability>; create(options: { signal: AbortSignal }): Promise<Detector> };
  Translator?: { availability(pair: Pair): Promise<Availability>; create(pair: Pair & { signal: AbortSignal }): Promise<Session> };
}
const models = () => globalThis as unknown as Models;
const unavailable = () => new Error('本地翻译暂不可用');
function usable(state: Availability) {
  return state === 'available' || ((state === 'downloadable' || state === 'downloading')
    && typeof navigator !== 'undefined' && navigator.userActivation?.isActive);
}
function fragment(html: string) {
  // Inert template: neither original mail nor model output enters the live document.
  const template = document.createElement('template');
  template.innerHTML = html;
  const nodes: Text[] = [];
  const walk = (node: Node) => {
    if (node.nodeType === 1 && (node as Element).matches('script,style,head,noscript,title,template,.notranslate,[translate="no"]')) return;
    if (node.nodeType === 3 && node.textContent?.trim()) nodes.push(node as Text);
    node.childNodes.forEach(walk);
  };
  walk(template.content);
  return { template, nodes };
}

/** undefined requests cloud fallback. Never return partially translated content. */
export async function tryLocalRequest(path: string, body: unknown): Promise<unknown | undefined> {
  if (!['/api/detect', '/api/translate', '/api/translate/word'].includes(path)
    || !body || typeof body !== 'object' || typeof document === 'undefined') return undefined;
  const { html, paragraphs, subject, to } = body as { html?: string; paragraphs?: string[]; subject?: string; to?: string };
  const word = path === '/api/translate/word';
  const detecting = path === '/api/detect';
  const input = word ? paragraphs : [html];
  if (!Array.isArray(input) || !input.length || input.length > 10000
    || input.some(value => typeof value !== 'string' || !value || value.length > (word ? 40000 : 1000000))
    || input.join('').length > 1000000 || (!detecting && (typeof to !== 'string' || !/^[a-z]{2,3}(?:-[A-Za-z]{2,8})?$/.test(to)))
    || (subject !== undefined && (typeof subject !== 'string' || subject.length > 10000))) return undefined;
  const { LanguageDetector, Translator } = models();
  if (!LanguageDetector || (!detecting && !Translator)) return undefined;
  const controller = new AbortController();
  const sessions: Session[] = [];
  let detector: Detector | undefined;
  let expired = false;
  let timer: ReturnType<typeof setTimeout> | undefined;
  const work = async () => {
    if (!usable(await LanguageDetector.availability())) throw unavailable();
    detector = await LanguageDetector.create({ signal: controller.signal });
    if (expired) { detector.destroy(); throw unavailable(); }
    const fragments = (input as string[]).map(fragment);
    const detect = async (text: string) => {
      const result = (await detector!.detect(text.slice(0, 10000), { signal: controller.signal }))[0];
      if (!result || !result.detectedLanguage || result.detectedLanguage === 'und'
        || !Number.isFinite(result.confidence) || result.confidence < 0.5 || result.confidence > 1) throw unavailable();
      return { language: result.detectedLanguage, score: result.confidence };
    };
    if (detecting) {
      const text = fragments[0].nodes.map(node => node.data.trim()).join(' ');
      return text ? detect(text) : null;
    }
    const pairs = new Map<string, Session>();
    const translate = async (text: string) => {
      if (!text.trim() || !/\p{L}/u.test(text)) return text;
      const { language } = await detect(text);
      // Different scripts such as zh-Hans and zh-Hant must remain distinct.
      if (language.toLowerCase() === to!.toLowerCase()) return text;
      let session = pairs.get(language);
      if (!session) {
        const pair = { sourceLanguage: language, targetLanguage: to! };
        if (!usable(await Translator!.availability(pair))) throw unavailable();
        session = await Translator!.create({ ...pair, signal: controller.signal });
        if (expired) { session.destroy(); throw unavailable(); }
        sessions.push(session); pairs.set(language, session);
      }
      let result = '';
      for (let offset = 0; offset < text.length;) {
        if (controller.signal.aborted) throw unavailable();
        let end = Math.min(offset + 4500, text.length);
        if (end < text.length && /[\uD800-\uDBFF]/.test(text[end - 1]) && /[\uDC00-\uDFFF]/.test(text[end])) --end;
        const part = text.slice(offset, end);
        const value = await session.translate(part, { signal: controller.signal });
        if (typeof value !== 'string' || (part.trim() && !value.trim())) throw unavailable();
        result += part.trim() ? part.match(/^\s*/)![0] + value.trim() + part.match(/\s*$/)![0] : part;
        if (result.length > 1000000) throw unavailable();
        offset = end;
      }
      return result;
    };
    for (const item of fragments) for (const node of item.nodes) node.data = await translate(node.data);
    const output = fragments.map((item, index) => item.nodes.length ? item.template.innerHTML : input[index]);
    if (output.join('').length > 1000000) throw unavailable();
    if (word) return { paragraphs: output };
    return { html: output[0], ...(subject === undefined ? {} : { subject: await translate(subject) }) };
  };
  try {
    return await Promise.race([work(), new Promise<never>((_, reject) => {
      timer = setTimeout(() => { expired = true; controller.abort(); reject(unavailable()); }, 75000);
    })]);
  } catch { return undefined; }
  finally {
    expired = true; clearTimeout(timer); controller.abort();
    for (const session of [detector, ...sessions]) { try { session?.destroy(); } catch { /* Already destroyed. */ } }
  }
}

export function setupLocalModels(): void {
  const button = document.getElementById('prepare-local-models') as HTMLButtonElement | null;
  if (!button) return;
  for (const [id, selected] of [['local-source', 'en'], ['local-target', 'zh-Hans']]) {
    const select = document.getElementById(id) as HTMLSelectElement;
    select.replaceChildren(...Object.entries(LANGUAGES).map(([code, label]) => new Option(label, code)));
    select.value = selected;
  }
  button.onclick = async () => {
    const status = document.getElementById('local-model-status')!;
    const source = (document.getElementById('local-source') as HTMLInputElement).value;
    const target = (document.getElementById('local-target') as HTMLInputElement).value;
    const { LanguageDetector, Translator } = models();
    if (!LanguageDetector || !Translator) { status.textContent = '当前 Office 环境未开放本地翻译 API，将使用 Azure F0。'; return; }
    button.disabled = true;
    status.textContent = '正在准备本地模型，首次下载可能需要几分钟…';
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), 300000);
    try {
      // Start both downloads in the click handler while user activation is present.
      const results = await Promise.allSettled([
        LanguageDetector.create({ signal: controller.signal }),
        Translator.create({ sourceLanguage: source, targetLanguage: target, signal: controller.signal }),
      ]);
      results.forEach(result => { if (result.status === 'fulfilled') result.value.destroy(); });
      status.textContent = results.every(result => result.status === 'fulfilled')
        ? '该语言对的本地模型已就绪，请重新翻译。' : '模型准备失败，请检查语言对、网络或当前 Office 环境权限。';
    } catch { status.textContent = '模型准备失败，请重试。'; }
    finally { clearTimeout(timer); controller.abort(); button.disabled = false; }
  };
}
