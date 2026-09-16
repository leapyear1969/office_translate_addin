const { detectAll } = require('tinyld');
const { recordRequest } = require('./analytics');

function detectLocal(text) {
  const best = detectAll(text)[0];
  return best ? { language: best.lang, score: best.accuracy } : null;
}

function createLocalTranslator(config) {
  let cachedLanguages, expires = 0;
  const endpoint = config.localTranslatorEndpoint.replace(/\/$/, '') + '/';
  const unavailable = message => Object.assign(new Error(message), { status: 503 });
  async function languages() {
    if (cachedLanguages && Date.now() < expires) return cachedLanguages;
    const response = await fetch(new URL('languages', endpoint), { signal: AbortSignal.timeout(3000) });
    if (!response.ok) throw unavailable('本地翻译服务无法读取语言列表。');
    const result = await response.json();
    if (!Array.isArray(result) || !result.every(item => typeof item.code === 'string'
      && Array.isArray(item.targets) && item.targets.every(target => typeof target === 'string'))) {
      throw unavailable('本地翻译服务语言列表格式无效。');
    }
    cachedLanguages = result;
    expires = Date.now() + 60000;
    return result;
  }
  async function translate(texts, to, deadline = Date.now() + 25000) {
    const target = to === 'zh-Hans' ? 'zh' : to;
    const routes = await languages();
    const context = detectLocal(texts.join(' '));
    const values = [];
    for (const text of texts) {
      if (!text.trim() || !/\p{L}/u.test(text)) { values.push(text); continue; }
      let detected = detectLocal(text);
      if ((!detected || detected.score < 0.5) && context?.score >= 0.5) detected = context;
      if (!detected || detected.score < 0.5) throw unavailable('无法可靠识别原文语言，本地翻译未执行。');
      const source = detected.language;
      if (source === target) { values.push(text); continue; }
      if (!routes.some(route => route.code === source && route.targets.includes(target))) {
        throw unavailable('本地翻译服务尚未安装该语言方向的模型。');
      }
      let translated = '';
      for (let offset = 0; offset < text.length;) {
        let end = Math.min(offset + 4500, text.length);
        if (/[\uD800-\uDBFF]/.test(text[end - 1])) end--;
        const part = text.slice(offset, end);
        offset = end;
        const remaining = deadline - Date.now();
        if (remaining <= 0) throw unavailable('本地翻译超时，请重试。');
        const started = Date.now(); let success = false;
        try {
          const response = await fetch(new URL('translate', endpoint), {
            method: 'POST', headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ text: part, source, target }), signal: AbortSignal.timeout(remaining),
          });
          if (!response.ok) throw unavailable(`本地翻译服务暂时不可用（HTTP ${response.status}）。`);
          const result = await response.json();
          if (typeof result?.translated_text !== 'string' || !result.translated_text.trim()
            || result.source !== source || result.target !== target) {
            throw unavailable('本地翻译服务响应格式无效或译文为空。');
          }
          translated += part.match(/^\s*/)[0] + result.translated_text.trim() + part.match(/\s*$/)[0];
          success = true;
        } finally { recordRequest('upstream', 'translate', started, success, part.length); }
      }
      values.push(translated);
    }
    return values;
  }
  return { translate, detect: detectLocal };
}

module.exports = { createLocalTranslator, detectLocal };
