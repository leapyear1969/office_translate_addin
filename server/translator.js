const cheerio = require('cheerio/slim');

function content(html) {
  const $ = cheerio.load(html, { xml: { xmlMode: false, decodeEntities: true, encodeEntities: 'utf8' } });
  const nodes = [];
  function visit(node, excluded = false) {
    const attrs = node.attribs || {};
    const skip = excluded || ['script', 'style', 'head', 'noscript', 'title'].includes(node.name)
      || (attrs.class || '').split(/\s+/).includes('notranslate') || attrs.translate === 'no';
    if (!skip && node.type === 'text' && node.data.trim()) nodes.push(node);
    if (node.children) node.children.forEach(child => visit(child, skip));
  }
  $.root().contents().each((_, node) => visit(node));
  return { $, nodes };
}

async function translateHtml(html, target, translate) {
  const started = Date.now();
  const { $, nodes } = content(html);
  const parts = [];
  for (const node of nodes) {
    let remaining = node.data;
    const translated = [];
    while (remaining.length) {
      let end = Math.min(4500, remaining.length);
      if (/[\uD800-\uDBFF]/.test(remaining[end - 1])) end--;
      parts.push({ text: remaining.slice(0, end), translated });
      remaining = remaining.slice(end);
    }
    node.translatedParts = translated;
  }
  for (let offset = 0; offset < parts.length;) {
    if (Date.now() - started > 75000) throw Object.assign(new Error('内容翻译耗时过长，原文保持不变，请重试。'), { status: 504 });
    const batch = [];
    let count = 0;
    while (offset < parts.length && batch.length < 100 && count + parts[offset].text.length <= 45000) {
      const part = parts[offset++];
      batch.push(part);
      count += part.text.length;
    }
    const results = await translate(batch.map(part => part.text), target);
    if (!Array.isArray(results) || results.length !== batch.length || results.some(t => typeof t !== 'string')) {
      throw new Error('翻译结果不完整，请重试。');
    }
    batch.forEach((part, index) => {
      const leading = part.text.match(/^\s*/)[0];
      const trailing = part.text.match(/\s*$/)[0];
      part.translated.push(part.text.trim() ? leading + results[index].trim() + trailing : part.text);
    });
  }
  if (!parts.length) return html;
  nodes.forEach(node => { node.data = node.translatedParts.join(''); });
  const result = $.html();
  if (result.length > 1000000) throw new Error('译文超过 1,000,000 字符限制。');
  return result;
}

async function detectHtml(html, detect) {
  const { nodes } = content(html);
  const text = nodes.map(node => node.data.trim()).join(' ').slice(0, 10000);
  return text ? detect(text) : null;
}

function createTranslator(config) {
  async function call(method, text, target) {
    if (!config.translatorKey) throw Object.assign(new Error('请在 .env 中配置 TRANSLATOR_KEY。'), { status: 503 });
    const url = new URL(method, `${config.translatorEndpoint.replace(/\/$/, '')}/`);
    url.searchParams.set('api-version', '3.0');
    if (target) url.searchParams.set('to', target);
    const headers = { 'Content-Type': 'application/json', 'Ocp-Apim-Subscription-Key': config.translatorKey };
    if (config.translatorRegion) headers['Ocp-Apim-Subscription-Region'] = config.translatorRegion;
    const response = await fetch(url, {
      method: 'POST', headers, body: JSON.stringify(text.map(Text => ({ Text }))),
      signal: AbortSignal.timeout(25000),
    });
    if (!response.ok) {
      const message = response.status === 429 ? '翻译请求过于频繁，请稍后重试。'
        : [401, 403].includes(response.status) ? '翻译服务认证失败，请检查 Key 和资源区域。'
          : `翻译服务暂时不可用（HTTP ${response.status}）。`;
      throw Object.assign(new Error(message), { status: 502 });
    }
    return response.json();
  }
  return {
    translate: (html, to) => translateHtml(html, to, async (texts, target) => {
      const response = await call('translate', texts, target);
      return Array.isArray(response) ? response.map(item => item.translations?.[0]?.text) : null;
    }),
    detect: html => detectHtml(html, async text => {
      const response = await call('detect', [text]);
      const result = response?.[0];
      if (typeof result?.language !== 'string' || typeof result?.score !== 'number') throw new Error('无法识别内容语言。');
      return { language: result.language, score: result.score };
    }),
  };
}
module.exports = { translateHtml, detectHtml, createTranslator };
