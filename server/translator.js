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
  async function call(method, text, target, textType) {
    if (!config.translatorKey) throw Object.assign(new Error('请在 .env 中配置 TRANSLATOR_KEY。'), { status: 503 });
    const url = new URL(method, `${config.translatorEndpoint.replace(/\/$/, '')}/`);
    url.searchParams.set('api-version', '3.0');
    if (target) url.searchParams.set('to', target);
    if (textType) url.searchParams.set('textType', textType);
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
    translateWord: async (paragraphs, to) => {
      const translated = [];
      const started = Date.now();
      for (let offset = 0; offset < paragraphs.length;) {
        if (Date.now() - started > 75000) throw new Error('全文翻译超时');
        const batch = [];
        let size = 0;
        while (offset < paragraphs.length && batch.length < 100 && size + paragraphs[offset].length <= 45000) {
          const paragraph = paragraphs[offset++];
          size += paragraph.length;
          batch.push(paragraph);
        }
        if (!batch.length) throw new Error('段落过长');
        // Send entire paragraphs in HTML mode so formatting spans share context.
        const response = await call('translate', batch, to, 'html');
        const values = Array.isArray(response) ? response.map(item => item.translations?.[0]?.text) : [];
        if (values.length !== batch.length || values.some(value => typeof value !== 'string' || !value.trim())) throw new Error('翻译结果不完整');
        for (let index = 0; index < values.length; index++) {
          let value = values[index];
          if (!hasWordMarkers(batch[index], value) || hasUntranslatedProse(batch[index], value)) {
            // HTML translation can move words outside spans, empty linked
            // spans or reorder them. Retry only this paragraph as plain text
            // fragments and rebuild its original markup locally. Mixed-language
            // paragraphs can also return unchanged when the provider detects
            // the target language for the entire paragraph.
            value = await translateHtml(batch[index], to, async texts => {
              if (Date.now() - started > 75000) throw new Error('全文翻译超时');
              const retry = await call('translate', texts, to);
              return Array.isArray(retry) ? retry.map(item => item.translations?.[0]?.text) : null;
            });
            if (!hasWordMarkers(batch[index], value)) throw new Error('译文格式标记无法安全对应原文，请重试。');
          }
          translated.push(value);
        }
      }
      if (translated.join('').length > 1000000) throw new Error('译文过长');
      return translated;
    },
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

function hasUntranslatedProse(source, translated) {
  const options = { xml: { xmlMode: false, decodeEntities: true } };
  const original = cheerio.load(source, options);
  const result = cheerio.load(translated, options);
  const returned = new Map(result('p > span').toArray().map(span => [span.attribs.id, result(span).text().trim()]));
  return original('p > span').toArray().some(span => {
    const text = original(span).text().trim();
    // Retry substantial unchanged prose, not isolated names such as Microsoft
    // Azure. This is a bounded retry, not a claim that all Latin text is wrong.
    const words = text.match(/\p{L}+/gu) || [];
    return words.length >= 5 && words.join('').length >= 40 && returned.get(span.attribs.id) === text;
  });
}

// Match the client's fail-closed mapping contract before returning paragraphs.
function hasWordMarkers(source, translated) {
  const options = { xml: { xmlMode: false, decodeEntities: true } };
  const original = cheerio.load(source, options);
  const result = cheerio.load(translated, options);
  const paragraphs = result.root().children();
  if (paragraphs.length !== 1 || paragraphs[0].name !== 'p') return false;
  const paragraph = paragraphs.first();
  const spans = paragraph.children();
  const expected = original('p').first().children();
  const looseText = nodes => nodes.toArray().some(node => node.type === 'text' && node.data.trim());
  if (spans.length !== expected.length || looseText(result.root().contents()) || looseText(paragraph.contents())) return false;
  return spans.toArray().every((span, i) => span.name === 'span' && span.attribs.id === `r${i}`
    && result(span).children().length === 0 && (!original(expected[i]).text().trim() || result(span).text().trim()));
}
module.exports = { translateHtml, detectHtml, createTranslator };
