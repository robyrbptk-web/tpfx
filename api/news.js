const FEEDS = [
  {
    name: 'google-en',
    url: 'https://news.google.com/rss/search?q=XAUUSD&hl=en-US&gl=US&ceid=US:en'
  },
  {
    name: 'google-id',
    url: 'https://news.google.com/rss/search?q=XAUUSD&hl=id&gl=ID&ceid=ID:id'
  }
];

const FALLBACK = 'https://www.bing.com/news/search?q=XAUUSD&format=rss';
const PRICE_URL = 'https://api.gold-api.com/price/XAU';

const MAX_ITEMS = 6;
const MAX_TITLE = 180;
const CACHE_MS = 5 * 60 * 1000;
const PRICE_CACHE_MS = 60 * 1000;
const REQUEST_TIMEOUT = 7000;
const MAX_AGE_MS = 21 * 24 * 60 * 60 * 1000;

let cache = { at: 0, items: [] };
let priceCache = { at: 0, value: 0, updated: 0 };

function decodeXml(value) {
  return String(value == null ? '' : value)
    .replace(/<!\[CDATA\[([\s\S]*?)\]\]>/g, '$1')
    .replace(/&#x([0-9a-f]+);/gi, (m, hex) => {
      const code = parseInt(hex, 16);
      return Number.isFinite(code) ? String.fromCodePoint(code) : m;
    })
    .replace(/&#(\d+);/g, (m, num) => {
      const code = Number(num);
      return Number.isFinite(code) ? String.fromCodePoint(code) : m;
    })
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&nbsp;/g, ' ')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&amp;/g, '&');
}

function tag(xml, name) {
  const pattern = new RegExp('<' + name + '(?:\\s[^>]*)?>([\\s\\S]*?)<\\/' + name + '>', 'i');
  const match = String(xml).match(pattern);
  return match ? decodeXml(match[1]) : '';
}

function cleanText(value, max) {
  const text = String(value == null ? '' : value)
    .replace(/<[^>]*>/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  return text.length > max ? text.slice(0, max - 1).trimEnd() + '…' : text;
}

function directUrl(value) {
  const link = String(value || '').trim();
  if (!link) return '';
  let url;
  try {
    url = new URL(link);
  } catch (e) {
    return '';
  }
  if (url.protocol !== 'http:' && url.protocol !== 'https:') return '';
  const host = url.hostname.toLowerCase();
  if (host === 'bing.com' || host.endsWith('.bing.com')) {
    const target = url.searchParams.get('url');
    if (target) {
      try {
        const next = new URL(target);
        if (next.protocol === 'http:' || next.protocol === 'https:') return next.href;
      } catch (e) {}
    }
  }
  return url.href;
}

function parseFeed(xml) {
  const items = [];
  const chunks = String(xml).split(/<item>/i).slice(1);
  chunks.forEach(chunk => {
    const block = chunk.split(/<\/item>/i)[0];
    let title = cleanText(tag(block, 'title'), MAX_TITLE);
    if (!title) return;
    // Buang sinyal trading personal yang tidak informatif.
    if (/\$?xauusd\s+(?:jual|beli)\s+di\b/i.test(title)) return;
    const source = cleanText(tag(block, 'source') || tag(block, 'News:Source'), 60);
    let sourceLabel = source;
    if (/^https?:\/\//i.test(sourceLabel)) {
      try {
        sourceLabel = new URL(sourceLabel).hostname.replace(/^www\./i, '');
      } catch (e) {
        sourceLabel = '';
      }
    }
    if (source) {
      const suffix = ' - ' + source;
      while (title.toLowerCase().endsWith(suffix.toLowerCase())) {
        title = title.slice(0, -suffix.length).trim();
      }
    }
    const link = directUrl(tag(block, 'link'));
    if (!link) return;
    const ts = Date.parse(tag(block, 'pubDate'));
    items.push({
      title,
      url: link,
      source: sourceLabel || 'Gold News',
      ts: Number.isFinite(ts) ? ts : 0
    });
  });
  return items;
}

function keyOf(item) {
  return item.title.toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
}

function prepare(items) {
  const seen = new Set();
  const fresh = [];
  const floor = Date.now() - MAX_AGE_MS;
  items
    .filter(item => item.title && item.url)
    .sort((a, b) => b.ts - a.ts)
    .forEach(item => {
      if (item.ts && item.ts < floor) return;
      const key = keyOf(item);
      if (!key || seen.has(key)) return;
      seen.add(key);
      fresh.push(item);
    });
  return fresh;
}

// Ambil bergiliran dari tiap feed agar berita XAUUSD tidak tenggelam
// oleh berita harga emas Antam yang lebih sering terbit.
function balanced(buckets) {
  const out = [];
  const seen = new Set();
  for (let rank = 0; rank < MAX_ITEMS * 2 && out.length < MAX_ITEMS; rank++) {
    for (const bucket of buckets) {
      const item = bucket[rank];
      if (!item) continue;
      const key = keyOf(item);
      if (seen.has(key)) continue;
      seen.add(key);
      out.push(item);
      if (out.length >= MAX_ITEMS) break;
    }
  }
  return out.sort((a, b) => b.ts - a.ts);
}

async function fetchText(url) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), REQUEST_TIMEOUT);
  try {
    const response = await fetch(url, {
      signal: controller.signal,
      redirect: 'follow',
      headers: {
        'User-Agent': 'Mozilla/5.0 (compatible; TPFXNews/1.0)',
        Accept: 'application/rss+xml, application/xml, text/xml;q=0.9, */*;q=0.8'
      }
    });
    if (!response.ok) throw new Error('HTTP ' + response.status);
    return await response.text();
  } finally {
    clearTimeout(timer);
  }
}

async function loadNews() {
  const results = await Promise.allSettled(FEEDS.map(feed => fetchText(feed.url)));
  const buckets = results.map(result => (
    result.status === 'fulfilled' ? prepare(parseFeed(result.value)) : []
  ));
  if (buckets.reduce((total, bucket) => total + bucket.length, 0) < 3) {
    try {
      buckets.push(prepare(parseFeed(await fetchText(FALLBACK))));
    } catch (e) {}
  }
  return balanced(buckets).slice(0, MAX_ITEMS);
}

async function loadPrice() {
  if (priceCache.value && Date.now() - priceCache.at < PRICE_CACHE_MS) return priceCache;
  try {
    const response = await fetch(PRICE_URL, {
      headers: { Accept: 'application/json' },
      signal: AbortSignal.timeout(REQUEST_TIMEOUT)
    });
    if (!response.ok) throw new Error('HTTP ' + response.status);
    const data = await response.json();
    const value = Number(data && data.price);
    if (!Number.isFinite(value) || value <= 0) throw new Error('Harga tidak valid');
    priceCache = {
      at: Date.now(),
      value,
      updated: Date.parse(data.updatedAt) || Date.now()
    };
  } catch (e) {}
  return priceCache;
}

module.exports = async (req, res) => {
  res.setHeader('Content-Type', 'application/json; charset=utf-8');
  res.setHeader('Cache-Control', 'public, max-age=60, s-maxage=60, stale-while-revalidate=300');

  if (req.method === 'OPTIONS') {
    res.statusCode = 204;
    return res.end();
  }
  if (req.method && req.method !== 'GET' && req.method !== 'HEAD') {
    res.statusCode = 405;
    res.setHeader('Allow', 'GET, HEAD');
    return res.end(JSON.stringify({ ok: false, items: [], err: 'Method not allowed' }));
  }

  const now = Date.now();
  const tasks = [loadPrice()];
  if (!cache.items.length || now - cache.at > CACHE_MS) {
    tasks.push((async () => {
      try {
        const items = await loadNews();
        if (items.length) cache = { at: Date.now(), items };
      } catch (e) {}
    })());
  }
  const results = await Promise.all(tasks);
  const price = results[0] || priceCache;

  if (!cache.items.length && !price.value) {
    res.statusCode = 200;
    return res.end(JSON.stringify({
      ok: false,
      items: [],
      price: null,
      updated: now,
      err: 'Sumber berita dan harga emas sedang tidak tersedia'
    }));
  }

  res.statusCode = 200;
  res.end(JSON.stringify({
    ok: true,
    items: cache.items,
    price: price.value
      ? { value: price.value, currency: 'USD', updated: price.updated, source: 'gold-api.com' }
      : null,
    updated: cache.at || now,
    stale: !cache.at || now - cache.at > CACHE_MS
  }));
};
