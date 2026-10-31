const CACHE_MS = 24 * 60 * 60 * 1000;
const CACHE_MAX = 800;
const TIMEOUT = 6500;
const ENDPOINT = 'https://nominatim.openstreetmap.org/reverse';

const cache = new Map();

function pickAddress(data) {
  const src = data && data.address && typeof data.address === 'object' ? data.address : {};
  const road = [src.road, src.house_number].filter(Boolean).join(' ');
  const parts = [];
  if (road) parts.push(road);
  [
    src.neighbourhood, src.suburb, src.village, src.hamlet,
    src.city_district, src.town, src.city, src.municipality,
    src.county, src.state
  ].forEach(value => {
    const text = String(value || '').trim();
    if (text && parts.indexOf(text) < 0) parts.push(text);
  });
  const short = parts.slice(0, 4).join(', ');
  if (short) return short;
  return String((data && data.display_name) || '')
    .split(',')
    .map(part => part.trim())
    .filter(Boolean)
    .slice(0, 3)
    .join(', ');
}

module.exports = async (req, res) => {
  res.setHeader('Content-Type', 'application/json; charset=utf-8');
  res.setHeader('Cache-Control', 'public, max-age=86400, s-maxage=604800, stale-while-revalidate=3600');

  if (req.method === 'OPTIONS') {
    res.statusCode = 204;
    return res.end();
  }
  if (req.method && req.method !== 'GET' && req.method !== 'HEAD') {
    res.statusCode = 405;
    res.setHeader('Allow', 'GET, HEAD');
    return res.end(JSON.stringify({ ok: false, address: '', err: 'Method not allowed' }));
  }

  const params = new URL(req.url || '/', 'http://local').searchParams;
  const lat = Number(params.get('lat'));
  const lng = Number(params.get('lng'));

  if (!Number.isFinite(lat) || !Number.isFinite(lng) ||
      Math.abs(lat) > 90 || Math.abs(lng) > 180) {
    res.statusCode = 400;
    return res.end(JSON.stringify({ ok: false, address: '', err: 'Koordinat tidak valid' }));
  }

  const key = lat.toFixed(5) + ',' + lng.toFixed(5);
  const hit = cache.get(key);
  if (hit && Date.now() - hit.at < CACHE_MS) {
    res.statusCode = 200;
    return res.end(JSON.stringify({ ok: !!hit.address, address: hit.address, source: 'cache' }));
  }

  const url = ENDPOINT + '?' + new URLSearchParams({
    format: 'jsonv2',
    lat: String(lat),
    lon: String(lng),
    zoom: '18',
    addressdetails: '1',
    'accept-language': 'id,en;q=0.8'
  }).toString();

  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), TIMEOUT);

  try {
    const response = await fetch(url, {
      signal: controller.signal,
      headers: {
        Accept: 'application/json',
        'Accept-Language': 'id,en;q=0.8',
        'User-Agent': 'TPFXWorkspace/1.0 (+https://tpfx.vercel.app)'
      }
    });
    if (!response.ok) throw new Error('HTTP ' + response.status);
    const data = await response.json();
    const address = pickAddress(data);
    cache.set(key, { at: Date.now(), address: address });
    if (cache.size > CACHE_MAX) {
      const oldest = cache.keys().next().value;
      if (oldest) cache.delete(oldest);
    }
    res.statusCode = 200;
    res.end(JSON.stringify({ ok: !!address, address: address, source: 'nominatim' }));
  } catch (e) {
    res.statusCode = 200;
    res.end(JSON.stringify({ ok: false, address: '', err: 'Alamat tidak tersedia' }));
  } finally {
    clearTimeout(timer);
  }
};
