const ID_RE = /^[A-Za-z0-9_-]{10,}$/;
const MAX_BYTES = 16 * 1024 * 1024;
const MIME_RE = /^(?:image\/(?:jpeg|png|webp|gif)|audio\/(?:webm|ogg|mp4|mpeg|mp3|wav|x-m4a|m4a)|video\/(?:mp4|webm)|application\/pdf|text\/plain|text\/csv|application\/msword|application\/vnd\.ms-excel|application\/vnd\.ms-powerpoint|application\/vnd\.openxmlformats-officedocument\.(?:wordprocessingml\.document|spreadsheetml\.sheet|presentationml\.presentation)|application\/zip|application\/x-zip-compressed|application\/vnd\.rar)$/;

function safeName(value) {
  return String(value || 'lampiran')
    .replace(/[\\/:*?"<>|\r\n\t]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, 120) || 'lampiran';
}

module.exports = async (req, res) => {
  if (req.method && req.method !== 'GET' && req.method !== 'HEAD') {
    res.statusCode = 405;
    res.setHeader('Allow', 'GET, HEAD');
    res.setHeader('Content-Type', 'application/json; charset=utf-8');
    return res.end(JSON.stringify({ ok: false, err: 'Method not allowed' }));
  }

  let url;
  try {
    url = new URL(req.url, 'https://tpfx.local');
  } catch (e) {
    url = new URL('https://tpfx.local/');
  }
  const id = String(url.searchParams.get('id') || '').trim();
  if (!ID_RE.test(id)) {
    res.statusCode = 400;
    res.setHeader('Content-Type', 'application/json; charset=utf-8');
    return res.end(JSON.stringify({ ok: false, err: 'ID lampiran tidak valid' }));
  }

  const requestedMime = String(url.searchParams.get('mime') || '').toLowerCase().split(';')[0].trim();
  const mime = MIME_RE.test(requestedMime) ? requestedMime : 'application/octet-stream';
  const name = safeName(url.searchParams.get('name'));
  const download = url.searchParams.get('dl') === '1';

  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), 15000);
  let upstream;
  let buffer;
  try {
    upstream = await fetch(
      'https://drive.usercontent.google.com/download?id=' + encodeURIComponent(id) + '&export=download&confirm=t',
      {
        signal: controller.signal,
        redirect: 'follow',
        headers: { 'User-Agent': 'Mozilla/5.0 (compatible; TPFXMedia/1.0)' }
      }
    );
    if (!upstream.ok) throw new Error('HTTP ' + upstream.status);
    const length = Number(upstream.headers.get('content-length')) || 0;
    if (length > MAX_BYTES) throw new Error('FILE_TOO_LARGE');
    buffer = Buffer.from(await upstream.arrayBuffer());
  } catch (e) {
    res.statusCode = e.message === 'FILE_TOO_LARGE' ? 413 : 502;
    res.setHeader('Content-Type', 'application/json; charset=utf-8');
    res.setHeader('Cache-Control', 'no-store');
    return res.end(JSON.stringify({
      ok: false,
      err: e.message === 'FILE_TOO_LARGE' ? 'Lampiran terlalu besar' : 'Lampiran tidak dapat diambil'
    }));
  } finally {
    clearTimeout(timer);
  }

  if (buffer.length > MAX_BYTES) {
    res.statusCode = 413;
    res.setHeader('Content-Type', 'application/json; charset=utf-8');
    return res.end(JSON.stringify({ ok: false, err: 'Lampiran terlalu besar' }));
  }

  res.setHeader('Content-Type', mime);
  res.setHeader('Accept-Ranges', 'bytes');
  res.setHeader('Cache-Control', 'public, max-age=31536000, s-maxage=31536000, immutable');
  res.setHeader('X-Content-Type-Options', 'nosniff');
  res.setHeader(
    'Content-Disposition',
    (download ? 'attachment' : 'inline') +
    '; filename="' + name.replace(/["\\]/g, '') + '"'
  );

  const range = req.headers && req.headers.range;
  if (range) {
    const match = /^bytes=(\d*)-(\d*)$/.exec(String(range).trim());
    if (match) {
      const total = buffer.length;
      let start = match[1] === '' ? -1 : Number(match[1]);
      let end = match[2] === '' ? -1 : Number(match[2]);
      if (start < 0) {
        start = Math.max(0, total - Math.max(0, end));
        end = total - 1;
      } else if (end < 0 || end >= total) {
        end = total - 1;
      }
      if (start >= total || start > end) {
        res.statusCode = 416;
        res.setHeader('Content-Range', 'bytes */' + total);
        res.setHeader('Content-Length', '0');
        return res.end();
      }
      res.statusCode = 206;
      res.setHeader('Content-Range', 'bytes ' + start + '-' + end + '/' + total);
      res.setHeader('Content-Length', String(end - start + 1));
      if (req.method === 'HEAD') return res.end();
      return res.end(buffer.subarray(start, end + 1));
    }
  }

  res.statusCode = 200;
  res.setHeader('Content-Length', String(buffer.length));
  if (req.method === 'HEAD') return res.end();
  res.end(buffer);
};
