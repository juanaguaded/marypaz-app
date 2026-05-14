// netlify/functions/sharepoint-proxy.js
// Downloads Excel files from SharePoint server-side, bypassing browser CORS restrictions

const https = require('https');
const http  = require('http');
const { URL } = require('url');

exports.handler = async function(event) {
  if (event.httpMethod !== 'GET') {
    return { statusCode: 405, body: 'Method Not Allowed' };
  }

  let targetUrl = event.queryStringParameters && event.queryStringParameters.url;
  if (!targetUrl) {
    return { statusCode: 400, body: 'Missing url parameter' };
  }

  // Security: only allow Microsoft domains
  let parsed;
  try { parsed = new URL(targetUrl); } catch(e) {
    return { statusCode: 400, body: 'Invalid URL' };
  }
  const allowed = ['sharepoint.com','onedrive.com','onedrive.live.com','1drv.ms','microsoft.com','microsoftonline.com'];
  if (!allowed.some(h => parsed.hostname.endsWith(h))) {
    return { statusCode: 403, body: 'Domain not allowed: ' + parsed.hostname };
  }

  // Force download=1 on SharePoint URLs (avoid preview page)
  if (parsed.hostname.includes('sharepoint.com')) {
    parsed.searchParams.set('download', '1');
    targetUrl = parsed.toString();
  }

  console.log('[proxy] Starting download:', targetUrl);

  try {
    const { buffer, contentType, finalUrl } = await fetchFollowRedirects(targetUrl, 10);

    console.log('[proxy] Final URL:', finalUrl);
    console.log('[proxy] Content-Type:', contentType);
    console.log('[proxy] Bytes received:', buffer.length);

    // Detect if SharePoint returned HTML instead of the file
    const isHtml = contentType && contentType.includes('text/html');
    const startsWithHtml = buffer.length > 0 && buffer.slice(0, 5).toString('utf8').trim().startsWith('<');

    if (isHtml || startsWithHtml) {
      console.error('[proxy] Got HTML instead of file — SharePoint returned a preview page');
      console.error('[proxy] First 500 chars:', buffer.slice(0, 500).toString('utf8'));
      return {
        statusCode: 502,
        headers: { 'Access-Control-Allow-Origin': '*', 'Content-Type': 'application/json' },
        body: JSON.stringify({
          error: 'SharePoint returned HTML instead of the file',
          contentType,
          finalUrl,
          preview: buffer.slice(0, 200).toString('utf8')
        })
      };
    }

    return {
      statusCode: 200,
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Methods': 'GET',
        'Cache-Control': 'no-cache',
      },
      body: buffer.toString('base64'),
      isBase64Encoded: true,
    };

  } catch(err) {
    console.error('[proxy] Error:', err.message);
    return {
      statusCode: 500,
      headers: { 'Access-Control-Allow-Origin': '*', 'Content-Type': 'application/json' },
      body: JSON.stringify({ error: err.message }),
    };
  }
};

// Follows up to maxRedirects redirects, returns buffer + headers
function fetchFollowRedirects(url, maxRedirects, cookieJar = '') {
  return new Promise((resolve, reject) => {
    const parsed = new URL(url);
    const protocol = parsed.protocol === 'https:' ? https : http;

    const options = {
      hostname: parsed.hostname,
      port:     parsed.port || (parsed.protocol === 'https:' ? 443 : 80),
      path:     parsed.pathname + parsed.search,
      method:   'GET',
      headers: {
        // Mimic a real browser to avoid SharePoint blocking
        'User-Agent':      'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Accept':          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/octet-stream,*/*',
        'Accept-Language': 'es-ES,es;q=0.9,en;q=0.8',
        'Accept-Encoding': 'identity', // avoid gzip so we get raw bytes
        ...(cookieJar ? { 'Cookie': cookieJar } : {}),
      },
      timeout: 20000,
    };

    const req = protocol.request(options, (res) => {
      console.log('[proxy] HTTP', res.statusCode, url.substring(0, 80));
      console.log('[proxy] Content-Type:', res.headers['content-type']);

      // Collect cookies for redirect chain (SharePoint auth uses cookies)
      let newCookies = cookieJar;
      if (res.headers['set-cookie']) {
        const cookies = res.headers['set-cookie'].map(c => c.split(';')[0]);
        newCookies = [cookieJar, ...cookies].filter(Boolean).join('; ');
      }

      // Follow redirects
      const isRedirect = [301,302,303,307,308].includes(res.statusCode);
      if (isRedirect && res.headers.location && maxRedirects > 0) {
        let nextUrl = res.headers.location;
        // Handle relative redirects
        if (!nextUrl.startsWith('http')) {
          nextUrl = new URL(nextUrl, url).toString();
        }
        console.log('[proxy] Redirect →', nextUrl.substring(0, 80));
        // Consume response body before following redirect
        res.resume();
        resolve(fetchFollowRedirects(nextUrl, maxRedirects - 1, newCookies));
        return;
      }

      if (res.statusCode !== 200) {
        res.resume();
        reject(new Error('HTTP ' + res.statusCode + ' from ' + url.substring(0, 60)));
        return;
      }

      const chunks = [];
      res.on('data', chunk => chunks.push(chunk));
      res.on('end', () => {
        const buffer = Buffer.concat(chunks);
        resolve({
          buffer,
          contentType: res.headers['content-type'] || '',
          finalUrl: url,
        });
      });
      res.on('error', reject);
    });

    req.on('error', reject);
    req.on('timeout', () => { req.destroy(); reject(new Error('Request timeout')); });
    req.end();
  });
}
