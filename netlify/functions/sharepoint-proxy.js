// netlify/functions/sharepoint-proxy.js
// Downloads Excel files from SharePoint server-side, bypassing browser CORS restrictions

const https = require('https');
const http = require('http');

exports.handler = async function(event, context) {
  // Only allow GET
  if (event.httpMethod !== 'GET') {
    return { statusCode: 405, body: 'Method Not Allowed' };
  }

  const targetUrl = event.queryStringParameters && event.queryStringParameters.url;
  if (!targetUrl) {
    return { statusCode: 400, body: 'Missing url parameter' };
  }

  // Only allow SharePoint and OneDrive domains for security
  let parsedUrl;
  try {
    parsedUrl = new URL(targetUrl);
  } catch(e) {
    return { statusCode: 400, body: 'Invalid URL' };
  }

  const allowedHosts = [
    'sharepoint.com',
    'onedrive.com',
    'onedrive.live.com',
    '1drv.ms',
    'microsoft.com',
  ];
  const isAllowed = allowedHosts.some(h => parsedUrl.hostname.endsWith(h));
  if (!isAllowed) {
    return { statusCode: 403, body: 'Domain not allowed' };
  }

  try {
    const data = await fetchWithRedirects(targetUrl, 5);
    return {
      statusCode: 200,
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Methods': 'GET',
        'Cache-Control': 'no-cache',
      },
      body: data.toString('base64'),
      isBase64Encoded: true,
    };
  } catch(err) {
    console.error('Proxy error:', err.message);
    return {
      statusCode: 500,
      headers: { 'Access-Control-Allow-Origin': '*' },
      body: JSON.stringify({ error: err.message }),
    };
  }
};

// Follows redirects (SharePoint often redirects before serving the file)
function fetchWithRedirects(url, maxRedirects) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    const protocol = url.startsWith('https') ? https : http;

    const req = protocol.get(url, {
      headers: {
        'User-Agent': 'Mozilla/5.0 (compatible; Marypaz-Proxy/1.0)',
        'Accept': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,*/*',
      }
    }, (res) => {
      // Follow redirects
      if ((res.statusCode === 301 || res.statusCode === 302 || res.statusCode === 307 || res.statusCode === 308)
          && res.headers.location && maxRedirects > 0) {
        const redirectUrl = res.headers.location.startsWith('http')
          ? res.headers.location
          : new URL(res.headers.location, url).toString();
        resolve(fetchWithRedirects(redirectUrl, maxRedirects - 1));
        return;
      }

      if (res.statusCode !== 200) {
        reject(new Error(`HTTP ${res.statusCode} from SharePoint`));
        return;
      }

      res.on('data', chunk => chunks.push(chunk));
      res.on('end', () => resolve(Buffer.concat(chunks)));
      res.on('error', reject);
    });

    req.on('error', reject);
    req.setTimeout(15000, () => {
      req.destroy();
      reject(new Error('Request timeout'));
    });
  });
}
