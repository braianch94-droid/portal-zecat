// Cloudflare Pages Function — proxy /portal/* → lupita.pythonanywhere.com/portal/*
// Users see portal-zecat.pages.dev throughout; PythonAnywhere is invisible.

const BACKEND = 'https://lupita.pythonanywhere.com';
const FRONTEND = 'https://portal-zecat.pages.dev';

export async function onRequest(context) {
  const url = new URL(context.request.url);

  // Build target URL preserving path and query
  const targetUrl = BACKEND + url.pathname + url.search;

  // Forward headers, set correct Host
  const reqHeaders = new Headers(context.request.headers);
  reqHeaders.set('Host', 'lupita.pythonanywhere.com');
  reqHeaders.set('X-Forwarded-Host', 'portal-zecat.pages.dev');
  // Remove Cloudflare-specific headers that could confuse the backend
  reqHeaders.delete('cf-ray');
  reqHeaders.delete('cf-visitor');
  reqHeaders.delete('cf-ipcountry');
  reqHeaders.delete('cf-connecting-ip');

  const body = ['GET', 'HEAD'].includes(context.request.method)
    ? undefined
    : context.request.body;

  const backendResp = await fetch(targetUrl, {
    method: context.request.method,
    headers: reqHeaders,
    body: body,
    redirect: 'manual',   // handle redirects ourselves so we can rewrite them
  });

  // Build response headers — rewrite any references to the backend domain
  const respHeaders = new Headers();
  for (const [key, val] of backendResp.headers.entries()) {
    const lk = key.toLowerCase();

    if (lk === 'location') {
      // Rewrite redirect destination so the browser stays on pages.dev
      respHeaders.set(key, val.replace(BACKEND, FRONTEND));

    } else if (lk === 'set-cookie') {
      // Strip Domain= attribute so the cookie is scoped to portal-zecat.pages.dev
      const cleaned = val.replace(/;\s*domain=[^;]*/gi, '');
      respHeaders.append(key, cleaned);

    } else if (lk === 'content-security-policy') {
      // Skip CSP that might block our proxied content
      // (uncomment if needed in future)
      // respHeaders.set(key, val);

    } else {
      respHeaders.append(key, val);
    }
  }

  return new Response(backendResp.body, {
    status: backendResp.status,
    statusText: backendResp.statusText,
    headers: respHeaders,
  });
}
