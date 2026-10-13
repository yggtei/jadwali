/**
 * نقطة الدخول: يخدم الموقع (الملفات في public/) ويضيف:
 *   GET  /proxy?url=...     يجلب رابط ICS من Outlook (لأن Outlook يمنع المتصفح من قراءته مباشرة)
 *   GET  /api/vapid         المفتاح العام للإشعارات
 *   POST /api/sync          تسجيل الجهاز وإعدادات التنبيه
 *   POST /api/pending       الـ Service Worker يطلب نص الإشعار عند وصول push
 *   POST /api/test          إشعار تجريبي بعد 20 ثانية
 *   POST /api/unsubscribe   إيقاف التنبيهات
 */
import { ICS_HOSTS, PUSH_HOSTS, hostOk, json, sha256hex } from './shared.mjs';
export { Subscriber, KeyStore } from './durable.mjs';

export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    const path = url.pathname.replace(/\/+$/, '') || '/';
    if (path === '/proxy') return proxy(url, request);
    if (path.startsWith('/api/')) return api(path, request, env, url);
    return env.ASSETS.fetch(request);
  }
};

async function proxy(url, request) {
  const headers = { 'Cache-Control': 'no-store' };
  if (request.method !== 'GET') return new Response('Method not allowed', { status: 405, headers });
  const target = url.searchParams.get('url');
  if (!target) return new Response('Missing ?url=', { status: 400, headers });
  let u;
  try { u = new URL(target.replace(/^webcal:/i, 'https:')); }
  catch { return new Response('Bad url', { status: 400, headers }); }
  if (u.protocol !== 'https:' || !hostOk(u.hostname, ICS_HOSTS)) return new Response('Host not allowed', { status: 403, headers });
  try {
    const res = await fetch(u.toString(), { headers: { 'User-Agent': 'Mozilla/5.0 MySchedule' }, redirect: 'follow' });
    const text = await res.text();
    if (!res.ok) return new Response('Upstream error ' + res.status, { status: 502, headers });
    if (!/BEGIN:VCALENDAR/i.test(text)) return new Response('Not an ICS file', { status: 422, headers });
    return new Response(text, { headers: { ...headers, 'Content-Type': 'text/calendar; charset=utf-8' } });
  } catch {
    return new Response('Fetch failed', { status: 502, headers });
  }
}

const ROUTES = { '/api/sync': '/sync', '/api/pending': '/pending', '/api/unsubscribe': '/unsubscribe', '/api/test': '/test' };

async function api(path, request, env, url) {
  if (path === '/api/vapid') {
    if (request.method !== 'GET') return json({ error: 'method' }, 405);
    const k = await (await env.KEYS.get(env.KEYS.idFromName('main')).fetch('https://do/keys')).json();
    return json({ key: k.publicKey });
  }
  if (path === '/api/feedback') {
    if (request.method !== 'POST') return json({ error: 'method' }, 405);
    try {
      const b = JSON.parse((await request.text()).slice(0, 8000));
      console.log('FEEDBACK', JSON.stringify({ stars: Number(b.stars) || 0, text: String(b.text || '').slice(0, 1500), lang: String(b.lang || '').slice(0, 5), at: Date.now() }));
      return json({ ok: true });
    } catch { return json({ error: 'bad_json' }, 400); }
  }
  const route = ROUTES[path];
  if (!route) return json({ error: 'not_found' }, 404);
  if (request.method !== 'POST') return json({ error: 'method' }, 405);

  let body;
  try {
    const txt = await request.text();
    if (txt.length > 700000) return json({ error: 'too_large' }, 413);
    body = JSON.parse(txt);
  } catch { return json({ error: 'bad_json' }, 400); }

  let eu;
  try { eu = new URL(String(body.endpoint || '')); } catch { return json({ error: 'bad_endpoint' }, 400); }
  if (eu.protocol !== 'https:' || !hostOk(eu.hostname, PUSH_HOSTS)) return json({ error: 'endpoint_not_allowed' }, 400);

  // كل جهاز له كائن خاص به، واسمه مشتق من رابط الاشتراك (سرّي وما يتخمّن)
  const stub = env.SUB.get(env.SUB.idFromName(await sha256hex(body.endpoint)));
  return stub.fetch('https://do' + route, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ ...body, origin: url.origin })
  });
}
