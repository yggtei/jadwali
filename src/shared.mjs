// أدوات مشتركة بين الملفات

// المواقع المسموح جلب التقويم منها (حماية من استغلال الخادم)
export const ICS_HOSTS = [
  'outlook.office365.com',
  'outlook.office.com',
  'outlook.live.com',
  'calendar.live.com',
  'calendar.google.com'
];

// خوادم الإشعارات المعروفة للمتصفحات (Chrome/Android، Firefox، Safari/iPhone، Edge)
export const PUSH_HOSTS = [
  'fcm.googleapis.com',
  'android.googleapis.com',
  'updates.push.services.mozilla.com',
  'push.services.mozilla.com',
  'push.apple.com',
  'notify.windows.com'
];

export const hostOk = (host, list) => list.some(h => host === h || host.endsWith('.' + h));

export const json = (obj, status = 200) =>
  new Response(JSON.stringify(obj), {
    status,
    headers: { 'Content-Type': 'application/json; charset=utf-8', 'Cache-Control': 'no-store' }
  });

export function b64u(buf) {
  const b = buf instanceof Uint8Array ? buf : new Uint8Array(buf);
  let s = '';
  for (let i = 0; i < b.length; i++) s += String.fromCharCode(b[i]);
  return btoa(s).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
}

export async function sha256hex(text) {
  const d = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(text));
  return [...new Uint8Array(d)].map(x => x.toString(16).padStart(2, '0')).join('');
}
