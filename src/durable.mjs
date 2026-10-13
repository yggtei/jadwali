/**
 * Durable Objects:
 *  - KeyStore:   يولّد مفاتيح VAPID مرة وحدة ويحفظها (ما تحتاج تسويها بنفسك).
 *  - Subscriber: كائن لكل جهاز مشترك. يحفظ إعداداته وجدوله، ويصحى بمنبّه (alarm)
 *                في الأوقات المطلوبة فقط (قبل كل محاضرة + كل 15 دقيقة لتحديث الجدول)،
 *                ثم يرسل إشعار Web Push للجهاز.
 */
import ICS from '../public/ics.js';
import { ICS_HOSTS, hostOk, json, b64u } from './shared.mjs';

const MIN = 60000;
const DAY = 86400000;
const REFRESH_MS = 15 * MIN;   // كل كم نجلب الجدول من Outlook
const HORIZON_URL = 14;        // أيام نحتفظ بها من الجدول (رابط)
const HORIZON_FILE = 120;      // أيام نحتفظ بها من الجدول (ملف مرفوع)

/* ======================= KeyStore ======================= */
export class KeyStore {
  constructor(state) {
    this.state = state;
    state.blockConcurrencyWhile(async () => {
      let k = await state.storage.get('keys');
      if (!k) {
        const pair = await crypto.subtle.generateKey({ name: 'ECDSA', namedCurve: 'P-256' }, true, ['sign', 'verify']);
        const privateJwk = await crypto.subtle.exportKey('jwk', pair.privateKey);
        const raw = new Uint8Array(await crypto.subtle.exportKey('raw', pair.publicKey));
        k = { publicKey: b64u(raw), privateJwk };
        await state.storage.put('keys', k);
      }
      this.keys = k;
    });
  }
  async fetch() { return json(this.keys); }
}

/* ======================= نصوص الإشعارات ======================= */
const TXT = {
  ar: {
    lecture: 'محاضرة', task: 'موعد مهمة', now: 'الآن', inMin: m => `بعد ${m} دقيقة`,
    changed: 'تغيّر جدولك',
    k: { removed: 'حُذفت', cancelled: 'أُلغيت', moved: 'تغيّر موعدها', room: 'تغيّرت القاعة', added: 'محاضرة جديدة' },
    testTitle: 'جدولي', testBody: 'تنبيه تجريبي: وصلك والتطبيق مقفل ✓', loc: 'ar-AE-u-nu-latn'
  },
  en: {
    lecture: 'Lecture', task: 'Task due', now: 'Now', inMin: m => `in ${m} min`,
    changed: 'Your schedule changed',
    k: { removed: 'Removed', cancelled: 'Cancelled', moved: 'Rescheduled', room: 'Room changed', added: 'New' },
    testTitle: 'My Schedule', testBody: 'Test alert: it reached you while the app was closed ✓', loc: 'en-GB'
  }
};

function fmtTime(ms, cfg) {
  const opts = { hour: 'numeric', minute: '2-digit', hour12: !!cfg.hour12 };
  const loc = TXT[cfg.lang].loc;
  try { return new Date(ms).toLocaleTimeString(loc, { ...opts, timeZone: cfg.tz }); }
  catch { return new Date(ms).toLocaleTimeString(loc, { ...opts, timeZone: 'UTC' }); }
}
function fmtDayTime(ms, cfg) {
  const loc = TXT[cfg.lang].loc;
  try {
    return new Date(ms).toLocaleString(loc, { weekday: 'short', day: 'numeric', month: 'short', hour: 'numeric', minute: '2-digit', hour12: !!cfg.hour12, timeZone: cfg.tz });
  } catch { return new Date(ms).toISOString(); }
}

/* ======================= تنقية المدخلات ======================= */
const clampInt = (v, lo, hi, d) => (Number.isFinite(+v) && v !== '' && v !== null ? Math.min(hi, Math.max(lo, Math.round(+v))) : d);

function validTz(tz) {
  if (typeof tz !== 'string' || !tz || tz.length > 64) return 'UTC';
  try { new Intl.DateTimeFormat('en', { timeZone: tz }); return tz; } catch { return 'UTC'; }
}

export function sanitize(b) {
  const cfg = {
    endpoint: String(b.endpoint || ''),
    origin: typeof b.origin === 'string' ? b.origin : '',
    icsUrl: '',
    reminders: [],
    taskRemind: clampInt(b.taskRemind, 0, 10080, 60),
    lang: b.lang === 'en' ? 'en' : 'ar',
    tz: validTz(b.tz),
    hour12: !!b.hour12,
    courses: {},
    tasks: []
  };
  let badHost = false;
  if (typeof b.icsUrl === 'string' && b.icsUrl.trim()) {
    try {
      const u = new URL(b.icsUrl.trim().replace(/^webcal:/i, 'https:'));
      if (u.protocol === 'https:' && hostOk(u.hostname, ICS_HOSTS)) cfg.icsUrl = u.toString();
      else badHost = true;
    } catch { badHost = true; }
  }
  if (Array.isArray(b.reminders)) {
    cfg.reminders = [...new Set(b.reminders.map(x => clampInt(x, 0, 1440, -1)).filter(x => x >= 0))].slice(0, 8);
  }
  if (b.courses && typeof b.courses === 'object') {
    Object.entries(b.courses).slice(0, 200).forEach(([name, v]) => {
      if (!v || typeof v !== 'object') return;
      const rem = v.remind === 'off' ? 'off' : (Number.isInteger(v.remind) ? clampInt(v.remind, 0, 1440, null) : null);
      if (v.hidden || rem !== null) cfg.courses[name.slice(0, 200)] = { hidden: !!v.hidden, remind: rem };
    });
  }
  if (Array.isArray(b.tasks)) {
    cfg.tasks = b.tasks.slice(0, 200).map(k => ({
      id: String(k && k.id || '').slice(0, 40),
      title: String(k && k.title || '').slice(0, 120),
      due: Number(k && k.due),
      rem: Array.isArray(k && k.rem) ? [...new Set(k.rem.map(x => clampInt(x, 0, 10080, -1)).filter(x => x >= 0))].slice(0, 6) : []
    })).filter(k => k.id && Number.isFinite(k.due));
  }
  return { cfg, badHost };
}

/* ======================= الجدول ======================= */
export function parseToOcc(text, cfg, now, days) {
  return ICS.parse(text, now - DAY, now + days * DAY, { tz: cfg.tz })
    .map(e => ({ i: e.id, u: e.uid, t: e.title, l: e.location, s: e.start, e: e.end, c: e.cancelled ? 1 : 0, a: e.allDay ? 1 : 0 }))
    .slice(0, 1500);
}

export function diffOcc(old, neu, now) {
  if (!old || !old.length) return [];
  const limit = now + 7 * DAY, nm = new Map(neu.map(o => [o.i, o])), out = [], removed = [], added = [];
  for (const o of old) {
    if (o.s < now || o.s > limit) continue;
    const n = nm.get(o.i);
    if (!n) { removed.push(o); continue; }
    if (n.c && !o.c) out.push({ k: 'cancelled', t: n.t, s: n.s });
    else if (n.s !== o.s || n.e !== o.e) out.push({ k: 'moved', t: n.t, s: n.s });
    else if (n.l !== o.l) out.push({ k: 'room', t: n.t, s: n.s, d: n.l || '—' });
  }
  const oldIds = new Set(old.map(o => o.i));
  for (const n of neu) if (!oldIds.has(n.i) && n.s >= now && n.s <= limit) added.push(n);
  for (const r of removed) {
    const j = added.findIndex(a => a.u && a.u === r.u && a.t === r.t);
    if (j >= 0) { out.push({ k: 'moved', t: r.t, s: added[j].s }); added.splice(j, 1); }
    else out.push({ k: 'removed', t: r.t, s: r.s });
  }
  added.forEach(a => out.push({ k: 'added', t: a.t, s: a.s }));
  return out.slice(0, 30);
}

/* ======================= VAPID ======================= */
async function vapidAuth(endpoint, keys, subject) {
  const header = b64u(new TextEncoder().encode(JSON.stringify({ typ: 'JWT', alg: 'ES256' })));
  const payload = b64u(new TextEncoder().encode(JSON.stringify({
    aud: new URL(endpoint).origin,
    exp: Math.floor(Date.now() / 1000) + 12 * 3600,
    sub: subject
  })));
  const key = await crypto.subtle.importKey('jwk', keys.privateJwk, { name: 'ECDSA', namedCurve: 'P-256' }, false, ['sign']);
  const sig = await crypto.subtle.sign({ name: 'ECDSA', hash: 'SHA-256' }, key, new TextEncoder().encode(header + '.' + payload));
  return `vapid t=${header}.${payload}.${b64u(sig)}, k=${keys.publicKey}`;
}

const freshMeta = () => ({ sent: {}, pending: [], lastFetch: 0, lastError: null, fixed: false, retryAt: null, retries: 0, testAt: null });

/* ======================= Subscriber ======================= */
export class Subscriber {
  constructor(state, env) { this.state = state; this.env = env; this.keys = null; }

  async fetch(request) {
    const path = new URL(request.url).pathname;
    const body = await request.json().catch(() => ({}));
    try {
      if (path === '/sync') return json(await this.sync(body));
      if (path === '/pending') return json({ items: await this.takePending() });
      if (path === '/unsubscribe') { await this.wipe(); return json({ ok: true }); }
      if (path === '/test') return json(await this.test());
    } catch (e) {
      return json({ error: String((e && e.message) || e).slice(0, 160) }, 500);
    }
    return json({ error: 'not_found' }, 404);
  }

  async vapid() {
    if (this.keys) return this.keys;
    const stub = this.env.KEYS.get(this.env.KEYS.idFromName('main'));
    this.keys = await (await stub.fetch('https://do/keys')).json();
    return this.keys;
  }

  async wipe() {
    await this.state.storage.deleteAlarm();
    await this.state.storage.deleteAll();
  }

  /* ---- مزامنة الإعدادات من التطبيق ---- */
  async sync(b) {
    const st = this.state.storage, now = Date.now();
    const { cfg, badHost } = sanitize(b);
    const prev = await st.get('cfg');
    await st.put('cfg', cfg);
    const meta = (await st.get('meta')) || freshMeta();
    let occ = (await st.get('occ')) || null;
    let icsOk = null, error = badHost ? 'ics_host_not_allowed' : null;

    if (!cfg.icsUrl && typeof b.icsText === 'string' && b.icsText.length > 0 && b.icsText.length < 600000) {
      try {
        occ = parseToOcc(b.icsText, cfg, now, HORIZON_FILE);
        meta.fixed = true; meta.lastFetch = now; icsOk = true;
      } catch (e) { icsOk = false; error = 'bad_ics'; }
    } else if (cfg.icsUrl) {
      meta.fixed = false;
      if (!prev || prev.icsUrl !== cfg.icsUrl || prev.tz !== cfg.tz || !occ) {
        try { occ = (await this.refresh(cfg, [], meta, now)).occ; icsOk = true; }
        catch (e) { icsOk = false; error = String((e && e.message) || e).slice(0, 80); meta.lastFetch = now - REFRESH_MS + MIN; }
      } else icsOk = !meta.lastError;
    } else if (!meta.fixed) {
      occ = [];
    }
    occ = occ || [];
    await st.put('occ', occ);
    await st.put('meta', meta);
    await this.schedule(cfg, occ, meta, now);
    return { ok: true, icsOk, error, count: occ.length };
  }

  async refresh(cfg, oldOcc, meta, now) {
    const res = await fetch(cfg.icsUrl, { headers: { 'User-Agent': 'Mozilla/5.0 MySchedule' }, redirect: 'follow' });
    if (!res.ok) throw new Error('HTTP ' + res.status);
    const text = await res.text();
    if (!/BEGIN:VCALENDAR/i.test(text)) throw new Error('not_ics');
    const occ = parseToOcc(text, cfg, now, HORIZON_URL);
    meta.lastFetch = now; meta.lastError = null;
    return { occ, changes: diffOcc(oldOcc, occ, now) };
  }

  async test() {
    const cfg = await this.state.storage.get('cfg');
    if (!cfg) return { ok: false, error: 'no_subscription' };
    const meta = (await this.state.storage.get('meta')) || freshMeta();
    meta.testAt = Date.now() + 20000;
    await this.state.storage.put('meta', meta);
    const occ = (await this.state.storage.get('occ')) || [];
    await this.schedule(cfg, occ, meta, Date.now());
    return { ok: true };
  }

  async takePending() {
    const meta = await this.state.storage.get('meta');
    if (!meta) return [];
    const cutoff = Date.now() - 3600000;
    const items = (meta.pending || []).filter(x => !x.ts || x.ts > cutoff);
    meta.pending = [];
    await this.state.storage.put('meta', meta);
    return items;
  }

  /* ---- تحديد الإشعارات المستحقة الآن ---- */
  remindersFor(cfg, o) {
    const c = cfg.courses[o.t];
    if (c && c.hidden) return [];
    if (c && c.remind === 'off') return [];
    if (c && Number.isInteger(c.remind)) return [c.remind];
    return cfg.reminders;
  }

  collect(cfg, occ, meta, now, out) {
    const T = TXT[cfg.lang];
    for (const o of occ) {
      if (o.c || o.a) continue;
      for (const m of this.remindersFor(cfg, o)) {
        const key = 'e|' + o.i + '|' + m, fire = o.s - m * MIN;
        if (meta.sent[key] || now < fire - 1500) continue;
        meta.sent[key] = now;
        if (now > (m === 0 ? o.s + 2 * MIN : o.s)) continue; // فاتت
        const left = Math.max(0, Math.ceil((o.s - now) / MIN));
        out.push({
          tag: key, lang: cfg.lang, dir: cfg.lang === 'ar' ? 'rtl' : 'ltr', ts: now,
          title: `${T.lecture}: ${o.t}`,
          body: `${left > 0 ? T.inMin(left) : T.now} · ${fmtTime(o.s, cfg)}${o.l ? ' · ' + o.l : ''}`
        });
      }
    }
    for (const k of cfg.tasks) for (const m of (k.rem && k.rem.length ? k.rem : [cfg.taskRemind])) {
      const key = 't|' + k.id + '|' + m, fire = k.due - m * MIN;
      if (meta.sent[key] || now < fire - 1500) continue;
      meta.sent[key] = now;
      if (now > k.due + 6 * 3600000) continue;
      out.push({ tag: key, lang: cfg.lang, dir: cfg.lang === 'ar' ? 'rtl' : 'ltr', ts: now, title: `${T.task}: ${k.title}`, body: fmtDayTime(k.due, cfg) });
    }
    if (meta.testAt && now >= meta.testAt - 1500) {
      meta.testAt = null;
      out.push({ tag: 'test', lang: cfg.lang, dir: cfg.lang === 'ar' ? 'rtl' : 'ltr', ts: now, title: T.testTitle, body: T.testBody });
    }
  }

  changeItem(cfg, changes, now) {
    const T = TXT[cfg.lang];
    const lines = changes.slice(0, 3).map(c => `${T.k[c.k]}: ${c.t}${c.k === 'room' ? ' → ' + c.d : ''}`);
    if (changes.length > 3) lines.push(`+${changes.length - 3}`);
    return { tag: 'sched-change', lang: cfg.lang, dir: cfg.lang === 'ar' ? 'rtl' : 'ltr', ts: now, title: T.changed, body: lines.join('\n') };
  }

  /* ---- أقرب وقت نحتاج نصحى فيه ---- */
  async schedule(cfg, occ, meta, now) {
    let next = Infinity;
    if (cfg.icsUrl && !meta.fixed) next = Math.min(next, meta.lastFetch + REFRESH_MS);
    for (const o of occ) {
      if (o.c || o.a) continue;
      for (const m of this.remindersFor(cfg, o)) {
        const fire = o.s - m * MIN;
        if (!meta.sent['e|' + o.i + '|' + m] && fire > now) next = Math.min(next, fire);
      }
    }
    for (const k of cfg.tasks) for (const m of (k.rem && k.rem.length ? k.rem : [cfg.taskRemind])) {
      const fire = k.due - m * MIN;
      if (!meta.sent['t|' + k.id + '|' + m] && fire > now) next = Math.min(next, fire);
    }
    if (meta.testAt) next = Math.min(next, meta.testAt);
    if (meta.retryAt) next = Math.min(next, meta.retryAt);
    if (next === Infinity) await this.state.storage.deleteAlarm();
    else await this.state.storage.setAlarm(Math.max(next, now + 1000));
  }

  async push(cfg) {
    const keys = await this.vapid();
    const auth = await vapidAuth(cfg.endpoint, keys, cfg.origin || 'https://example.com');
    const res = await fetch(cfg.endpoint, { method: 'POST', headers: { Authorization: auth, TTL: '120', Urgency: 'high' } });
    return res.status;
  }

  /* ---- المنبّه ---- */
  async alarm() {
    const st = this.state.storage;
    const cfg = await st.get('cfg');
    if (!cfg) return;
    let occ = (await st.get('occ')) || [];
    const meta = (await st.get('meta')) || freshMeta();
    const now = Date.now(), out = [];

    if (cfg.icsUrl && !meta.fixed && now >= meta.lastFetch + REFRESH_MS - 2000) {
      try {
        const r = await this.refresh(cfg, occ, meta, now);
        occ = r.occ; await st.put('occ', occ);
        if (r.changes.length) out.push(this.changeItem(cfg, r.changes, now));
      } catch (e) {
        meta.lastError = String((e && e.message) || e).slice(0, 100);
        meta.lastFetch = now - REFRESH_MS + 5 * MIN; // نعيد المحاولة بعد 5 دقائق
      }
    }

    this.collect(cfg, occ, meta, now, out);
    if (out.length) meta.pending = [...(meta.pending || []), ...out].slice(-20);

    if (meta.pending.length && (out.length || (meta.retryAt && now >= meta.retryAt - 1500))) {
      meta.retryAt = null;
      let status = 0;
      try { status = await this.push(cfg); } catch { status = 0; }
      if (status === 404 || status === 410) { await this.wipe(); return; }   // الاشتراك انتهى
      if (status >= 200 && status < 300) { meta.retries = 0; meta.lastError = null; }
      else if (status === 0 || status === 429 || status >= 500) {
        meta.retries = (meta.retries || 0) + 1;
        if (meta.retries > 5) { meta.pending = []; meta.retries = 0; }
        else meta.retryAt = now + MIN;
      } else { meta.lastError = 'push_' + status; meta.pending = []; }
    }

    const cut = now - 3 * DAY;
    Object.keys(meta.sent).forEach(k => { if (meta.sent[k] < cut) delete meta.sent[k]; });
    await st.put('meta', meta);
    await this.schedule(cfg, occ, meta, now);
  }
}
