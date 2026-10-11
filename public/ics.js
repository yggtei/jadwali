/*
 * ics.js — قارئ ملفات ICS بسيط مع دعم التكرار (RRULE) والمناطق الزمنية
 * ICS reader with recurrence + timezone support. Works in browser and Node.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.ICS = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  // Windows timezone names (used by Outlook) -> fixed offset in minutes (no DST rules)
  var WIN_TZ = {
    'Arabian Standard Time': 240, 'Gulf Standard Time': 240, 'Arab Standard Time': 180,
    'Arabic Standard Time': 180, 'Iran Standard Time': 210, 'India Standard Time': 330,
    'Pakistan Standard Time': 300, 'UTC': 0, 'GMT Standard Time': 0, 'Greenwich Standard Time': 0,
    'Egypt Standard Time': 120, 'Turkey Standard Time': 180, 'Jordan Standard Time': 120,
    'W. Europe Standard Time': 60, 'Central Europe Standard Time': 60, 'Romance Standard Time': 60
  };
  // approximate IANA fallbacks if Intl can't resolve
  function ianaOffset(tz, utcMs) {
    try {
      var dtf = new Intl.DateTimeFormat('en-US', {
        timeZone: tz, hourCycle: 'h23', year: 'numeric', month: 'numeric', day: 'numeric',
        hour: 'numeric', minute: 'numeric', second: 'numeric'
      });
      var p = {};
      dtf.formatToParts(new Date(utcMs)).forEach(function (x) { p[x.type] = +x.value; });
      var asUTC = Date.UTC(p.year, p.month - 1, p.day, p.hour, p.minute, p.second);
      return Math.round((asUTC - Math.floor(utcMs / 1000) * 1000) / 60000);
    } catch (e) { return null; }
  }

  // wall-clock fields in a timezone -> epoch ms
  function wallToMs(y, mo, d, h, mi, s, tz, isUTC) {
    if (isUTC) return Date.UTC(y, mo - 1, d, h, mi, s);
    if (!tz) return new Date(y, mo - 1, d, h, mi, s).getTime(); // floating -> device local
    if (WIN_TZ.hasOwnProperty(tz)) return Date.UTC(y, mo - 1, d, h, mi, s) - WIN_TZ[tz] * 60000;
    var guess = Date.UTC(y, mo - 1, d, h, mi, s);
    var off = ianaOffset(tz, guess);
    if (off === null) return new Date(y, mo - 1, d, h, mi, s).getTime();
    var ms = guess - off * 60000;
    var off2 = ianaOffset(tz, ms);
    if (off2 !== null && off2 !== off) ms = guess - off2 * 60000;
    return ms;
  }

  function unfold(text) {
    return text.replace(/\r\n/g, '\n').replace(/\r/g, '\n').replace(/\n[ \t]/g, '').split('\n');
  }

  function unescapeText(s) {
    return (s || '').replace(/\\n/gi, '\n').replace(/\\,/g, ',').replace(/\\;/g, ';').replace(/\\\\/g, '\\');
  }

  function parseLine(line) {
    var idx = -1, inQ = false;
    for (var i = 0; i < line.length; i++) {
      var c = line[i];
      if (c === '"') inQ = !inQ;
      else if (c === ':' && !inQ) { idx = i; break; }
    }
    if (idx < 0) return null;
    var left = line.slice(0, idx), value = line.slice(idx + 1);
    var parts = left.split(';');
    var name = parts.shift().toUpperCase();
    var params = {};
    parts.forEach(function (p) {
      var eq = p.indexOf('=');
      if (eq > 0) params[p.slice(0, eq).toUpperCase()] = p.slice(eq + 1).replace(/^"|"$/g, '');
    });
    return { name: name, params: params, value: value };
  }

  var DEFAULT_TZ = null; // set per parse() call: timezone for floating (no TZID, no Z) times

  // returns {y,mo,d,h,mi,s,utc,tz,allDay,ms}
  function parseDate(value, params) {
    var m = /^(\d{4})(\d{2})(\d{2})(?:T(\d{2})(\d{2})(\d{2})?(Z)?)?$/.exec(value.trim());
    if (!m) return null;
    var allDay = !m[4];
    var o = {
      y: +m[1], mo: +m[2], d: +m[3], h: +(m[4] || 0), mi: +(m[5] || 0), s: +(m[6] || 0),
      utc: !!m[7], tz: (params && params.TZID) || null, allDay: allDay
    };
    if (allDay) o.tz = null;
    else if (!o.tz && !o.utc && DEFAULT_TZ) o.tz = DEFAULT_TZ;
    o.ms = wallToMs(o.y, o.mo, o.d, o.h, o.mi, o.s, o.tz, o.utc);
    return o;
  }

  function parseEvents(text) {
    var lines = unfold(text), events = [], cur = null, depth = 0;
    lines.forEach(function (ln) {
      if (!ln) return;
      var u = ln.toUpperCase();
      if (u === 'BEGIN:VEVENT') { cur = { exdates: [], raw: {} }; depth = 0; return; }
      if (u === 'END:VEVENT') { if (cur) events.push(cur); cur = null; return; }
      if (!cur) return;
      if (u.indexOf('BEGIN:') === 0) { depth++; return; }  // e.g. VALARM
      if (u.indexOf('END:') === 0) { depth--; return; }
      if (depth > 0) return;
      var p = parseLine(ln);
      if (!p) return;
      switch (p.name) {
        case 'UID': cur.uid = p.value; break;
        case 'SUMMARY': cur.title = unescapeText(p.value); break;
        case 'LOCATION': cur.location = unescapeText(p.value); break;
        case 'DESCRIPTION': cur.description = unescapeText(p.value); break;
        case 'STATUS': cur.status = p.value.toUpperCase(); break;
        case 'DTSTART': cur.start = parseDate(p.value, p.params); break;
        case 'DTEND': cur.end = parseDate(p.value, p.params); break;
        case 'DURATION': cur.duration = p.value; break;
        case 'RRULE': cur.rrule = p.value; break;
        case 'RECURRENCE-ID': cur.recId = parseDate(p.value, p.params); break;
        case 'EXDATE':
          p.value.split(',').forEach(function (v) {
            var d = parseDate(v, p.params); if (d) cur.exdates.push(d.ms);
          });
          break;
      }
    });
    return events.filter(function (e) { return e.start; });
  }

  function parseRRule(s) {
    var r = {};
    s.split(';').forEach(function (kv) {
      var i = kv.indexOf('='); if (i < 0) return;
      r[kv.slice(0, i).toUpperCase()] = kv.slice(i + 1);
    });
    return r;
  }

  var DAYS = { SU: 0, MO: 1, TU: 2, WE: 3, TH: 4, FR: 5, SA: 6 };

  function durationMs(ev) {
    if (ev.end) return Math.max(0, ev.end.ms - ev.start.ms);
    if (ev.duration) {
      var m = /^P(?:(\d+)W)?(?:(\d+)D)?(?:T(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?)?$/.exec(ev.duration);
      if (m) return (((+m[1] || 0) * 7 + (+m[2] || 0)) * 86400 + (+m[3] || 0) * 3600 + (+m[4] || 0) * 60 + (+m[5] || 0)) * 1000;
    }
    return ev.start.allDay ? 86400000 : 3600000;
  }

  // generate start ms of occurrences between fromMs and toMs
  function expand(ev, fromMs, toMs) {
    var st = ev.start, dur = durationMs(ev);
    var out = [];
    function mk(y, mo, d) { // y, mo(1-12), d may overflow -> normalise through Date
      var n = new Date(Date.UTC(y, mo - 1, d));
      return wallToMs(n.getUTCFullYear(), n.getUTCMonth() + 1, n.getUTCDate(), st.h, st.mi, st.s, st.tz, st.utc);
    }
    if (!ev.rrule) {
      if (st.ms <= toMs && st.ms + dur >= fromMs) out.push(st.ms);
      return { starts: out, dur: dur };
    }
    var r = parseRRule(ev.rrule);
    var interval = Math.max(1, +(r.INTERVAL || 1));
    var count = r.COUNT ? +r.COUNT : Infinity;
    var until = null;
    if (r.UNTIL) { var u = parseDate(r.UNTIL, {}); if (u) until = u.allDay ? Date.UTC(u.y, u.mo - 1, u.d, 23, 59, 59) : u.ms; }
    var byday = r.BYDAY ? r.BYDAY.split(',').map(function (x) { return DAYS[x.slice(-2)]; }).filter(function (x) { return x !== undefined; }) : null;
    var bymonthday = r.BYMONTHDAY ? r.BYMONTHDAY.split(',').map(Number) : null;
    var n = 0, guard = 0;
    function push(ms) {
      if (ms < st.ms) return true;
      if (until !== null && ms > until) return false;
      if (n >= count) return false;
      n++;
      if (ms > toMs) return false;
      if (ms + dur >= fromMs) out.push(ms);
      return true;
    }
    var base = Date.UTC(st.y, st.mo - 1, st.d);
    if (r.FREQ === 'DAILY') {
      for (var k = 0; guard++ < 5000; k += interval) { if (!push(mk(st.y, st.mo, st.d + k))) break; }
    } else if (r.FREQ === 'WEEKLY') {
      var days = (byday && byday.length ? byday : [new Date(base).getUTCDay()]).slice();
      days.sort(function (a, b) { return ((a + 6) % 7) - ((b + 6) % 7); }); // week starts Monday
      var startDow = new Date(base).getUTCDay();
      var mondayOffset = -((startDow + 6) % 7);
      var go = true;
      for (var w = 0; go && guard++ < 3000; w += interval) {
        for (var i = 0; i < days.length; i++) {
          var dayOff = mondayOffset + w * 7 + ((days[i] + 6) % 7);
          var ms = mk(st.y, st.mo, st.d + dayOff);
          if (ms < st.ms) continue;
          if (!push(ms)) { go = false; break; }
        }
      }
    } else if (r.FREQ === 'MONTHLY') {
      var md = bymonthday || [st.d];
      var go2 = true;
      for (var mth = 0; go2 && guard++ < 1500; mth += interval) {
        for (var j = 0; j < md.length; j++) {
          var dt = new Date(Date.UTC(st.y, st.mo - 1 + mth, md[j]));
          if (dt.getUTCDate() !== md[j]) continue;
          if (!push(mk(dt.getUTCFullYear(), dt.getUTCMonth() + 1, dt.getUTCDate()))) { go2 = false; break; }
        }
      }
    } else if (r.FREQ === 'YEARLY') {
      for (var yy = 0; guard++ < 200; yy += interval) { if (!push(mk(st.y + yy, st.mo, st.d))) break; }
    } else {
      if (st.ms <= toMs && st.ms + dur >= fromMs) out.push(st.ms);
    }
    return { starts: out, dur: dur };
  }

  /** parse(text, fromMs, toMs, {tz}) -> sorted list of occurrences */
  function parse(text, fromMs, toMs, opts) {
    DEFAULT_TZ = (opts && opts.tz) || null;
    try { return parseInner(text, fromMs, toMs); } finally { DEFAULT_TZ = null; }
  }
  function parseInner(text, fromMs, toMs) {
    var evs = parseEvents(text);
    var overrides = {};
    evs.forEach(function (e) {
      if (e.recId && e.uid) overrides[e.uid + '|' + e.recId.ms] = e;
    });
    var result = [];
    evs.forEach(function (e) {
      if (e.recId) return;
      var ex = expand(e, fromMs, toMs);
      ex.starts.forEach(function (s) {
        if (e.exdates.indexOf(s) >= 0) return;
        var origKey = (e.uid || e.title) + '|' + s;
        var o = overrides[(e.uid || '') + '|' + s];
        var src = o || e;
        var start = o ? o.start.ms : s;
        var end = o ? start + durationMs(o) : s + ex.dur;
        var title = src.title || e.title || '';
        var cancelled = src.status === 'CANCELLED' || /^(canceled|cancelled)\s*:/i.test(title);
        result.push({
          id: origKey,
          uid: e.uid || '',
          title: title.replace(/^(canceled|cancelled)\s*:\s*/i, ''),
          location: src.location || e.location || '',
          description: src.description || e.description || '',
          start: start, end: end,
          allDay: !!e.start.allDay,
          cancelled: cancelled
        });
      });
    });
    // overrides whose original occurrence fell outside the window but moved inside
    evs.forEach(function (e) {
      if (!e.recId) return;
      var key = (e.uid || e.title) + '|' + e.recId.ms;
      if (result.some(function (r) { return r.id === key; })) return;
      var end = e.start.ms + durationMs(e);
      if (e.start.ms <= toMs && end >= fromMs) {
        var title = e.title || '';
        result.push({
          id: key, uid: e.uid || '', title: title.replace(/^(canceled|cancelled)\s*:\s*/i, ''),
          location: e.location || '', description: e.description || '',
          start: e.start.ms, end: end, allDay: !!e.start.allDay,
          cancelled: e.status === 'CANCELLED' || /^(canceled|cancelled)\s*:/i.test(title)
        });
      }
    });
    result.sort(function (a, b) { return a.start - b.start || a.title.localeCompare(b.title); });
    return result;
  }

  return { parse: parse, parseEvents: parseEvents };
});
