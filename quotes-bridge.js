/* Connects the Quote Builder to the company's Outlook account.
   Saved quotes and settings live as entries in a separate "EAM Quotes" calendar.
   Runs inside the Quotes tab of the dispatch app and uses its Microsoft sign-in. */
(function () {
  "use strict";
  var E = null;
  try { E = window.parent !== window && window.parent.EAM; } catch (x) {}
  if (!E) { location.replace("./#quotes"); return; }

  var TZ = E.TZ, BASE = E.BASE, QCAL = "EAM Quotes";
  var SET_SUBJECT = "EAM Quote Builder settings (do not delete)";
  var BEGIN = "-----EAM-DATA-BEGIN-----", END = "-----EAM-DATA-END-----";
  var SEL = "$select=id,subject,body,lastModifiedDateTime";
  var calId = null, items = {}, byNo = {}, settingsId = null, settingsData = null;
  var state = "loading", qSubs = [], sSubs = [], chain = Promise.resolve();

  function b64e(s) { return btoa(unescape(encodeURIComponent(s))).replace(/(.{76})/g, "$1\n"); }
  function b64d(s) { return decodeURIComponent(escape(atob(s.replace(/[^A-Za-z0-9+/=]/g, "")))); }
  function pack(kind, data, human) { return human + "\n\n" + BEGIN + "\n" + b64e(JSON.stringify({ kind: kind, data: data })) + "\n" + END; }
  function unpack(ev) {
    var b = ev && ev.body, t = (b && b.content) || "";
    if (b && /html/i.test(b.contentType || "")) t = t.replace(/<[^>]+>/g, "\n").replace(/&#43;/g, "+").replace(/&#61;/g, "=").replace(/&#47;/g, "/");
    var i = t.indexOf(BEGIN), j = t.indexOf(END);
    if (i < 0 || j < i) return null;
    try { return JSON.parse(b64d(t.slice(i + BEGIN.length, j))); } catch (x) { return null; }
  }
  function nowLocal(addMin) {
    var d = new Date(Date.now() + (addMin || 0) * 60000);
    return new Intl.DateTimeFormat("sv-SE", { timeZone: TZ, year: "numeric", month: "2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit", second: "2-digit", hourCycle: "h23" }).format(d).replace(" ", "T");
  }
  function money(n) { return "$" + (Math.round((+n || 0) * 100) / 100).toLocaleString("en-AU", { minimumFractionDigits: 2, maximumFractionDigits: 2 }); }

  /* ---- Outlook storage ---- */
  function ensureCal() {
    if (calId) return Promise.resolve(calId);
    return E.gfetch(BASE + "/calendars?$select=id,name&$top=100").then(function (r) {
      var hit = ((r && r.value) || []).filter(function (c) { return c.name === QCAL; })[0];
      if (hit) return hit;
      return E.gfetch(BASE + "/calendars", { method: "POST", body: { name: QCAL } });
    }).then(function (c) { calId = c.id; return calId; });
  }
  function absorb(ev) {
    var p = unpack(ev); if (!p) return;
    items[ev.id] = { mod: ev.lastModifiedDateTime || "", kind: p.kind, data: p.data };
  }
  function reindex() {
    byNo = {}; settingsId = null; settingsData = null;
    Object.keys(items).forEach(function (id) {
      var it = items[id];
      if (it.kind === "settings") { if (!settingsId || it.mod > items[settingsId].mod) { settingsId = id; settingsData = it.data; } }
      else if (it.kind === "quote" && it.data && it.data.no) { if (!byNo[it.data.no] || it.mod > items[byNo[it.data.no]].mod) byNo[it.data.no] = id; }
    });
  }
  function pages(url, each) {
    function go(u, n) { return E.gfetch(u).then(function (r) { ((r && r.value) || []).forEach(each); if (r && r["@odata.nextLink"] && n < 20) return go(r["@odata.nextLink"], n + 1); }); }
    return go(url, 0);
  }
  function loadAll() {
    state = "loading";
    return ensureCal().then(function (id) {
      var fresh = {};
      return pages(BASE + "/calendars/" + encodeURIComponent(id) + "/events?$top=50&" + SEL, function (ev) { var p = unpack(ev); if (p) fresh[ev.id] = { mod: ev.lastModifiedDateTime || "", kind: p.kind, data: p.data }; })
        .then(function () { items = fresh; reindex(); state = "ok"; emit(); banner(); });
    }).catch(function (e) { state = "error"; qSubs.forEach(function (s) { try { s.err && s.err(e); } catch (x) {} }); banner(); });
  }
  function poll() {
    if (state !== "ok" || document.hidden) return;
    var seen = {}, changed = [];
    pages(BASE + "/calendars/" + encodeURIComponent(calId) + "/events?$top=200&$select=id,lastModifiedDateTime", function (ev) {
      seen[ev.id] = 1; if (!items[ev.id] || items[ev.id].mod !== (ev.lastModifiedDateTime || "")) changed.push(ev.id);
    }).then(function () {
      var gone = Object.keys(items).filter(function (id) { return !seen[id]; });
      gone.forEach(function (id) { delete items[id]; });
      return changed.slice(0, 40).reduce(function (p, id) {
        return p.then(function () { return E.gfetch(BASE + "/events/" + encodeURIComponent(id) + "?" + SEL).then(absorb, function () {}); });
      }, Promise.resolve()).then(function () { if (gone.length || changed.length) { reindex(); emit(); banner(); } });
    }).catch(function () {});
  }
  function quoteDocs() {
    return Object.keys(byNo).map(function (no) { return items[byNo[no]].data; })
      .sort(function (a, b) { return String(b.createdAt || "").localeCompare(String(a.createdAt || "")); }).slice(0, 500);
  }
  function emit() {
    var docs = quoteDocs().map(function (d) { return { id: d.no, exists: true, data: function () { return d; } }; });
    qSubs.forEach(function (s) { try { s.next({ docs: docs, size: docs.length, empty: !docs.length }); } catch (x) {} });
    sSubs.forEach(function (s) { try { s.next({ id: "main", exists: !!settingsData, data: function () { return settingsData || undefined; } }); } catch (x) {} });
  }
  function serial(fn) { chain = chain.then(fn, fn); return chain; }
  function saveQuote(d) {
    return serial(function () {
      return ensureCal().then(function (cid) {
        var human = ["Quote " + d.no, "Customer: " + (d.customerName || ""), "Move date: " + (d.moveFlexible ? "Flexible" : (d.moveDate || "")), "From: " + (d.from || ""), "To: " + (d.to || ""), "Total: " + money(d.total), "Status: " + (d.status || ""), "Rep: " + (d.rep || ""), "", "Saved by the EAM app. Open the app to view or change this quote."].join("\n");
        var subject = [d.no, (d.customerName || "").trim(), money(d.total), "[" + (d.status || "Sent") + "]"].filter(Boolean).join(" ");
        var body = { contentType: "text", content: pack("quote", d, human) };
        var id = byNo[d.no];
        var req = id ? E.gfetch(BASE + "/events/" + encodeURIComponent(id) + "?" + SEL, { method: "PATCH", body: { subject: subject, body: body } })
          : E.gfetch(BASE + "/calendars/" + encodeURIComponent(cid) + "/events", { method: "POST", body: { subject: subject, body: body, start: { dateTime: nowLocal(0), timeZone: TZ }, end: { dateTime: nowLocal(30), timeZone: TZ }, showAs: "free", isReminderOn: false } });
        return req.then(function (ev) {
          var eid = (ev && ev.id) || id;
          items[eid] = { mod: (ev && ev.lastModifiedDateTime) || "", kind: "quote", data: d };
          reindex(); emit();
        });
      });
    });
  }
  function saveSettings(obj) {
    return serial(function () {
      return ensureCal().then(function (cid) {
        var body = { contentType: "text", content: pack("settings", obj, "Rates, wording and business details for the EAM app.\nDo not delete or edit this entry. Change settings inside the app.") };
        var req = settingsId ? E.gfetch(BASE + "/events/" + encodeURIComponent(settingsId) + "?" + SEL, { method: "PATCH", body: { body: body } })
          : E.gfetch(BASE + "/calendars/" + encodeURIComponent(cid) + "/events", { method: "POST", body: { subject: SET_SUBJECT, body: body, start: { dateTime: "2020-01-01T09:00:00", timeZone: TZ }, end: { dateTime: "2020-01-01T09:30:00", timeZone: TZ }, showAs: "free", isReminderOn: false } });
        return req.then(function (ev) {
          var eid = (ev && ev.id) || settingsId;
          items[eid] = { mod: (ev && ev.lastModifiedDateTime) || "", kind: "settings", data: obj };
          reindex(); emit(); banner();
        });
      });
    });
  }
  function fail(code) { var e = new Error(code); e.code = code; return Promise.reject(e); }
  function ready() { return state === "ok" ? Promise.resolve() : fail("unavailable"); }

  /* ---- the storage, download and user interfaces the Quote Builder expects ---- */
  var db = {
    doc: function (path) {
      var parts = String(path).split("/"), col = parts[0], id = parts[1];
      return {
        id: id, path: path,
        set: function (data) {
          return ready().then(function () {
            var d = JSON.parse(JSON.stringify(data));
            if (col === "settings") return saveSettings(d);
            if (col === "quotes") { d.no = d.no || id; return saveQuote(d); }
            return fail("invalid_argument");
          });
        },
        update: function (patch) {
          return ready().then(function () {
            if (col === "settings") return saveSettings(Object.assign({}, settingsData || {}, patch));
            if (col === "quotes") { if (!byNo[id]) return fail("invalid_argument"); return saveQuote(Object.assign({}, items[byNo[id]].data, patch)); }
            return fail("invalid_argument");
          });
        },
        onSnapshot: function (next, err) {
          var s = { next: next, err: err }; if (col === "settings") sSubs.push(s);
          if (state === "ok" && col === "settings") { try { next({ id: "main", exists: !!settingsData, data: function () { return settingsData || undefined; } }); } catch (x) {} }
          return function () { sSubs = sSubs.filter(function (x) { return x !== s; }); };
        }
      };
    },
    collection: function () {
      var q = {
        orderBy: function () { return q; }, limit: function () { return q; }, where: function () { return q; },
        onSnapshot: function (next, err) {
          var s = { next: next, err: err }; qSubs.push(s);
          if (state === "ok") emit();
          return function () { qSubs = qSubs.filter(function (x) { return x !== s; }); };
        }
      };
      return q;
    }
  };
  var downloads = {
    save: function (o) {
      var blob = o.data instanceof Blob ? o.data : new Blob([o.data]);
      var url = URL.createObjectURL(blob), a = document.createElement("a");
      a.href = url; a.download = o.filename || "download"; a.style.display = "none";
      document.body.appendChild(a); a.click();
      setTimeout(function () { a.remove(); URL.revokeObjectURL(url); }, 60000);
      return Promise.resolve();
    }
  };
  var user = { canEdit: function () { return Promise.resolve(true); }, id: function () { return Promise.resolve(E.who || null); } };
  window.claude = { use: function (n) { return Promise.resolve(n === "db" ? db : n === "downloads" ? downloads : n === "user" ? user : null); } };

  window.EAM_BRIDGE = {
    createBooking: function (ev) {
      return E.gfetch(BASE + "/events", { method: "POST", body: {
        subject: ev.subject, body: { contentType: "text", content: ev.body },
        start: { dateTime: String(ev.start).slice(0, 19), timeZone: TZ }, end: { dateTime: String(ev.end).slice(0, 19), timeZone: TZ },
        location: { displayName: ev.location || "" }
      } }).then(function (r) { try { E.refreshBoard(); } catch (x) {} return r; });
    }
  };

  /* ---- banner: connection problems and first-time settings ---- */
  function banner() {
    var el = document.getElementById("eamBridgeBar");
    var need = state === "error" ? "error" : (state === "ok" && !settingsData ? "settings" : "");
    if (!need) { if (el) el.remove(); return; }
    if (!document.body) return;
    if (!el) {
      el = document.createElement("div"); el.id = "eamBridgeBar";
      el.style.cssText = "position:sticky;top:0;z-index:9999;background:#FFF4DD;color:#5a3a00;border-bottom:1px solid #B45309;padding:10px 14px;font:600 14px/1.4 system-ui,sans-serif;display:flex;flex-wrap:wrap;gap:10px;align-items:center";
      document.body.insertBefore(el, document.body.firstChild);
    }
    if (el.dataset.kind === need) return;
    el.dataset.kind = need;
    if (need === "error") {
      el.innerHTML = '<span style="flex:1;min-width:220px">Quotes and rates couldn\'t be loaded from Outlook.</span><button type="button" id="eamRetry" style="font:inherit;padding:6px 12px;border-radius:8px;border:1px solid #B45309;background:#fff;cursor:pointer">Try again</button>';
      document.getElementById("eamRetry").onclick = function () { el.remove(); loadAll(); };
    } else {
      el.innerHTML = '<span style="flex:1;min-width:220px">Your rates are not loaded yet. Choose the settings file to load them into Outlook. This is needed once.</span><label for="eamSetFile" style="padding:6px 12px;border-radius:8px;border:1px solid #B45309;background:#fff;cursor:pointer">Choose settings file</label><input type="file" id="eamSetFile" accept=".json,application/json" style="position:absolute;width:1px;height:1px;opacity:0"><span id="eamSetMsg" style="flex-basis:100%;font-weight:500"></span>';
      document.getElementById("eamSetFile").onchange = function (ev) {
        var f = ev.target.files && ev.target.files[0], msg = document.getElementById("eamSetMsg"); if (!f) return;
        var rd = new FileReader();
        rd.onload = function () {
          var obj; try { obj = JSON.parse(String(rd.result)); } catch (x) { msg.textContent = "That file couldn't be read. Choose the eam-quote-settings.json file."; return; }
          if (!obj || !obj.calc || !Array.isArray(obj.calc.crews) || !obj.calc.crews.length) { msg.textContent = "That file doesn't contain rates. Choose the eam-quote-settings.json file."; return; }
          delete obj.sample; msg.textContent = "Saving to Outlook…";
          saveSettings(obj).then(function () {}, function () { msg.textContent = "Couldn't save to Outlook. Check the connection and try again."; });
        };
        rd.readAsText(f);
      };
    }
  }

  /* ---- truck availability beside the move date and start time ---- */
  var availKey = "", availSeq = 0;
  function startMin(t) {
    var s = String(t || "").toLowerCase(), m = s.match(/(\d{1,2})(?::(\d{2}))?\s*(am|pm)?/); if (!m) return 480;
    var hr = +m[1], mn = +(m[2] || 0), ap = m[3] || (s.match(/(am|pm)/) || [])[1] || "am";
    if (hr === 12) return 720 + mn; if (ap === "pm") hr += 12; return hr * 60 + mn;
  }
  function availBox() {
    var el = document.getElementById("eamAvail"); if (el) return el;
    var t = document.getElementById("m-time"), lab = t && t.closest("label"); if (!lab) return null;
    el = document.createElement("div"); el.id = "eamAvail"; el.setAttribute("role", "status");
    el.style.cssText = "grid-column:1/-1;border-radius:9px;padding:9px 12px;font-size:14px;font-weight:600;line-height:1.4;display:none";
    lab.insertAdjacentElement("afterend", el); return el;
  }
  function paint(el, kind, text) {
    var c = kind === "ok" ? ["#E3F5EA", "#1B7F4C"] : kind === "warn" ? ["#FFF4DD", "#8A4B00"] : ["#EEF2F8", "#5E6673"];
    el.style.display = "block"; el.style.background = c[0]; el.style.color = c[1]; el.textContent = text;
  }
  function checkAvail() {
    var d = document.getElementById("m-date"), t = document.getElementById("m-time"), h = document.getElementById("p-hours"), f = document.getElementById("m-flex");
    if (!d || !t) return;
    var el = availBox(); if (!el) return;
    var key = [d.value, t.value, h ? h.value : "", f && f.checked ? 1 : 0].join("|");
    if (key === availKey) return; availKey = key;
    if (!d.value || (f && f.checked) || !E.availability) { el.style.display = "none"; return; }
    var st = startMin(t.value), hours = parseFloat(h && h.value) > 0 ? parseFloat(h.value) : 2, seq = ++availSeq;
    paint(el, "", "Checking the trucks for that time…");
    E.availability(d.value, st, st + Math.round(hours * 60)).then(function (r) {
      if (seq !== availSeq) return;
      var free = r.trucks.filter(function (x) { return x.free; }), busy = r.trucks.filter(function (x) { return !x.free; });
      var wait = r.waiting ? " " + r.waiting + " booked job" + (r.waiting === 1 ? "" : "s") + " in that window still need" + (r.waiting === 1 ? "s" : "") + " a truck." : "";
      var nm = function (x) { return x.name + (x.size ? " (" + x.size + ")" : ""); };
      if (!r.trucks.length) paint(el, "", "No trucks are set up on the Dispatch page yet, so availability can't be checked.");
      else if (!free.length) paint(el, "warn", "No truck of ours is free then. " + busy.map(function (x) { return x.name + " is on " + x.busy; }).join(". ") + ". This job would need a subcontractor or another time." + wait);
      else if (free.length <= r.waiting) paint(el, "warn", "Free then: " + free.map(nm).join(", ") + ", but" + wait.replace(/^ /, " ") + " Check Dispatch before promising this time.");
      else paint(el, "ok", "Free then: " + free.map(nm).join(", ") + "." + (busy.length ? " " + busy.map(function (x) { return x.name + " is on " + x.busy; }).join(". ") + "." : "") + wait);
    }, function () { if (seq !== availSeq) return; availKey = ""; paint(el, "", "Checking the trucks for that time…"); });
  }
  setInterval(checkAvail, 1500);
  document.addEventListener("change", function () { setTimeout(checkAvail, 50); });

  loadAll();
  setInterval(poll, 120000);
  document.addEventListener("visibilitychange", function () { if (!document.hidden) poll(); });
  document.addEventListener("DOMContentLoaded", banner);
})();
