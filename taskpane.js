/* KAHI 메일 도우미 — Outlook 작업창 (읽기 전용)
 * 열린 메일의 고유번호(internetMessageId)로 'KAHI 메일 AI' 웹 앱에 번역·요약을 조회해 보여 준다.
 * 연결 코드(웹 앱 주소 + 조회 키)는 각자 한 번 붙여 넣으면 Outlook 설정(roamingSettings)에 저장된다.
 * 이 파일에는 비밀 정보가 없다. 메일을 수정·이동·발송하지 않는다.
 */
(function () {
  'use strict';
  var KEY = 'kahiConn';
  var app = document.getElementById('app');
  var state = { conn: null, item: null, data: null, busy: false, retry: 0, timer: null, settings: false };

  function esc(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }
  function hasOffice() { return typeof Office !== 'undefined' && Office.context && Office.context.mailbox; }

  /* ---------- 연결 코드 ---------- */
  function parseCode(code) {
    try {
      var s = String(code || '').trim().replace(/-/g, '+').replace(/_/g, '/');
      while (s.length % 4) s += '=';
      var bin = atob(s), bytes = new Uint8Array(bin.length);
      for (var i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
      var o = JSON.parse(new TextDecoder().decode(bytes));
      if (!o || !/^https:\/\/script\.google\.com\/macros\/s\/[\w-]+\/exec$/.test(o.u) || !o.k) return null;
      return { u: o.u, k: String(o.k), code: String(code).trim() };
    } catch (e) { return null; }
  }
  function loadConn() {
    var code = '';
    try { if (hasOffice() && Office.context.roamingSettings) code = Office.context.roamingSettings.get(KEY) || ''; } catch (e) {}
    if (!code) { try { code = localStorage.getItem(KEY) || ''; } catch (e) {} }
    return code ? parseCode(code) : null;
  }
  function saveConn(code, done) {
    try { localStorage.setItem(KEY, code); } catch (e) {}
    if (hasOffice() && Office.context.roamingSettings) {
      try {
        Office.context.roamingSettings.set(KEY, code);
        Office.context.roamingSettings.saveAsync(function () { done(); });
        return;
      } catch (e) {}
    }
    done();
  }
  function clearConn() {
    try { localStorage.removeItem(KEY); } catch (e) {}
    try { if (hasOffice() && Office.context.roamingSettings) { Office.context.roamingSettings.remove(KEY); Office.context.roamingSettings.saveAsync(function () {}); } } catch (e) {}
  }

  /* ---------- 조회 ---------- */
  function api(action, id) {
    var url = state.conn.u + '?a=' + encodeURIComponent(action) + '&id=' + encodeURIComponent(id || '') + '&k=' + encodeURIComponent(state.conn.k);
    var ctl = typeof AbortController !== 'undefined' ? new AbortController() : null;
    var t = setTimeout(function () { if (ctl) ctl.abort(); }, action === 'process' ? 90000 : 30000);
    return fetch(url, { method: 'GET', redirect: 'follow', cache: 'no-store', signal: ctl ? ctl.signal : undefined })
      .then(function (r) { if (!r.ok) throw new Error('HTTP ' + r.status); return r.text(); })
      .then(function (t2) { try { return JSON.parse(t2); } catch (e) { throw new Error('응답 형식 오류(웹 앱 배포·접근 권한을 확인하세요)'); } })
      .finally(function () { clearTimeout(t); });
  }

  function currentItem() {
    if (!hasOffice()) return null;
    var it = Office.context.mailbox.item;
    if (!it) return null;
    return { id: it.internetMessageId || '', subject: it.subject || '' };
  }

  function load(action) {
    clearTimeout(state.timer);
    state.settings = false;
    state.conn = loadConn();
    if (!state.conn) { renderSetup(); return; }
    state.item = currentItem();
    if (!state.item) { app.innerHTML = '<p class="muted">메일을 하나 선택하면 번역·요약이 여기에 보입니다.</p>'; return; }
    if (!state.item.id) { app.innerHTML = '<div class="notice">이 메일은 아직 고유번호가 없습니다(작성 중인 메일 등). 받은 메일이나 보낸 메일을 열어 주세요.</div>'; return; }
    state.busy = true;
    if (action !== 'silent') app.innerHTML = '<p class="muted"><span class="spin"></span> ' + (action === 'process' ? '지금 번역·요약을 만드는 중… (최대 1분)' : '불러오는 중…') + '</p>';
    var myId = state.item.id;
    api(action === 'process' ? 'process' : 'mail', myId).then(function (d) {
      if (!state.item || state.item.id !== myId) return;
      state.busy = false; state.data = d; render(d);
    }).catch(function (e) {
      state.busy = false;
      app.innerHTML = '<div class="notice"><b class="err">불러오지 못했습니다.</b><span class="muted">' + esc(e && e.message || e) + '</span><div class="row"><button class="btn" type="button" id="bRetry">다시 시도</button></div></div>';
      bind('bRetry', function () { load(); });
    });
  }

  /* ---------- 화면 ---------- */
  function bind(id, fn) { var el = document.getElementById(id); if (el) el.onclick = fn; }

  function renderSetup(msg) {
    state.settings = true;
    var cur = loadConn();
    app.innerHTML =
      '<div class="card"><h3>연결 코드 입력</h3>' +
      '<p class="muted" style="margin:0">관리자에게 받은 <b>연결 코드</b>를 붙여 넣으세요. 한 번 넣으면 이 Outlook 계정에 저장됩니다. 코드는 팀 4명 외에는 공유하지 마세요.</p>' +
      (msg ? '<p class="err" style="margin:0">' + esc(msg) + '</p>' : '') +
      '<textarea id="code" placeholder="연결 코드" spellcheck="false">' + esc(cur ? cur.code : '') + '</textarea>' +
      '<div class="row"><button class="btn primary" type="button" id="bSave">저장하고 연결 확인</button>' +
      (cur ? '<button class="btn" type="button" id="bCancel">취소</button><button class="btn" type="button" id="bClear">연결 지우기</button>' : '') + '</div>' +
      '<p class="meta" style="margin:0">연결 코드는 Apps Script ‘KAHI 메일 AI’의 showConnectionCode 실행 결과입니다.</p></div>';
    bind('bSave', function () {
      var code = document.getElementById('code').value.trim();
      var c = parseCode(code);
      if (!c) { renderSetup('연결 코드 형식이 맞지 않습니다. 앞뒤 공백 없이 전체를 붙여 넣어 주세요.'); return; }
      var btn = document.getElementById('bSave'); btn.disabled = true; btn.innerHTML = '<span class="spin"></span> 확인 중…';
      state.conn = c;
      api('health').then(function (h) {
        if (!h || !h.ok) { renderSetup(h && h.error === 'auth' ? '조회 키가 맞지 않습니다. 최신 연결 코드인지 확인하세요.' : '연결을 확인하지 못했습니다: ' + (h && h.error || '알 수 없음')); return; }
        saveConn(code, function () { load(); });
      }).catch(function (e) { renderSetup('연결을 확인하지 못했습니다: ' + (e && e.message || e)); });
    });
    bind('bCancel', function () { load(); });
    bind('bClear', function () { clearConn(); renderSetup(); });
  }

  function replyBox(m) {
    var lv = String(m.reply || '');
    if (!lv) return '';
    var cls = lv === '회신 필요' ? 'need' : lv === '결정 회신' ? 'decide' : 'none';
    return '<div class="reply ' + cls + '"><b>' + esc(lv) + '</b>' + (m.replyReason ? '<span>' + esc(m.replyReason) + '</span>' : '') + '</div>';
  }
  function markHl(text, hls) {
    var out = esc(text);
    (hls || []).forEach(function (h) {
      h = String(h || '').trim(); if (!h) return;
      var e = esc(h), i = out.indexOf(e);
      if (i >= 0) out = out.slice(0, i) + '<mark>' + e + '</mark>' + out.slice(i + e.length);
    });
    return out;
  }
  function bullets(s) {
    var lines = String(s || '').split(/\n+/).map(function (x) { return x.replace(/^\s*[-•*·]\s*/, '').trim(); }).filter(Boolean);
    return lines.length ? '<ul>' + lines.map(function (x) { return '<li>' + esc(x) + '</li>'; }).join('') + '</ul>' : '';
  }
  function actList(acts) {
    acts = Array.isArray(acts) ? acts.filter(function (a) { return a && a.what; }) : [];
    if (!acts.length) return '';
    return '<ul class="acts">' + acts.map(function (a) {
      var them = a.who === '상대';
      return '<li><span class="w' + (them ? ' them' : '') + '">' + (them ? '상대' : '우리') + '</span><span>' + esc(a.what) + '</span>' + (a.due ? '<span class="due">~' + esc(String(a.due).slice(5).replace('-', '/')) + '</span>' : '') + '</li>';
    }).join('') + '</ul>';
  }
  function creatorCard(d) {
    var c = d.creator, t = d.thread || {}, links = [];
    if (t.doc) links.push('<a href="' + esc(t.doc) + '" target="_blank" rel="noopener">대화 문서</a>');
    if (c && c.folder) links.push('<a href="' + esc(c.folder) + '" target="_blank" rel="noopener">드라이브 폴더</a>');
    if (d.links && d.links.desk) links.push('<a href="' + esc(d.links.desk) + '" target="_blank" rel="noopener">크리에이터 데스크</a>');
    var head = c ? '<div class="who"><b>' + esc(c.name || c.handle || c.id) + '</b>' + (c.handle ? '<span class="muted">@' + esc(String(c.handle).replace(/^@/, '')) + '</span>' : '') + (c.core ? '<span class="chip core">Core</span>' : '') + '<span class="chip">' + esc(c.id) + '</span></div>'
      : '<div class="who"><b>연결된 크리에이터 없음</b><span class="chip">미분류</span></div>';
    var th = t.count ? '<div class="meta">이 대화 ' + t.count + '통' + (t.replied ? ' · 이 메일 뒤 우리 회신 ' + esc(t.repliedAt) : '') + (t.lastDt ? ' · 마지막 ' + esc(t.lastDir) + ' ' + esc(t.lastDt) : '') + '</div>' : '';
    return '<div class="card">' + head + th + (c && c.memo ? '<div class="muted">' + esc(c.memo) + '</div>' : '') + (links.length ? '<div class="links">' + links.join('') + '</div>' : '') + '</div>';
  }
  function history(d) {
    var h = d.history || [];
    if (!h.length) return '';
    return '<details class="card"><summary>앞선 흐름 (' + h.length + ')</summary><ul class="hist">' + h.map(function (x) { return '<li><span class="d">' + esc(String(x.dt).slice(5)) + ' ' + esc(x.dir) + '</span>' + esc(x.line) + '</li>'; }).join('') + '</ul></details>';
  }

  function render(d) {
    if (!d || d.ok === false) {
      if (d && d.error === 'auth') { renderSetup('조회 키가 바뀌었습니다. 새 연결 코드를 넣어 주세요.'); return; }
      app.innerHTML = '<div class="notice"><b class="err">오류</b><span>' + esc(d && d.error || '알 수 없는 오류') + '</span></div>'; return;
    }
    var st = d.status, html = creatorCard(d);
    if (st === 'ready') {
      var m = d.mail || {};
      html += replyBox(m);
      html += '<div class="card"><h3>요약</h3>' + (m.summaryLine ? '<div class="sumline">' + esc(m.summaryLine) + '</div>' : '') + bullets(m.summary) + '</div>';
      var acts = actList(m.actions);
      if (acts || m.deadline) html += '<div class="card"><h3>확인 필요 · 할 일' + (m.deadline ? ' <span class="due">(기한 ' + esc(m.deadline) + ')</span>' : '') + '</h3>' + acts + '</div>';
      if (m.attachments) html += '<div class="card"><h3>첨부 요약</h3><div>' + esc(m.attachments) + '</div></div>';
      if (m.translation) html += '<details class="card" open><summary>국문 번역</summary><div class="tr">' + markHl(m.translation, m.highlights) + '</div></details>';
      html += history(d);
      html += '<p class="meta">' + esc(m.dir || '') + ' ' + esc(m.dt || '') + (m.processedAt ? ' · AI 정리 ' + esc(String(m.processedAt).slice(0, 16)) : '') + '</p>';
    } else if (st === 'pending' || st === 'not_collected' || st === 'busy' || st === 'raw_missing') {
      var msg = st === 'not_collected' ? '아직 드라이브에 수집되지 않은 메일입니다. 보통 10~15분 안에 수집됩니다.'
        : st === 'raw_missing' ? '메일 목록에는 있지만 원본 파일을 아직 찾지 못했습니다. 잠시 뒤 다시 시도해 주세요.'
        : st === 'busy' ? '자동 처리가 진행 중입니다. 잠시 뒤 다시 시도해 주세요.'
        : '수집은 되었고 번역·요약을 기다리는 중입니다(5분마다 자동 처리' + (d.run && d.run.pending ? ' · 대기 ' + d.run.pending + '통' : '') + ').';
      html += '<div class="notice"><span>' + esc(msg) + '</span><div class="row">' +
        (st === 'pending' || st === 'raw_missing' ? '<button class="btn primary" type="button" id="bNow">지금 번역·요약 만들기</button>' : '') +
        '<button class="btn" type="button" id="bAgain">다시 확인</button></div><span class="meta">이 창을 열어 두면 1분마다 다시 확인합니다.</span></div>';
      html += history(d);
      scheduleRetry();
    } else if (st === 'skipped') {
      html += '<div class="notice">번역 대상이 아닌 메일입니다' + (d.reason ? ' (' + esc(d.reason) + ')' : '') + '.</div>';
    } else if (st === 'error') {
      html += '<div class="notice"><b class="err">처리 중 오류</b><span class="muted">' + esc(d.reason || '') + '</span><div class="row"><button class="btn" type="button" id="bAgain">다시 확인</button></div></div>';
    } else {
      html += '<div class="notice">' + esc(st || '알 수 없는 상태') + '</div>';
    }
    app.innerHTML = html;
    bind('bNow', function () { load('process'); });
    bind('bAgain', function () { load(); });
  }

  function scheduleRetry() {
    clearTimeout(state.timer);
    if (state.retry >= 10) return;
    state.timer = setTimeout(function () { state.retry++; if (!state.settings) load('silent'); }, 60000);
  }

  /* ---------- 시작 ---------- */
  document.getElementById('btnRefresh').onclick = function () { state.retry = 0; load(); };
  document.getElementById('btnSettings').onclick = function () { if (state.settings && loadConn()) load(); else renderSetup(); };

  function start() {
    if (hasOffice() && Office.context.mailbox.addHandlerAsync && Office.EventType && Office.EventType.ItemChanged) {
      try { Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, function () { state.retry = 0; load(); }); } catch (e) {}
    }
    load();
  }
  if (typeof Office !== 'undefined' && Office.onReady) {
    Office.onReady(function () {
      if (!hasOffice()) { app.innerHTML = '<div class="notice">이 화면은 Outlook 안에서 열어야 동작합니다. 설치 방법은 <a href="help.html">사용 안내</a>를 보세요.</div>'; return; }
      start();
    });
  } else {
    app.innerHTML = '<div class="notice">Office 라이브러리를 불러오지 못했습니다. Outlook 안에서 다시 열어 주세요.</div>';
  }
})();
