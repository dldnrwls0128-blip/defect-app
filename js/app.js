/** 하자관리 PWA — 화면·라우팅 */
import * as db from "./db.js";
import {
  PRESETS, LOCAL_STATUS, statusOf, parseQuery, applyFilter, groupByTradeVendor, countBy,
  hoLabel, shortDong, uniq, sortUnit, isMi, isPa, normHo, normDong, floorOfHo, migrateRows,
} from "./data.js";
import { importFile } from "./importer.js";
import { compressImage, downloadBlob, shareOrDownload } from "./photo.js";
import { exportBackup, importBackup } from "./backup.js";
import { toXlsx, toCsv, unitText, vendorText, shareText } from "./exporter.js";

const APP_VERSION = "1.0.1";
const $ = (s, el = document) => el.querySelector(s);
const $$ = (s, el = document) => [...el.querySelectorAll(s)];
const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const n = (x) => Number(x || 0).toLocaleString("ko-KR");
const todayStr = () => {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
};

const LS = {
  get(k, d) { try { return JSON.parse(localStorage.getItem("dfx:" + k)) ?? d; } catch { return d; } },
  set(k, v) { localStorage.setItem("dfx:" + k, JSON.stringify(v)); },
};

const S = {
  rows: [], byId: new Map(), edits: new Map(), pcount: new Map(), imports: [],
  filter: LS.get("filter", { preset: "mipa" }),
  unitPreset: LS.get("unitPreset", "mipa"),
  unitOpen: LS.get("unitOpen", false),
  dash: LS.get("dash", { dong: "", preset: "mipa" }),
  vend: LS.get("vend", { vendor: "", dong: "", preset: "mipa", open: true }),
  nav: [], // 상세 화면 이전/다음용 목록
};
const st = (r) => statusOf(r, S.edits.get(r.id));
const ctx = { edit: (id) => S.edits.get(id), status: st, photos: (id) => S.pcount.get(id) || 0 };

/* ---------- 공통 UI ---------- */
function toast(msg, ms = 2200) {
  const t = $("#toast");
  t.textContent = msg;
  t.classList.add("show");
  clearTimeout(toast._t);
  toast._t = setTimeout(() => t.classList.remove("show"), ms);
}
function busy(on, text = "처리 중…") {
  $("#busy").hidden = !on;
  $("#busy-text").textContent = text;
}
function setHeader(title, back = false) {
  $("#header-title").textContent = title;
  $("#btn-back").hidden = !back;
}
function setTab(tab) {
  $$(".tab-item").forEach((a) => a.classList.toggle("active", a.dataset.tab === tab));
}
const view = () => $("#view");
const STATUS_CLS = { 미처리: "s0", 작업중: "s1", 완료: "s2", 재접수: "s3" };
const pill = (s, btn = false) => `<span class="pill ${STATUS_CLS[s]}"${btn ? ' data-cycle="1" role="button" tabindex="0" aria-label="상태 변경"' : ""}>${s}</span>`;
const typeBadge = (r) => `<span class="tbadge ${isMi(r) ? "t-mi" : isPa(r) ? "t-pa" : ""}">${esc(r.type || "-")}</span>`;

function presetChips(cur, attr = "preset") {
  return `<div class="chips">${Object.entries(PRESETS)
    .map(([k, p]) => `<button type="button" class="chip ${cur === k ? "on" : ""}" data-${attr}="${k}">${p.label}</button>`)
    .join("")}</div>`;
}

function noData() {
  return `<div class="card empty-card">
    <div class="big-ico">📥</div>
    <h2>하자 엑셀을 먼저 가져오세요</h2>
    <p class="help">DTspace 하자관리 → 조회 → <b>Export Excel</b>로 받은 파일(.xlsx)을 선택하면 폰 안에만 저장됩니다. 인터넷·서버로 보내지 않습니다.</p>
    <button type="button" class="btn btn-primary btn-lg" data-act="import">엑셀 가져오기</button>
  </div>`;
}

function defectCard(r) {
  const s = st(r);
  const e = S.edits.get(r.id);
  const pc = S.pcount.get(r.id) || 0;
  return `<div class="defect ${STATUS_CLS[s]}" data-id="${r.id}">
    <div class="d-top">${pill(s, true)}${typeBadge(r)}<span class="d-loc">${esc(r.loc || "")}${r.sub ? " · " + esc(r.sub) : ""}</span></div>
    <div class="d-text">${esc(r.text || "(내용 없음)")}</div>
    <div class="d-meta">${r.dong}동 ${hoLabel(r.ho)}${r.part ? " · " + esc(r.part) : ""} · #${r.id}${pc ? ` · 📷${pc}` : ""}${e?.memo ? " · 📝" : ""}${e?.doneDate ? " · 완료 " + esc(e.doneDate.slice(5)) : ""}${r._gone ? ' · <b class="warn">최신 엑셀에 없음</b>' : ""}</div>
  </div>`;
}

function groupsHtml(rows, { openAll = true, showUnit = false } = {}) {
  const groups = groupByTradeVendor(rows);
  if (!groups.length) return `<div class="empty">해당하는 하자가 없습니다.</div>`;
  return groups.map((g, i) => {
    const open = g.rows.filter((r) => st(r) !== "완료").length;
    return `<details class="grp" data-g="${i}" ${openAll ? "open" : ""}>
      <summary><span class="g-name"><b>${esc(g.trade)}</b><small>${esc(g.vendor)}</small></span>
      <span class="g-cnt">${open < g.rows.length ? `<em>잔여 ${open}</em>` : ""}<b>${g.rows.length}</b></span></summary>
      <div class="g-body">${openAll ? g.rows.sort(sortUnit).map(defectCard).join("") : ""}</div>
    </details>`;
  }).join("");
}
function bindGroups(rows) {
  const groups = groupByTradeVendor(rows);
  $$("details.grp").forEach((d) => d.addEventListener("toggle", () => {
    const body = $(".g-body", d);
    if (d.open && !body.children.length) body.innerHTML = groups[+d.dataset.g].rows.sort(sortUnit).map(defectCard).join("");
  }));
  S.nav = groups.flatMap((g) => g.rows.sort(sortUnit).map((r) => r.id));
}

/* 카드 클릭: 상태 배지 = 상태 순환, 나머지 = 상세 */
document.addEventListener("click", async (ev) => {
  const cyc = ev.target.closest("[data-cycle]");
  const card = ev.target.closest(".defect[data-id]");
  if (cyc && card) {
    ev.preventDefault();
    const id = +card.dataset.id;
    const r = S.byId.get(id);
    const order = ["미처리", "작업중", "완료"];
    const cur = st(r);
    const next = order[(order.indexOf(cur) + 1) % order.length] || "작업중";
    await saveEdit(id, { status: next, ...(next === "완료" && !S.edits.get(id)?.doneDate ? { doneDate: todayStr() } : {}) });
    card.outerHTML = defectCard(r);
    toast(`#${id} → ${next}`);
    return;
  }
  if (card) {
    location.hash = `#/item/${card.dataset.id}`;
    return;
  }
  const act = ev.target.closest("[data-act]")?.dataset.act;
  if (act === "import") $("#input-xlsx").click();
});

async function saveEdit(id, patch) {
  const cur = S.edits.get(id) || { id };
  const next = { ...cur, ...patch, id, updatedAt: new Date().toISOString() };
  S.edits.set(id, next);
  await db.put("edits", next);
  return next;
}

/* ---------- 홈 ---------- */
function recentUnits() { return LS.get("recent", []); }
function pushRecent(dong, ho) {
  const k = `${dong}|${ho}`;
  LS.set("recent", [k, ...recentUnits().filter((x) => x !== k)].slice(0, 8));
}

function renderHome() {
  setHeader("하자관리");
  setTab("home");
  if (!S.rows.length) { view().innerHTML = noData(); return; }
  const mp = S.rows.filter(PRESETS.mipa.test);
  const done = S.rows.filter((r) => st(r) === "완료").length;
  const work = S.rows.filter((r) => st(r) === "작업중").length;
  const dongs = uniq(S.rows, "dong");
  const last = S.imports[0];
  view().innerHTML = `
    <form class="card search-card" id="f-search">
      <label class="lab" for="q">동·호수 검색</label>
      <div class="inline"><input id="q" inputmode="text" autocomplete="off" placeholder="예) 9동 1203 · 12 203 · 309-1203" />
      <button class="btn btn-primary" type="submit">찾기</button></div>
      <p class="help">9동=309동, 12동=312동. 동만 입력하면 세대 목록, 글자는 하자내용 검색.</p>
    </form>
    <div class="chips dong-chips">${dongs.map((d) => `<a class="chip" href="#/dong/${d}">${shortDong(d)}</a>`).join("")}</div>
    ${recentUnits().length ? `<div class="card"><h3 class="block-title">최근 본 세대</h3><div class="chip-row">${recentUnits()
      .map((k) => { const [d, h] = k.split("|"); return `<a class="chip" href="#/unit/${d}/${h}">${shortDong(d)} ${hoLabel(h)}</a>`; }).join("")}</div></div>` : ""}
    <div class="card stats">
      <div class="stat"><b>${n(S.rows.length)}</b><small>전체</small></div>
      <div class="stat hot"><b>${n(mp.length)}</b><small>미시공·파손</small></div>
      <div class="stat"><b>${n(work)}</b><small>작업중</small></div>
      <div class="stat ok"><b>${n(done)}</b><small>완료</small></div>
    </div>
    <p class="help center">데이터: ${esc(last?.fileName || "-")} · 최근 접수일 ${esc(S.maxRcv || "-")} · 가져온 날 ${last ? esc(last.importedAt.slice(0, 10)) : "-"}</p>`;
  $("#f-search").addEventListener("submit", (e) => {
    e.preventDefault();
    const q = $("#q").value;
    const p = parseQuery(q);
    if (!p) return;
    if (p.dong && p.ho) location.hash = `#/unit/${p.dong}/${p.ho}`;
    else if (p.dong) location.hash = `#/dong/${p.dong}`;
    else {
      S.filter = { preset: "all", ...(p.ho ? { ho: p.ho } : { text: p.text }) };
      LS.set("filter", S.filter);
      location.hash = "#/list";
    }
  });
}

/* ---------- 동: 세대 그리드 ---------- */
function renderDong(dong) {
  setHeader(`${dong}동 (${shortDong(dong)})`, true);
  setTab("home");
  if (!S.rows.length) { view().innerHTML = noData(); return; }
  const pre = PRESETS[S.unitPreset] || PRESETS.mipa;
  const rows = S.rows.filter((r) => r.dong === dong);
  const units = new Map();
  for (const r of rows) {
    if (!units.has(r.ho)) units.set(r.ho, { all: 0, hit: 0, open: 0 });
    const u = units.get(r.ho);
    u.all++;
    if (pre.test(r)) { u.hit++; if (st(r) !== "완료") u.open++; }
  }
  // 층·세대는 실제 데이터에서만 만든다(없는 층/세대는 생기지 않음). 층 = 호수 ÷ 100 (0901 → 9층)
  const hos = [...units.keys()].sort((a, b) => (+a || 0) - (+b || 0) || a.localeCompare(b));
  const floors = new Map();
  for (const h of hos) {
    const fl = floorOfHo(h);
    const f = Number.isFinite(fl) ? fl : -1; // -1 = 층을 알 수 없는 호수(기타)
    if (!floors.has(f)) floors.set(f, []);
    floors.get(f).push(h);
  }
  const hit = [...units.values()].reduce((a, u) => a + u.hit, 0);
  const open = [...units.values()].reduce((a, u) => a + u.open, 0);
  view().innerHTML = `
    ${presetChips(S.unitPreset)}
    <div class="card" id="dong-sum" data-hit="${hit}" data-units="${hos.length}" data-floors="${[...floors.keys()].filter((f) => f >= 0).length}"><b>${dong}동</b> ${[...floors.keys()].filter((f) => f >= 0).length}개 층 · 세대 ${hos.length} · ${pre.label} ${n(hit)}건 · 잔여 <b class="hot-t">${n(open)}</b>
      <div class="row-actions"><a class="btn" href="#/list" id="dong-list">이 동 공종별 목록</a></div></div>
    <div class="ho-grid">${[...floors.entries()].sort((a, b) => b[0] - a[0]).map(([f, list]) => `
      <div class="floor" data-floor="${f}"><span class="fl">${f < 0 ? "기타" : f + "F"}</span>${list.map((h) => {
        const u = units.get(h);
        const lvl = u.open === 0 ? (u.hit ? "z" : "n") : u.open >= 10 ? "h" : u.open >= 4 ? "m" : "l";
        return `<a class="ho ${lvl}" data-ho="${esc(h)}" data-hit="${u.hit}" href="#/unit/${dong}/${encodeURIComponent(h)}"><b>${esc(/^\d+$/.test(h) ? +h : h)}</b><small>${u.hit ? `${u.open}/${u.hit}` : "-"}</small></a>`;
      }).join("")}</div>`).join("")}</div>
    <p class="help">칸의 숫자 = 잔여/${pre.label} 건수. 진한 색일수록 잔여가 많습니다.</p>`;
  bindPreset((k) => { S.unitPreset = k; LS.set("unitPreset", k); renderDong(dong); });
  $("#dong-list").addEventListener("click", () => {
    S.filter = { preset: S.unitPreset, dong };
    LS.set("filter", S.filter);
  });
}

function bindPreset(fn, attr = "preset") {
  $$(`[data-${attr}]`).forEach((b) => b.addEventListener("click", () => fn(b.dataset[attr])));
}

/* ---------- 세대 ---------- */
function renderUnit(dong, ho) {
  setHeader(`${shortDong(dong)} ${hoLabel(ho)}`, true);
  setTab("home");
  if (!S.rows.length) { view().innerHTML = noData(); return; }
  const all = S.rows.filter((r) => r.dong === dong && r.ho === ho);
  if (!all.length) {
    view().innerHTML = `<div class="card empty">${dong}동 ${hoLabel(ho)} 하자가 없습니다.<br><small>호수를 다시 확인하세요.</small></div>`;
    return;
  }
  pushRecent(dong, ho);
  const pre = PRESETS[S.unitPreset] || PRESETS.mipa;
  let rows = all.filter(pre.test);
  const hitN = rows.length;
  if (S.unitOpen) rows = rows.filter((r) => st(r) !== "완료");
  const doneN = all.filter(pre.test).filter((r) => st(r) === "완료").length;
  const hos = uniq(S.rows.filter((r) => r.dong === dong), "ho");
  const i = hos.indexOf(ho);
  const prev = hos[i - 1], next = hos[i + 1];
  view().innerHTML = `
    <div class="card unit-head">
      <div class="u-title">${dong}동 ${hoLabel(ho)} <small>${all[0].py ? esc(all[0].py) + "타입" : ""}</small></div>
      <div class="u-sub">전체 ${all.length}건 · ${pre.label} ${hitN}건 · 완료 ${doneN} · 잔여 <b class="hot-t">${hitN - doneN}</b></div>
      <div class="u-nav">${prev ? `<a class="btn" href="#/unit/${dong}/${prev}">‹ ${+prev}호</a>` : "<span></span>"}
        ${next ? `<a class="btn" href="#/unit/${dong}/${next}">${+next}호 ›</a>` : "<span></span>"}</div>
    </div>
    ${presetChips(S.unitPreset)}
    <label class="toggle-row"><input type="checkbox" id="only-open" ${S.unitOpen ? "checked" : ""}/> 완료 숨기기</label>
    <div class="row-actions">
      <button type="button" class="btn btn-kakao" id="u-share">카톡 공유</button>
      <button type="button" class="btn" id="u-xlsx">엑셀</button>
    </div>
    ${groupsHtml(rows)}
    <p class="help">상태 배지를 누르면 미처리 → 작업중 → 완료로 바뀝니다. 항목을 누르면 사진·메모.</p>`;
  bindGroups(rows);
  bindPreset((k) => { S.unitPreset = k; LS.set("unitPreset", k); renderUnit(dong, ho); });
  $("#only-open").addEventListener("change", (e) => { S.unitOpen = e.target.checked; LS.set("unitOpen", S.unitOpen); renderUnit(dong, ho); });
  $("#u-share").addEventListener("click", async () => {
    const r = await shareText(unitText(dong, ho, rows, ctx, pre.label), `${dong}동 ${ho}호`);
    if (r === "copied") toast("복사했습니다. 카톡에 붙여넣으세요.");
  });
  $("#u-xlsx").addEventListener("click", async () => {
    const blob = await toXlsx(rows, ctx, `${dong}-${ho}`);
    await shareOrDownload(blob, `${dong}동${+ho}호_${pre.label.replace("·", "")}_${todayStr().replace(/-/g, "").slice(2)}.xlsx`);
  });
}

/* ---------- 조회(필터) ---------- */
function opt(list, cur, all = "전체", label = (x) => x) {
  return `<option value="">${all}</option>` + list.map((v) => `<option value="${esc(v)}" ${v === cur ? "selected" : ""}>${esc(label(v))}</option>`).join("");
}
function renderList() {
  setHeader("조회");
  setTab("list");
  if (!S.rows.length) { view().innerHTML = noData(); return; }
  const f = S.filter;
  const scope = S.rows.filter((r) => (!f.dong || r.dong === f.dong));
  const res = applyFilter(S.rows, f, st);
  const open = res.length <= 150;
  view().innerHTML = `
    ${presetChips(f.preset || "all")}
    <div class="card filters">
      <div class="grid2">
        <label>동<select data-f="dong">${opt(uniq(S.rows, "dong"), f.dong, "전체", (d) => `${d} (${shortDong(d)})`)}</select></label>
        <label>호수<input data-f="ho" inputmode="numeric" value="${esc(f.ho ? +f.ho : "")}" placeholder="예) 1203"/></label>
        <label>유형<select data-f="type">${opt(uniq(scope, "type"), f.type)}</select></label>
        <label>내 상태<select data-f="st"><option value="">전체</option><option value="open" ${f.st === "open" ? "selected" : ""}>미완료(완료 제외)</option>${LOCAL_STATUS.map((s) => `<option ${f.st === s ? "selected" : ""}>${s}</option>`).join("")}</select></label>
        <label>공종<select data-f="trade">${opt(uniq(scope, "trade"), f.trade)}</select></label>
        <label>업체<select data-f="vendor">${opt(uniq(scope, "vendor"), f.vendor)}</select></label>
        <label>DTspace 처리상태<select data-f="dts">${opt(uniq(S.rows, "dts"), f.dts)}</select></label>
        <label>내용 검색<input data-f="text" value="${esc(f.text || "")}" placeholder="예) 코킹"/></label>
      </div>
      <div class="row-actions"><button type="button" class="btn" id="f-reset">초기화</button>
        <button type="button" class="btn" id="l-xlsx">엑셀</button><button type="button" class="btn" id="l-csv">CSV</button></div>
    </div>
    <div class="result-bar"><b>${n(res.length)}</b>건 · 잔여 ${n(res.filter((r) => st(r) !== "완료").length)} ${open ? "" : '<small>(그룹을 눌러 펼치기)</small>'}</div>
    ${groupsHtml(res, { openAll: open })}`;
  bindGroups(res);
  bindPreset((k) => { S.filter = { ...S.filter, preset: k }; LS.set("filter", S.filter); renderList(); });
  $$("[data-f]").forEach((el) => el.addEventListener("change", () => {
    let v = el.value.trim();
    if (el.dataset.f === "ho" && v) v = normHo(v);
    S.filter = { ...S.filter, [el.dataset.f]: v };
    if (el.dataset.f === "dong") { delete S.filter.trade; delete S.filter.vendor; delete S.filter.type; }
    LS.set("filter", S.filter);
    renderList();
  }));
  $("#f-reset").addEventListener("click", () => { S.filter = { preset: "mipa" }; LS.set("filter", S.filter); renderList(); });
  const fname = () => `하자조회_${[f.dong, f.ho && +f.ho, PRESETS[f.preset]?.label.replace("·", ""), f.vendor].filter(Boolean).join("_")}_${todayStr().replace(/-/g, "").slice(2)}`;
  $("#l-xlsx").addEventListener("click", async () => { busy(true, "엑셀 만드는 중…"); try { await shareOrDownload(await toXlsx(res, ctx), fname() + ".xlsx"); } finally { busy(false); } });
  $("#l-csv").addEventListener("click", () => downloadBlob(toCsv(res, ctx), fname() + ".csv"));
}

/* ---------- 상세 ---------- */
async function renderItem(id) {
  const r = S.byId.get(id);
  setTab("");
  if (!r) { setHeader("하자", true); view().innerHTML = `<div class="empty">#${id} 항목이 없습니다.</div>`; return; }
  setHeader(`${shortDong(r.dong)} ${hoLabel(r.ho)} · #${id}`, true);
  const e = S.edits.get(id) || {};
  const s = st(r);
  const photos = await db.photosOf(id);
  const idx = S.nav.indexOf(id);
  const prev = idx > 0 ? S.nav[idx - 1] : null, next = idx >= 0 && idx < S.nav.length - 1 ? S.nav[idx + 1] : null;
  const ph = (kind) => photos.filter((p) => p.kind === kind).sort((a, b) => a.createdAt.localeCompare(b.createdAt));
  const thumbs = (kind) => ph(kind).map((p) => `<div class="pth"><img data-pid="${p.id}" alt="${kind}" src="${URL.createObjectURL(p.thumb || p.blob)}"/><button type="button" class="pdel" data-del="${p.id}" aria-label="삭제">×</button></div>`).join("");
  view().innerHTML = `
    <div class="card item-head">
      <div class="d-top">${typeBadge(r)}<span class="d-loc">${esc(r.trade || "")} · ${esc(r.vendor || "")}</span></div>
      <div class="item-text">${esc(r.text || "(내용 없음)")}</div>
      <div class="muted">${esc(r.loc || "")} / ${esc(r.part || "")} / ${esc(r.sub || "")}</div>
      ${r._gone ? `<p class="warn">최신 DTspace 엑셀에 이 순번이 없습니다(삭제·이관 가능성).</p>` : ""}
    </div>
    <div class="card">
      <h3 class="block-title">내 처리 상태</h3>
      <div class="seg4">${LOCAL_STATUS.map((x) => `<button type="button" class="seg ${x === s ? "on " + STATUS_CLS[x] : ""}" data-st="${x}">${x}</button>`).join("")}</div>
      <label class="lab">완료일<input type="date" id="done-date" value="${esc(e.doneDate || "")}"/></label>
      <label class="lab">메모<textarea id="memo" placeholder="조치 내용, 업체 통화, 자재 입고 등">${esc(e.memo || "")}</textarea></label>
      ${e.updatedAt ? `<p class="help">마지막 수정 ${esc(new Date(e.updatedAt).toLocaleString("ko-KR"))}</p>` : ""}
    </div>
    ${["조치전", "조치후"].map((k) => `<div class="card">
      <h3 class="block-title">${k} 사진 <small class="muted">${ph(k).length}장</small></h3>
      <div class="pgrid">${thumbs(k)}</div>
      <div class="row-actions"><button type="button" class="btn btn-primary" data-cam="${k}">📷 촬영</button><button type="button" class="btn" data-alb="${k}">앨범</button></div>
    </div>`).join("")}
    <div class="card"><h3 class="block-title">DTspace 원본</h3>
      <table class="kv">${[["순번", r.id], ["단계", r.stage], ["접수일", r.rcv], ["평형", r.py], ["대공종", r.big], ["공종", r.trade], ["세부공종", r.sub], ["유형", r.type], ["시공업체", r.vendor],
        ["처리상태", r.dts + (r._prevDts !== undefined && r._prevDts !== r.dts ? ` (이전: ${r._prevDts || "-"})` : "")], ["처리경과일", r.days], ["이미지", r.img], ["작업완료일", r.wdone], ["최종완료일", r.fdone], ["재접수일", r.rercv], ["E/V", r.ev], ["AS예약", r.as], ["참고", r.note]]
        .filter(([, v]) => v !== undefined && v !== "").map(([k, v]) => `<tr><th>${k}</th><td>${esc(v)}</td></tr>`).join("")}</table></div>
    <div class="u-nav sticky-nav">${prev ? `<a class="btn" href="#/item/${prev}">‹ 이전</a>` : "<span></span>"}
      <a class="btn" href="#/unit/${r.dong}/${r.ho}">세대 목록</a>
      ${next ? `<a class="btn" href="#/item/${next}">다음 ›</a>` : "<span></span>"}</div>`;

  $$("[data-st]").forEach((b) => b.addEventListener("click", async () => {
    const v = b.dataset.st;
    const patch = { status: v };
    if (v === "완료" && !$("#done-date").value) { patch.doneDate = todayStr(); }
    await saveEdit(id, patch);
    toast(`상태: ${v}`);
    renderItem(id);
  }));
  $("#done-date").addEventListener("change", (ev) => saveEdit(id, { doneDate: ev.target.value }).then(() => toast("완료일 저장")));
  let mt;
  $("#memo").addEventListener("input", (ev) => {
    clearTimeout(mt);
    mt = setTimeout(() => saveEdit(id, { memo: ev.target.value }), 500);
  });
  $("#memo").addEventListener("blur", (ev) => { clearTimeout(mt); saveEdit(id, { memo: ev.target.value }); });
  const pick = (input, kind) => { S.photoTarget = { id, kind }; input.value = ""; input.click(); };
  $$("[data-cam]").forEach((b) => b.addEventListener("click", () => pick($("#input-camera"), b.dataset.cam)));
  $$("[data-alb]").forEach((b) => b.addEventListener("click", () => pick($("#input-album"), b.dataset.alb)));
  $$("[data-del]").forEach((b) => b.addEventListener("click", async () => {
    if (!confirm("이 사진을 삭제할까요?")) return;
    await db.del("photos", b.dataset.del);
    S.pcount.set(id, Math.max(0, (S.pcount.get(id) || 1) - 1));
    renderItem(id);
  }));
  $$("img[data-pid]").forEach((img) => img.addEventListener("click", () => {
    const p = photos.find((x) => x.id === img.dataset.pid);
    openViewer(p.blob);
  }));
}

function openViewer(blob) {
  const url = URL.createObjectURL(blob);
  const d = document.createElement("div");
  d.className = "viewer";
  d.innerHTML = `<img src="${url}" alt="사진"/><button type="button" class="btn">닫기</button>`;
  d.addEventListener("click", () => { URL.revokeObjectURL(url); d.remove(); });
  document.body.appendChild(d);
}

async function onPhotoFiles(files) {
  const t = S.photoTarget;
  if (!t || !files?.length) return;
  const r = S.byId.get(t.id);
  busy(true, "사진 저장 중…");
  try {
    for (const f of files) {
      const d = new Date();
      const when = `${todayStr()} ${String(d.getHours()).padStart(2, "0")}:${String(d.getMinutes()).padStart(2, "0")}`;
      const stamp = `${r.dong}동 ${hoLabel(r.ho)} ${r.loc || ""} ${r.sub || ""} #${r.id}\n${t.kind} · ${when}`;
      const { blob, thumb, w, h } = await compressImage(f, { stamp });
      await db.put("photos", { id: crypto.randomUUID(), defectId: t.id, kind: t.kind, blob, thumb, w, h, size: blob.size, createdAt: new Date().toISOString() });
      S.pcount.set(t.id, (S.pcount.get(t.id) || 0) + 1);
    }
    if (t.kind === "조치후" && st(r) === "미처리") await saveEdit(t.id, { status: "작업중" });
    toast(`${t.kind} 사진 ${files.length}장 저장`);
  } catch (e) {
    toast("사진 저장 실패: " + e.message, 4000);
  } finally {
    busy(false);
    if (location.hash === `#/item/${t.id}`) renderItem(t.id);
  }
}

/* ---------- 현황 ---------- */
function barTable(title, map, rowsFor, key, limit = 0) {
  const ent = [...map.entries()].sort((a, b) => b[1].all - a[1].all);
  const max = Math.max(1, ...ent.map(([, v]) => v.all));
  const shown = limit ? ent.slice(0, limit) : ent;
  return `<div class="card"><h3 class="block-title">${title}</h3><div class="bars">${shown.map(([k, v]) => `
    <button type="button" class="bar-row" data-k="${esc(key)}" data-v="${esc(k)}">
      <span class="b-lab">${esc(key === "dong" ? `${k} (${shortDong(k)})` : k)}</span>
      <span class="b-bar"><i style="width:${(v.all / max) * 100}%"></i><i class="done" style="width:${(v.done / max) * 100}%"></i></span>
      <span class="b-n">${n(v.all)}<small>${v.all - v.done ? ` 잔${n(v.all - v.done)}` : " ✔"}</small></span>
    </button>`).join("")}</div></div>`;
}
function agg(rows, fn) {
  const m = new Map();
  for (const r of rows) {
    const k = fn(r) || "(없음)";
    if (!m.has(k)) m.set(k, { all: 0, done: 0 });
    const v = m.get(k);
    v.all++;
    if (st(r) === "완료") v.done++;
  }
  return m;
}
function renderDash() {
  setHeader("현황");
  setTab("dash");
  if (!S.rows.length) { view().innerHTML = noData(); return; }
  const D = S.dash;
  const rows = S.rows.filter((r) => (!D.dong || r.dong === D.dong) && (PRESETS[D.preset] || PRESETS.all).test(r));
  const sc = countBy(rows, st);
  const done = sc.get("완료") || 0;
  view().innerHTML = `
    ${presetChips(D.preset)}
    <label class="lab">범위<select id="d-dong">${opt(uniq(S.rows, "dong"), D.dong, "전체 동", (d) => `${d} (${shortDong(d)})`)}</select></label>
    <div class="card stats">
      <div class="stat"><b>${n(rows.length)}</b><small>건수</small></div>
      <div class="stat hot"><b>${n(sc.get("미처리") || 0)}</b><small>미처리</small></div>
      <div class="stat"><b>${n(sc.get("작업중") || 0)}</b><small>작업중</small></div>
      <div class="stat ok"><b>${n(done)}</b><small>완료 ${rows.length ? Math.round((done / rows.length) * 100) : 0}%</small></div>
    </div>
    ${(sc.get("재접수") || 0) ? `<p class="help">재접수 ${n(sc.get("재접수"))}건</p>` : ""}
    ${D.dong ? "" : barTable("동별", agg(rows, (r) => r.dong), rows, "dong")}
    ${barTable("업체별", agg(rows, (r) => r.vendor), rows, "vendor")}
    ${barTable("공종별", agg(rows, (r) => r.trade), rows, "trade")}
    ${barTable("유형별", agg(rows, (r) => r.type), rows, "type")}
    ${barTable("DTspace 처리상태", agg(rows, (r) => r.dts), rows, "dts")}
    <p class="help">막대: 진한 부분 = 완료. 줄을 누르면 해당 조건으로 조회합니다.</p>`;
  bindPreset((k) => { S.dash = { ...D, preset: k }; LS.set("dash", S.dash); renderDash(); });
  $("#d-dong").addEventListener("change", (e) => { S.dash = { ...D, dong: e.target.value }; LS.set("dash", S.dash); renderDash(); });
  $$(".bar-row").forEach((b) => b.addEventListener("click", () => {
    S.filter = { preset: D.preset, ...(D.dong ? { dong: D.dong } : {}), [b.dataset.k]: b.dataset.v === "(없음)" ? "" : b.dataset.v };
    LS.set("filter", S.filter);
    location.hash = "#/list";
  }));
}

/* ---------- 업체별 조치리스트 ---------- */
function renderVendor() {
  setHeader("업체별 조치리스트");
  setTab("vendor");
  if (!S.rows.length) { view().innerHTML = noData(); return; }
  const V = S.vend;
  const pre = PRESETS[V.preset] || PRESETS.mipa;
  const scope = S.rows.filter((r) => (!V.dong || r.dong === V.dong) && pre.test(r));
  const vc = agg(scope, (r) => r.vendor);
  const vendors = [...vc.entries()].sort((a, b) => (b[1].all - b[1].done) - (a[1].all - a[1].done));
  if (V.vendor && !vc.has(V.vendor)) V.vendor = "";
  let rows = V.vendor ? scope.filter((r) => r.vendor === V.vendor) : [];
  if (V.open) rows = rows.filter((r) => st(r) !== "완료");
  const units = new Map();
  for (const r of rows.sort(sortUnit)) {
    const k = `${r.dong}|${r.ho}`;
    units.set(k, (units.get(k) || 0) + 1);
  }
  view().innerHTML = `
    ${presetChips(V.preset)}
    <div class="card filters"><div class="grid2">
      <label class="span-2">업체<select id="v-vendor"><option value="">업체 선택 (${vendors.length}곳)</option>${vendors.map(([k, v]) => `<option value="${esc(k)}" ${k === V.vendor ? "selected" : ""}>${esc(k)} — 잔여 ${v.all - v.done}/${v.all}</option>`).join("")}</select></label>
      <label>동<select id="v-dong">${opt(uniq(S.rows, "dong"), V.dong, "전체 동", (d) => `${d} (${shortDong(d)})`)}</select></label>
      <label class="toggle-row"><input type="checkbox" id="v-open" ${V.open ? "checked" : ""}/> 완료 제외</label>
    </div></div>
    ${V.vendor ? `
      <div class="result-bar"><b>${esc(V.vendor)}</b> · 세대 ${units.size} · ${n(rows.length)}건</div>
      <div class="stack-actions">
        <button type="button" class="btn btn-kakao btn-lg" id="v-share">카톡으로 조치요청 보내기(텍스트)</button>
        <div class="row-actions"><button type="button" class="btn" id="v-xlsx">엑셀 파일</button><button type="button" class="btn" id="v-csv">CSV</button><button type="button" class="btn" id="v-copy">복사</button></div>
      </div>
      <div class="card"><h3 class="block-title">세대별</h3><div class="chip-row">${[...units.entries()].map(([k, c]) => { const [d, h] = k.split("|"); return `<a class="chip" href="#/unit/${d}/${h}">${shortDong(d)} ${+h} <b>${c}</b></a>`; }).join("") || '<span class="muted">없음</span>'}</div></div>
      <details class="card prev"><summary>보낼 내용 미리보기</summary><pre class="pre">${esc(vendorText(V.vendor, rows, ctx, pre.label))}</pre></details>`
    : `<div class="card"><h3 class="block-title">잔여 많은 업체</h3><div class="bars">${vendors.slice(0, 15).map(([k, v]) => `<button type="button" class="bar-row" data-vend="${esc(k)}"><span class="b-lab">${esc(k)}</span><span class="b-n">${n(v.all - v.done)}<small>/${n(v.all)}</small></span></button>`).join("")}</div></div>`}`;
  const save = () => LS.set("vend", S.vend);
  bindPreset((k) => { S.vend = { ...V, preset: k }; save(); renderVendor(); });
  $("#v-vendor").addEventListener("change", (e) => { S.vend = { ...V, vendor: e.target.value }; save(); renderVendor(); });
  $("#v-dong").addEventListener("change", (e) => { S.vend = { ...V, dong: e.target.value }; save(); renderVendor(); });
  $("#v-open").addEventListener("change", (e) => { S.vend = { ...V, open: e.target.checked }; save(); renderVendor(); });
  $$("[data-vend]").forEach((b) => b.addEventListener("click", () => { S.vend = { ...V, vendor: b.dataset.vend }; save(); renderVendor(); }));
  if (!V.vendor) return;
  const base = `${V.vendor.replace(/[()（）\s/\\]/g, "")}_${V.dong || "전체"}_${pre.label.replace("·", "")}_${todayStr().replace(/-/g, "").slice(2)}`;
  const text = () => vendorText(V.vendor, rows, ctx, pre.label);
  $("#v-share").addEventListener("click", async () => { const r = await shareText(text(), V.vendor); if (r === "copied") toast("복사했습니다. 카톡에 붙여넣으세요."); });
  $("#v-copy").addEventListener("click", async () => { await navigator.clipboard.writeText(text()).catch(() => {}); toast("복사했습니다."); });
  $("#v-xlsx").addEventListener("click", async () => { busy(true, "엑셀 만드는 중…"); try { await shareOrDownload(await toXlsx(rows, ctx, V.vendor), base + ".xlsx", V.vendor); } finally { busy(false); } });
  $("#v-csv").addEventListener("click", () => downloadBlob(toCsv(rows, ctx), base + ".csv"));
}

/* ---------- 설정 ---------- */
async function renderSettings() {
  setHeader("설정");
  setTab("settings");
  const last = S.imports[0];
  const est = await navigator.storage?.estimate?.().catch(() => null);
  const persisted = await navigator.storage?.persisted?.().catch(() => false);
  const editsN = [...S.edits.values()].filter((e) => e.status || e.memo || e.doneDate).length;
  const photosN = [...S.pcount.values()].reduce((a, b) => a + b, 0);
  view().innerHTML = `
    <div class="card"><h3 class="block-title">하자 데이터 (DTspace 엑셀)</h3>
      ${last ? `<table class="kv">
        <tr><th>파일</th><td>${esc(last.fileName)}</td></tr>
        <tr><th>가져온 시각</th><td>${esc(new Date(last.importedAt).toLocaleString("ko-KR"))}</td></tr>
        <tr><th>저장 건수</th><td>${n(S.rows.length)}건 (동 ${uniq(S.rows, "dong").length}개)</td></tr>
        <tr><th>최근 접수일</th><td>${esc(S.maxRcv || "-")} <small class="muted">(이번 파일 ${esc(last.maxRcv || "-")})</small></td></tr>
        <tr><th>지난 가져오기</th><td>신규 ${n(last.added)} · 변경 ${n(last.changed)} · 동일 ${n(last.same)}${last.gone ? ` · 엑셀에 없음 ${n(last.gone)}` : ""} · ${(last.totalMs / 1000).toFixed(1)}초</td></tr>
      </table>` : `<p class="help">아직 가져온 데이터가 없습니다.</p>`}
      <button type="button" class="btn btn-primary btn-lg" data-act="import">${last ? "새 엑셀로 업데이트" : "엑셀 가져오기"}</button>
      <p class="help">DTspace 하자관리(a.dtspace.co.kr) 조회 화면의 <b>Export Excel</b> 파일을 그대로 선택하세요. 같은 순번은 최신 내용으로 바뀌고, 내 상태·메모·사진은 그대로 남습니다. 데이터는 이 폰 브라우저 안(IndexedDB)에만 저장됩니다.</p>
    </div>
    <div class="card"><h3 class="block-title">백업 · 복원</h3>
      <p class="help">내 처리기록 ${n(editsN)}건 · 사진 ${n(photosN)}장. 폰 교체·앱 삭제 전 백업 zip을 받아 두세요(하자 원본은 엑셀로 다시 가져오면 됩니다).</p>
      <div class="row-actions"><button type="button" class="btn btn-primary" id="b-export">백업 zip 받기</button><button type="button" class="btn" id="b-import">백업 복원</button></div>
      ${(LS.get("lastBackup", "")) ? `<p class="help">마지막 백업 ${esc(new Date(LS.get("lastBackup")).toLocaleString("ko-KR"))}</p>` : ""}
    </div>
    <div class="card"><h3 class="block-title">저장공간</h3>
      <p class="help">사용 ${est ? (est.usage / 1048576).toFixed(1) + "MB" : "-"} / 가능 ${est ? Math.round(est.quota / 1048576).toLocaleString() + "MB" : "-"} · 영구 보관 ${persisted ? "켜짐" : "꺼짐"}</p>
      ${persisted ? "" : `<button type="button" class="btn" id="s-persist">영구 보관 요청(자동 삭제 방지)</button>`}
    </div>
    <div class="card"><h3 class="block-title">초기화</h3>
      <button type="button" class="btn btn-danger" id="x-defects">하자 원본 삭제(내 기록·사진은 유지)</button>
      <button type="button" class="btn btn-danger" id="x-all">전체 초기화(기록·사진 포함)</button>
    </div>
    <p class="help center">하자관리 v${APP_VERSION} · 오프라인 동작 · 서버 전송 없음</p>`;
  $("#b-export").addEventListener("click", async () => {
    busy(true, "백업 만드는 중…");
    try {
      const blob = await exportBackup();
      LS.set("lastBackup", new Date().toISOString());
      await shareOrDownload(blob, `하자관리_백업_${todayStr().replace(/-/g, "")}.zip`);
    } catch (e) { toast("백업 실패: " + e.message, 4000); } finally { busy(false); renderSettings(); }
  });
  $("#b-import").addEventListener("click", () => { $("#input-backup").value = ""; $("#input-backup").click(); });
  $("#s-persist")?.addEventListener("click", async () => { const ok = await navigator.storage?.persist?.(); toast(ok ? "영구 보관 켜짐" : "브라우저가 거절했습니다(앱 설치 후 다시 시도)"); renderSettings(); });
  $("#x-defects").addEventListener("click", async () => {
    if (!confirm("하자 원본 데이터를 지울까요? 내 상태·메모·사진은 남습니다.")) return;
    await db.clear("defects"); await db.setMeta("imports", []);
    await loadAll(); toast("삭제했습니다."); renderSettings();
  });
  $("#x-all").addEventListener("click", async () => {
    if (!confirm("모든 데이터(하자·내 기록·사진)를 지웁니다. 백업했나요?")) return;
    if (!confirm("정말 전체 초기화할까요? 되돌릴 수 없습니다.")) return;
    for (const s of ["defects", "edits", "photos", "meta"]) await db.clear(s);
    await loadAll(); toast("초기화했습니다."); renderSettings();
  });
}

/* ---------- 데이터 로드·가져오기 ---------- */
async function loadAll() {
  S.rows = await db.getAll("defects");
  // 이전 버전에 저장된 동·호수 표기 정리(901 → 0901 등). 내 기록·사진은 순번(id)으로 연결돼 그대로 유지
  const fixed = migrateRows(S.rows);
  if (fixed.length) await db.putMany("defects", fixed);
  const rec = LS.get("recent", []);
  const rec2 = [...new Set(rec.map((k) => { const [d, h] = k.split("|"); return `${normDong(d)}|${normHo(h)}`; }))];
  if (rec2.join() !== rec.join()) LS.set("recent", rec2);
  if (S.filter?.ho && S.filter.ho !== normHo(S.filter.ho)) { S.filter.ho = normHo(S.filter.ho); LS.set("filter", S.filter); }
  S.rows.sort(sortUnit);
  S.byId = new Map(S.rows.map((r) => [r.id, r]));
  S.edits = new Map((await db.getAll("edits")).map((e) => [e.id, e]));
  S.pcount = await db.photoCounts();
  S.imports = (await db.getMeta("imports")) || [];
  S.maxRcv = S.rows.reduce((m, r) => (r.rcv && r.rcv > m ? r.rcv : m), "");
}

$("#input-xlsx").addEventListener("change", async (ev) => {
  const f = ev.target.files?.[0];
  ev.target.value = "";
  if (!f) return;
  busy(true, "엑셀 읽는 중…");
  try {
    const info = await importFile(f, S.rows, (m) => busy(true, m), async (fm, sm) =>
      confirm(`이 엑셀의 최근 접수일(${fm})이 저장된 데이터(${sm})보다 이전입니다.\n오래된 파일일 수 있습니다. 그래도 덮어쓸까요?\n(내 상태·메모·사진은 유지됩니다)`));
    await loadAll();
    navigator.storage?.persist?.().catch(() => {});
    toast(`가져오기 완료: ${n(info.rows)}건 (신규 ${n(info.added)}, 변경 ${n(info.changed)}) ${(info.totalMs / 1000).toFixed(1)}초`, 4000);
    route();
  } catch (e) {
    console.error(e);
    toast("가져오기 실패: " + e.message, 5000);
  } finally {
    busy(false);
  }
});
$("#input-backup").addEventListener("change", async (ev) => {
  const f = ev.target.files?.[0];
  if (!f) return;
  busy(true, "복원 중…");
  try {
    const r = await importBackup(f);
    await loadAll();
    toast(`복원: 기록 ${r.edits}/${r.total}건, 사진 ${r.photos}/${r.totalPhotos}장`, 4000);
    route();
  } catch (e) { toast("복원 실패: " + e.message, 5000); } finally { busy(false); }
});
$("#input-camera").addEventListener("change", (e) => onPhotoFiles([...e.target.files]));
$("#input-album").addEventListener("change", (e) => onPhotoFiles([...e.target.files]));
$("#btn-back").addEventListener("click", () => (history.length > 1 ? history.back() : (location.hash = "#/home")));

/* ---------- 라우터 ---------- */
function route() {
  const parts = (location.hash.replace(/^#\/?/, "") || "home").split("/");
  const [p, a, b] = parts;
  window.scrollTo(0, 0);
  if (p === "dong" && a) return renderDong(normDong(decodeURIComponent(a)));
  if (p === "unit" && a && b) return renderUnit(normDong(decodeURIComponent(a)), normHo(decodeURIComponent(b)));
  if (p === "item" && a) return renderItem(+a);
  if (p === "list") return renderList();
  if (p === "dash") return renderDash();
  if (p === "vendor") return renderVendor();
  if (p === "settings") return renderSettings();
  return renderHome();
}
window.addEventListener("hashchange", route);

(async () => {
  busy(true, "불러오는 중…");
  try { await loadAll(); } catch (e) { toast("데이터 열기 실패: " + e.message, 5000); }
  busy(false);
  route();
  if ("serviceWorker" in navigator) navigator.serviceWorker.register("./sw.js").catch(() => {});
  window.__defectApp = { S, loadAll, parseQuery };
})();
