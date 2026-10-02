/* 盖章页前端逻辑 —— 1:1 移植自 invoice-cloud /stamp（v3 线上已验证版）
   坐标体系（v2，修复跑偏）：positions 一律存 scale=1 显示点（pdf.js viewport 口径 = CropBox 显示区域）；
   盖章输出用 viewport.convertToPdfPoint() 换算 user space（自动消化 CropBox/MediaBox 偏移、页面旋转、缩放），
   保证所见即所得。依赖：app.js 的 $/esc/toast/api；pdfjsLib 3.11 UMD；PDFLib 1.17 UMD。
   本地版改动（均已标注）：worker/cmaps/standard_fonts 走本地 vendor；印章接口改本地目录版；
   去掉 R2 归档/Word·Excel 转换；新增印章目录设置。盖章核心逻辑零改动。 */
pdfjsLib.GlobalWorkerOptions.workerSrc = "/vendor/pdf.worker.min.js"; // 本地版：离线
const CMAP_URL = "/vendor/cmaps/";                 // 本地版：离线
const STANDARD_FONTS = "/vendor/standard_fonts/";  // 本地版：离线（云端未设，本地补上更稳）

const SEAL_PT = 119.055;  // 章物理尺寸 42mm（磅）= 42 × 72 / 25.4
const BASE_SCALE = 1.5;   // canvas 渲染基础倍率（清晰度）

// ---------- 状态 ----------
let pdfBytes = null, pdfDoc = null, pdfName = "";
let pageNum = 1, pageCount = 0, zoomPct = 100;
let seals = [], selSeal = null;
let mode = "normal";            // normal | qiefeng
let positions = [];             // normal {mode,sealKey,name,page,x,y} / qiefeng {mode,sealKey,name,y}（x,y=章左上，scale=1 显示点）
let rendering = false, needRender = false;
let dispSize = {};              // pageNum -> {w,h}（scale=1 显示磅，rotation 已应用）
let vp1 = {};                   // pageNum -> scale=1 viewport（convertToPdfPoint 用，含 CropBox/旋转变换）
let qfCache = null;             // 骑缝切片缓存 {sealKey, n, slices:[dataURL]}
let drag = null;                // 拖拽状态 {p, k, sx, sy, ox, oy}

// ---------- 印章 ----------
async function loadSeals() {
  try {
    seals = await api("/api/seals");
    for (const s of seals) {
      const r = await fetch("/api/seal?key=" + encodeURIComponent(s.key)); // 本地版
      if (!r.ok) continue;
      s.bytes = new Uint8Array(await r.arrayBuffer());
      s.url = URL.createObjectURL(new Blob([s.bytes], { type: "image/png" }));
      s.imgEl = await new Promise((ok) => { const im = new Image(); im.onload = () => ok(im); im.onerror = () => ok(null); im.src = s.url; });
    }
  } catch (e) { toast("印章加载失败：" + e.message, "err"); }
  renderSealList();
}

function renderSealList() {
  const el = $("sealList");
  if (!seals.length) { el.innerHTML = '<div style="grid-column:1/-1;color:#9aa7b5;font-size:12px">暂无印章，先上传一枚</div>'; return; }
  el.innerHTML = seals.map((s) =>
    `<div class="seal-item ${selSeal && selSeal.key === s.key ? "sel" : ""}" data-key="${esc(s.key)}" title="${esc(s.name)}">`
    + `<img src="${s.url}" alt="${esc(s.name)}">`
    + `<button class="rm" data-del="${esc(s.key)}" title="删除印章">✕</button></div>`).join("");
}

$("sealList").addEventListener("click", async (e) => {
  const del = e.target.closest("[data-del]");
  if (del) {
    e.stopPropagation();
    if (!confirm("确认删除该印章？")) return;
    await api("/api/seal-del", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ key: del.dataset.del }) }); // 本地版
    if (selSeal && selSeal.key === del.dataset.del) selSeal = null;
    await loadSeals();
    return toast("印章已删除", "ok");
  }
  const item = e.target.closest(".seal-item");
  if (!item) return;
  const s = seals.find((x) => x.key === item.dataset.key);
  if (!s) return;
  selSeal = selSeal && selSeal.key === s.key ? null : s;
  renderSealList();
  $("hint").textContent = selSeal ? `已选「${s.name}」` : "";
  if (selSeal && mode === "qiefeng") buildQf();
});

$("btnSealUpload").addEventListener("click", () => $("sealInput").click());
$("sealInput").addEventListener("change", async (e) => {
  const f = e.target.files[0];
  e.target.value = "";
  if (!f) return;
  await api("/api/seal-upload", { method: "POST", headers: { "X-Filename": encodeURIComponent(f.name), "Content-Type": "application/octet-stream" }, body: f }); // 本地版：原始字节+文件名头
  toast("印章已上传", "ok");
  await loadSeals();
});

// ---------- 印章目录（本地版新增） ----------
async function loadDir() {
  try {
    const d = await api("/api/stampdir");
    $("dirVal").textContent = d.dir;
    $("dirRow").title = d.dir;
  } catch (e) { $("dirVal").textContent = "读取失败"; }
}
$("btnDir").addEventListener("click", async () => {
  const d = await api("/api/stampdir").catch(() => ({ dir: "" }));
  $("dirInput").value = d.dir || "";
  $("dlgDir").showModal();
});
$("btnDirClose").addEventListener("click", () => $("dlgDir").close());
$("btnDirSave").addEventListener("click", async () => {
  try {
    const d = await api("/api/stampdir", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ dir: $("dirInput").value.trim() }) });
    $("dlgDir").close();
    $("dirVal").textContent = d.dir;
    $("dirRow").title = d.dir;
    seals = []; selSeal = null;
    await loadSeals();
    toast("印章目录已保存，下次免设", "ok");
  } catch (e) { toast(e.message, "err"); }
});

// ---------- 模式 ----------
function setMode(m) {
  mode = m;
  $("modeNormal").className = "btn sm" + (m === "normal" ? " pri" : "");
  $("modeQiefeng").className = "btn sm" + (m === "qiefeng" ? " pri" : "");
  $("modeTip").textContent = m === "normal"
    ? "选一枚章，在页面上点击放置（点击点为章中心），落章后可拖拽微调；可跨页多次放章。"
    : "选一枚章后在页面点击一次定纵向位置（可拖拽微调），全部页右缘自动切条盖好（第 1 页盖最左条，拼合即完整章）。";
  if (m === "qiefeng" && selSeal && pdfDoc) buildQf();
  hideFloat();
}
$("modeNormal").addEventListener("click", () => setMode("normal"));
$("modeQiefeng").addEventListener("click", () => setMode("qiefeng"));

async function buildQf() {
  if (!selSeal || !selSeal.imgEl || !pageCount) return;
  const img = selSeal.imgEl, n = pageCount;
  const cw = Math.ceil(img.naturalWidth / n), ch = img.naturalHeight;
  const slices = await Promise.all([...Array(n)].map((_, i) => new Promise((res) => {
    const c = document.createElement("canvas");
    c.width = cw; c.height = ch;
    c.getContext("2d").drawImage(img, -i * cw, 0);
    c.toBlob((b) => { const fr = new FileReader(); fr.onload = () => res(fr.result); fr.readAsDataURL(b); }, "image/png");
  })));
  qfCache = { sealKey: selSeal.key, n, slices };
  renderOverlay();
}

// ---------- 渲染 ----------
async function renderPage() {
  if (rendering) { needRender = true; return; }
  rendering = true;
  try {
    const page = await pdfDoc.getPage(pageNum);
    const v1v = page.getViewport({ scale: 1 });
    vp1[pageNum] = v1v;
    dispSize[pageNum] = { w: v1v.width, h: v1v.height };
    const vp = page.getViewport({ scale: zoomPct / 100 * BASE_SCALE });
    const canvas = $("pdfCanvas");
    canvas.width = Math.round(vp.width);
    canvas.height = Math.round(vp.height);
    await page.render({ canvasContext: canvas.getContext("2d"), viewport: vp }).promise;
    $("pageInfo").textContent = `${pageNum} / ${pageCount}`;
    renderOverlay();
  } finally {
    rendering = false;
    if (needRender) { needRender = false; renderPage(); }
  }
}

// canvas CSS 显示宽 → scale=1 显示点的比例（自动覆盖 zoom 与 CSS 压缩）
function cssK() {
  const rect = $("pdfCanvas").getBoundingClientRect();
  return rect.width / dispSize[pageNum].w;
}

function renderOverlay() {
  const ov = $("overlay"), canvas = $("pdfCanvas");
  ov.innerHTML = "";
  const d = dispSize[pageNum];
  if (!d || !canvas.width) return;
  const k = cssK();
  const sealCss = SEAL_PT * k;
  const rect = canvas.getBoundingClientRect();
  for (const p of positions) {
    const seal = seals.find((s) => s.key === p.sealKey);
    if (!seal) continue;
    const el = document.createElement("div");
    el.className = "stamped";
    el.style.cursor = "move";
    if (p.mode === "normal") {
      if (p.page !== pageNum) continue;
      el.style.cssText += `left:${p.x * k}px;top:${p.y * k}px;width:${sealCss}px;height:${sealCss}px`;
      el.innerHTML = `<img src="${seal.url}" alt=""><button class="rm" title="删除">✕</button>`;
    } else {
      if (!qfCache || qfCache.sealKey !== p.sealKey) continue;
      const sw = sealCss / qfCache.n;
      el.style.cssText += `left:${rect.width - sw}px;top:${p.y * k}px;width:${sw}px;height:${sealCss}px`;
      el.innerHTML = `<img src="${qfCache.slices[pageNum - 1]}" alt=""><button class="rm" title="删除">✕</button>`;
    }
    el.querySelector(".rm").addEventListener("click", () => {
      positions.splice(positions.indexOf(p), 1);
      renderOverlay();
      renderPosList();
    });
    el.addEventListener("pointerdown", (ev) => startDrag(ev, p));
    ov.appendChild(el);
  }
}

function renderPosList() {
  const el = $("posList");
  if (!positions.length) { el.innerHTML = '<div style="color:#9aa7b5">尚未盖章</div>'; return; }
  el.innerHTML = positions.map((p, i) =>
    `<div class="pos-row"><span class="k">${p.mode === "qiefeng" ? "骑缝" : "第 " + p.page + " 页"} · ${esc(p.name)}</span><button class="rm" data-i="${i}">删除</button></div>`).join("");
}

$("posList").addEventListener("click", (e) => {
  const b = e.target.closest(".rm");
  if (!b) return;
  positions.splice(+b.dataset.i, 1);
  renderOverlay();
  renderPosList();
});

// ---------- 拖拽微调 ----------
const wrap = $("canvasWrap");

function startDrag(ev, p) {
  if (ev.target.closest(".rm")) return;
  ev.preventDefault();
  const rect = $("pdfCanvas").getBoundingClientRect();
  drag = { p, k: dispSize[pageNum].w / rect.width, sx: ev.clientX, sy: ev.clientY, ox: p.x, oy: p.y };
}
wrap.addEventListener("pointermove", (ev) => {
  if (!drag) return;
  const dx = (ev.clientX - drag.sx) * drag.k, dy = (ev.clientY - drag.sy) * drag.k;
  if (drag.p.mode === "normal") {
    drag.p.x = drag.ox + dx;
    drag.p.y = drag.oy + dy;
  } else {
    drag.p.y = drag.oy + dy;
  }
  renderOverlay();
});
wrap.addEventListener("pointerup", () => { drag = null; });
wrap.addEventListener("pointerleave", () => { drag = null; });

// ---------- 摆章交互（浮动虚影 + 点击放章；坐标统一 scale=1 显示点） ----------

function hideFloat() { $("floatSeal").style.display = "none"; }

wrap.addEventListener("mousemove", (e) => {
  if (!selSeal || !pdfDoc || !dispSize[pageNum] || drag) return hideFloat();
  const rect = $("pdfCanvas").getBoundingClientRect();
  if (e.clientX < rect.left || e.clientX > rect.right || e.clientY < rect.top || e.clientY > rect.bottom) return hideFloat();
  const k = rect.width / dispSize[pageNum].w;
  const sq = SEAL_PT * k;
  const f = $("floatSeal"), fi = $("floatImg");
  if (mode === "normal") {
    fi.src = selSeal.url;
    f.style.cssText = `display:block;left:${e.clientX - rect.left - sq / 2}px;top:${e.clientY - rect.top - sq / 2}px;width:${sq}px;height:${sq}px`;
  } else {
    const hasSlice = qfCache && qfCache.sealKey === selSeal.key;
    const n = hasSlice ? qfCache.n : pageCount;
    fi.src = hasSlice ? qfCache.slices[pageNum - 1] : selSeal.url;
    f.style.cssText = `display:block;opacity:${hasSlice ? ".65" : ".3"};left:${rect.width - sq / n}px;top:${e.clientY - rect.top - sq / 2}px;width:${sq / n}px;height:${sq}px`;
  }
});
wrap.addEventListener("mouseleave", hideFloat);

wrap.addEventListener("click", (e) => {
  if (!selSeal || !pdfDoc) return;
  if (e.target.closest(".stamped")) return;
  const canvas = $("pdfCanvas");
  const rect = canvas.getBoundingClientRect();
  const d = dispSize[pageNum];
  const x1 = (e.clientX - rect.left) * d.w / rect.width;   // 点击点 → scale1 显示点
  const y1 = (e.clientY - rect.top) * d.h / rect.height;
  if (x1 < 0 || y1 < 0 || x1 > d.w || y1 > d.h) return;
  const half = SEAL_PT / 2;
  if (mode === "normal") {
    positions.push({ mode: "normal", sealKey: selSeal.key, name: selSeal.name, page: pageNum, x: x1 - half, y: y1 - half });
    selSeal = null;                            // 放完释放选中（同原版）
    renderSealList();
    $("hint").textContent = "";
    hideFloat();
    toast("印章已放置，可拖拽微调", "ok");
  } else {
    if (!qfCache || qfCache.sealKey !== selSeal.key) { buildQf(); return toast("切片生成中，请再点一次", "err"); }
    positions = positions.filter((p) => p.mode !== "qiefeng" || p.sealKey !== selSeal.key);  // 同章覆盖
    positions.push({ mode: "qiefeng", sealKey: selSeal.key, name: selSeal.name, y: y1 - half });
    toast(`骑缝章已放置（${pageCount} 页右缘），可拖拽微调`, "ok");
  }
  renderOverlay();
  renderPosList();
});

// ---------- 翻页 / 缩放 ----------
$("btnPrev").addEventListener("click", () => { if (pageNum > 1) { pageNum--; hideFloat(); renderPage(); } });
$("btnNext").addEventListener("click", () => { if (pageNum < pageCount) { pageNum++; hideFloat(); renderPage(); } });
$("zoom").addEventListener("input", (e) => {
  zoomPct = +e.target.value;
  $("zoomVal").textContent = zoomPct + "%";
  if (pdfDoc) renderPage();
});

// ---------- 文件加载（本地版：PDF 直接盖；jpg/png 前端转 PDF；其余类型提示不支持） ----------
async function loadFile(file) {
  const ext = (file.name.split(".").pop() || "").toLowerCase();
  try {
    let bytes;
    if (ext === "pdf") {
      bytes = new Uint8Array(await file.arrayBuffer());
    } else if (["jpg", "jpeg", "png"].includes(ext)) {
      bytes = await imageToPdf(file, ext);
    } else {
      throw new Error("本地版仅支持 PDF / JPG / PNG（Word、Excel 请先另存为 PDF）");
    }
    await loadPdf(bytes, file.name.replace(/\.[^.]+$/, ""));
  } catch (e) {
    toast(e.message, "err");
  }
}

async function imageToPdf(file, ext) {
  const { PDFDocument } = PDFLib;
  const doc = await PDFDocument.create();
  const buf = await file.arrayBuffer();
  const img = ext === "png" ? await doc.embedPng(buf) : await doc.embedJpg(buf);
  const A4W = 595.28, A4H = 841.89, pad = 24;
  const k = Math.min((A4W - pad * 2) / img.width, (A4H - pad * 2) / img.height, 1);
  const w = img.width * k, h = img.height * k;
  const page = doc.addPage([A4W, A4H]);
  page.drawImage(img, { x: (A4W - w) / 2, y: (A4H - h) / 2, width: w, height: h });
  return new Uint8Array(await doc.save());
}

async function loadPdf(bytes, name) {
  pdfBytes = bytes;
  pdfName = name;
  pdfDoc = await pdfjsLib.getDocument({ data: bytes.slice(0), cMapUrl: CMAP_URL, cMapPacked: true, standardFontDataUrl: STANDARD_FONTS }).promise;  // slice 隔离，防 worker transfer 掏空 pdfBytes
  pageCount = pdfDoc.numPages;
  pageNum = 1;
  positions = [];
  qfCache = null;
  dispSize = {};
  vp1 = {};
  $("dropzone").style.display = "none";
  $("workspace").style.display = "";
  $("fileName").textContent = name;
  $("posList").innerHTML = "";
  renderPosList();
  await renderPage();
}

// 上传交互
$("btnPick").addEventListener("click", () => $("fileInput").click());
$("fileInput").addEventListener("change", (e) => { if (e.target.files[0]) loadFile(e.target.files[0]); e.target.value = ""; });
const dz = $("dropzone");
dz.addEventListener("click", (e) => { if (!e.target.closest("button")) $("fileInput").click(); });
dz.addEventListener("dragover", (e) => { e.preventDefault(); dz.classList.add("drag"); });
dz.addEventListener("dragleave", () => dz.classList.remove("drag"));
dz.addEventListener("drop", (e) => { e.preventDefault(); dz.classList.remove("drag"); if (e.dataTransfer.files[0]) loadFile(e.dataTransfer.files[0]); });

$("btnReset").addEventListener("click", () => {
  if (pdfDoc) { try { pdfDoc.destroy(); } catch {} pdfDoc = null; }
  pdfBytes = null; positions = []; qfCache = null; selSeal = null; dispSize = {}; vp1 = {};
  $("workspace").style.display = "none";
  $("dropzone").style.display = "";
});

// ---------- 盖章输出（pdf-lib 前端完成；convertToPdfPoint 自动处理 CropBox/MediaBox 偏移与页面旋转） ----------
async function ensureViewport(no) {
  if (!vp1[no]) {   // 未渲染过的页（多页骑缝必然包含）现场补建
    const pg = await pdfDoc.getPage(no);
    const v = pg.getViewport({ scale: 1 });
    vp1[no] = v;
    dispSize[no] = { w: v.width, h: v.height };
  }
  return vp1[no];
}
async function stampOut() {
  const { PDFDocument } = PDFLib;
  const doc = await PDFDocument.load(pdfBytes);
  const embeds = new Map();
  const embedOf = async (key) => {
    if (!embeds.has(key)) embeds.set(key, await doc.embedPng(seals.find((s) => s.key === key).bytes));
    return embeds.get(key);
  };
  // 显示矩形（scale1 点，左上+右下）→ user space 包围盒（两点转换覆盖旋转/裁剪偏移）
  const box = (v, x1, y1, x2, y2) => {
    const [ax, ay] = v.convertToPdfPoint(x1, y1);
    const [bx, by] = v.convertToPdfPoint(x2, y2);
    return { x: Math.min(ax, bx), y: Math.min(ay, by), width: Math.abs(bx - ax), height: Math.abs(by - ay) };
  };
  for (const p of positions.filter((q) => q.mode === "normal")) {
    const v = await ensureViewport(p.page);
    const page = doc.getPage(p.page - 1);
    page.drawImage(await embedOf(p.sealKey), box(v, p.x, p.y, p.x + SEAL_PT, p.y + SEAL_PT));
  }
  for (const p of positions.filter((q) => q.mode === "qiefeng")) {
    if (!qfCache || qfCache.sealKey !== p.sealKey) throw new Error("骑缝切片未就绪，请重放骑缝章");
    const sw = SEAL_PT / pageCount;
    for (let i = 0; i < pageCount; i++) {
      const v = await ensureViewport(i + 1);
      const page = doc.getPage(i);
      page.drawImage(await doc.embedPng(qfCache.slices[i]), box(v, dispSize[i + 1].w - sw, p.y, dispSize[i + 1].w, p.y + SEAL_PT));
    }
  }
  return doc.save();
}

$("btnDownload").addEventListener("click", async () => {
  if (!positions.length) return toast("请先放置印章", "err");
  const btn = $("btnDownload");
  btn.disabled = true;
  try {
    const out = await stampOut();
    const blob = new Blob([out], { type: "application/pdf" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = pdfName + "_盖章.pdf";   // 本地版：直接下载，无归档
    document.body.appendChild(a);
    a.click();
    a.remove();
    toast("已下载：" + a.download, "ok");
  } catch (e) {
    toast("盖章失败：" + e.message, "err");
  }
  btn.disabled = false;
});

// ---------- 初始化 ----------
setMode("normal");
loadSeals();
loadDir();   // 本地版

// 本地版心跳：与本地服务保持一条 SSE 长连接；全部页面关闭后服务 8s 自动退出（cmd 窗口随之关闭）。
// 断线由浏览器自动重连，刷新页面不会误杀（服务端有 8s 宽限）。
try { new EventSource("/api/events"); } catch (e) {}
