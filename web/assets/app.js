/* 本地版共享工具 —— 移植自 invoice-cloud public/assets/app.js（去掉登录跳转与 SheetJS） */
const $ = (id) => document.getElementById(id);
const esc = (s) => String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

function toast(msg, ok) {
  const t = $("toast");
  t.textContent = msg;
  t.className = "toast " + (ok ? "ok" : "err");
  t.style.display = "block";
  clearTimeout(t._tm);
  t._tm = setTimeout(() => (t.style.display = "none"), 3200);
}

async function api(path, opts) {
  let r;
  try {
    r = await fetch(path, opts);
  } catch (e) {
    throw new Error("网络错误（本地服务未响应，请重新运行程序）：" + e.message);
  }
  const ct = (r.headers.get("content-type") || "").toLowerCase();
  if (!ct.includes("application/json")) {
    const raw = await r.text().catch(() => "");
    const brief = raw.replace(/<[^>]*>/g, " ").replace(/\s+/g, " ").trim().slice(0, 100);
    throw new Error(`服务异常 HTTP ${r.status}（${ct.split(";")[0] || "无Content-Type"}）${brief ? "：" + brief : ""}`);
  }
  const j = await r.json();
  if (!j.ok) throw new Error(j.error || "接口错误");
  return j.data;
}
