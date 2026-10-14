# Office 文档转 PDF（本地直出）更新计划

> StampTool · 文件盖章本地版
> 目标：补回 invoice-cloud「文件盖章」页的 office→pdf 能力，但后端改为**本地直接调 MS Office**，不再依赖远程 office-convert 服务。

## 一、背景与差距

本地版 `stamp.js` 的 `loadFile` 只接受 pdf/jpg/png，其余类型直接 `throw`；`server.js` 无任何转换接口。invoice-cloud 原版通过 `POST /api/convert`（打到远程 office-convert Python 服务）把 doc/docx/xls/xlsx 转成 PDF 再 `loadPdf`。本工具已本地化、要分发，应**在本机直接调 Office COM** 完成转换，保持单 exe、零依赖、数据不出本机。

## 二、已定决策

1. **后端方式**：Node 直接 `spawn` PowerShell COM，**不引入 Python 服务**。理由：本工具定位是绿色单 exe / 零依赖 / 数据不出本机，Python 服务与这一定位冲突；转换为偶发操作，每次 ~2.5s 冷启动可接受。
2. **检测方式**：**注册表快筛（缺席门）+ COM 实探（权威）**，启动时一次性完成，结果缓存入 `/api/health` 的 `office / word / excel`。
   - 注册表只用于"肯定缺席"快筛（都没有才跳过 COM 启动，省 ~2.5s），**绝不以"注册表存在"作为有 Office 的依据**。
   - COM 实探（`New-Object -ComObject Word/Excel.Application` 并读 `.Version`）是唯一权威。Office 卸载后即使注册表有残留，COM 探针也会正确判否，**不会误报**。
3. **转换映射**（沿用 office-convert 已实证做法）：
   - Word：`Documents.Open` → `SaveAs(FileFormat=17=wdFormatPDF)`
   - Excel：`Workbooks.Open(UpdateLinks=0)` → `ExportAsFixedFormat(Type=0=xlTypePDF)`

## 三、改动面（纯增量，不动现有接口与盖章逻辑）

### 1. `src/server.js`
- `require` 增加 `os` 与 `execFile / execFileSync`（保留 `exec`）。
- 新增 `OFFICE_CAP` 状态与 `detectOffice()`：单 powershell 进程内完成"注册表扫描 + Word/Excel COM 实探"，解析 `WORD=...;EXCEL=...`。异常（无 powershell / COM 失败）→ 判无，绝不误报有。
- 启动流程：现有 `loadConfig()` / `ensureDefaultDir()` 之后 `await detectOffice()`，再 `findPort` / `listen`。浏览器地址日志附 Office 能力。
- `GET /api/health`：在 `data` 中增加 `office`(bool)、`word`(bool)、`excel`(bool)。
- 新增 `POST /api/convert`：
  - 读 `X-Filename` 头取扩展名；非 doc/docx/xls/xlsx → 400 明确错误。
  - 缺对应组件（docx 无 Word / xlsx 无 Excel）→ 400 灰字文案。
  - 读 body（上限 50MB）→ 写临时文件（guid 命名，避中文）→ 调 `convertOffice()` → 返回 PDF 字节。
  - `convertOffice()`：临时目录 `os.tmpdir()/stamptool-<guid>`；通过环境变量把源/目标/类型传给内联 PowerShell 脚本（`-STA` 必需），超时 180s；**串行锁**防两个 Office 实例并发抢 COM；无论成败清理临时目录；失败映射为 500 带原因。

### 2. `web/assets/stamp.js`
- `loadFile`：`doc/docx/xls/xlsx` 分支改为调用新增的 `officeToPdf(file, ext)`（fetch `/api/convert`，拿 PDF 字节喂 `loadPdf`）；其余非支持类型保留 `throw`。
- 新增 `officeToPdf()`：转换前显示 `#convStatus`"正在将 XX 转为 PDF…"，失败映射错误信息，结束后清空。
- 初始化新增 `applyOfficeCapability()`：`fetch /api/health`，按 `word/excel` 动态设置 `fileInput.accept`（有能力才接受 office 扩展名），并按 `office` 显隐灰字提示 `#officeHint`。

### 3. `web/index.html`
- `fileInput` 的 `accept` 初始保留，最终由 JS 按能力动态收紧。
- 拖拽区（`upcard`）新增两个元素：
  - `#officeHint`：默认 `display:none`，灰字（#9aa7b5）"本地无 MS Office，不支持文档转 PDF 直出，仅支持 PDF 盖章"。
  - `#convStatus`：转换进度行（橙 #b06a00 粗体，min-height 占位防跳动）。

## 四、关键技术点 / 坑位

- **PowerShell 必须 `-STA`**：Office COM 自动化需在单线程套间运行，否则报 RPC 错误（office-convert.py 里 `pythoncom.CoInitialize()` 同理）。探测与转换两条命令都加 `-STA -NoProfile`。
- **注册表残留不致误判**：仅用注册表"缺席"作快筛，COM 实探为唯一权威；卸载残留骗不过 COM 探针。
- **中文路径**：上传字节由 Node 落临时文件，文件名用 guid（仅借原始名取扩展名），目标 PDF 也用 guid；PS 脚本通过环境变量收路径，彻底规避引号/中文编码问题。
- **串行**：Node 侧用 Promise 链做转换锁，避免两文件同时触发两个 COM 实例互抢。
- **版本无关**：`SaveAs(17)` / `ExportAsFixedFormat(0)` 为稳定枚举，2013–2021/365 通用。
- **无桌面会话**：若以后当服务跑（无交互桌面）COM 会失败 → 自然走灰字分支，逻辑无需变。
- **超时与清理**：转换超时 180s；`finally` + 回调后都 `fs.rmSync(recursive)` 清理临时目录。

## 五、测试

1. 本机（已确认 Word/Excel COM 可用）：拖入 .docx / .xlsx → 转换成功并进入盖章；`/api/health` 返回 `office=true`。
2. 用 PowerShell COM 实测一次 docx→pdf 端到端，确认 `CONVERT_PS` 脚本可用。
3. `curl /api/convert` 传非 office 文件 → 400 明确错误（不崩）。
4. 模拟无 Office（临时让探测返回 false）→ 验证灰字提示显示、`accept` 收紧、拖 docx 报灰字文案。
5. 大文件 / 含公式 xlsx 转换正确性；转换失败返回明确错误而非 500 裸崩。

## 六、回滚

纯增量，不触碰现有 `/api/seals`、`/api/stampdir`、盖章 `stampOut` 等逻辑。若验收未过，回退 `server.js` / `stamp.js` / `index.html` 三处改动即可，不影响已发布 exe 的现有功能。
