# StampTool · 文件盖章（本地版）

把发票云（invoice-cloud）的「文件盖章」页面 1:1 移植到本地，打包成**绿色单文件 exe**。双击运行、零安装零配置，所有渲染 / 预览 / 印章融合都在浏览器本地完成，文件不出本机。

## 功能

- 单页盖章（同页可多枚）+ 骑缝章（整章竖切 N 条，N 页右缘各一条）
- 拖入 PDF / JPG / PNG（图片自动转 A4 PDF）→ 选章 → 拖拽或方向键微调（0.5mm）→ 下载盖章后的 PDF
- 印章来源：`exe 同目录/stamps/`；也可在页面设置任意文件夹，持久化到 `StampTool.json`，下次免设
- 固定直径 42mm；点「下载」时才用 pdf-lib 把印章 PNG 融合进 PDF（不会先正片叠底等）

## 使用（已构建）

去 **Releases** 下载 `StampTool.exe`，双击即可：

```
双击 exe → 自动开浏览器 http://127.0.0.1:8731/（端口占用自动 +1）
→ 拖 PDF → 选章盖章 → 下载
```

需要印章：把 PNG/JPG 放进 `StampTool.exe` 同目录的 `stamps/` 文件夹（可在页面上传/删除）。

## 自行构建

```bash
npm install
npm run build        # → build/StampTool.exe（Node SEA 单文件）
```

- 需要 Node 22+（SEA 打包依赖）。
- `build/tools/` 下放入 `upx-*/upx.exe` 时，`npm run build` 会自动 UPX 压缩（86MB → ~22MB）；不放则输出未压缩版，功能一致。
- `build/sea-config.json` 与 `build/sea-prep.blob` 为打包中间产物，勿入库，运行 build 自动再生。

## 目录结构

```
StampTool/
├── src/server.js          # 本地 HTTP 服务（端口 8731，回环）
├── web/                   # SPA：index.html + assets + vendor(pdf.js 3.11 UMD)
├── build/
│   ├── build_sea.mjs      # SEA 打包脚本（node server.js + web/ → exe）
│   └── tools/upx-*/       # 可选 UPX 压缩
├── stamps/                # 用户印章（运行时，git 忽略）
├── StampTool.json         # 运行配置（自动生成，git 忽略）
└── package.json
```

## 仓库说明（脱敏）

本仓库**不含任何密钥 / 域名 / 业务数据**——StampTool 是纯本地工具。

`.gitignore` 排除：`node_modules/`、`build/StampTool.exe`（产物走 Release）、`build/sea-prep.blob`、`build/sea-config.json`（含本地路径，可重建）、`StampTool.json`（运行时配置）、`stamps/`（用户印章）、`.workbuddy/`。
