# 沐塔 MUTTA 分單系統 V4（合併版）

## 架構
- 單一後端 main.py，依前端 `shipper` 參數內部分流
- 7-11 模式：PDF + Excel → 重繪 V4 卡片（擷取撕線區條碼）
- 順豐模式：PDF + Excel + CSV → 重繪 V4 卡片（上半部留白）+ 產出順豐批量下單 Excel
- 兩模式共用 config、core_logic、purchase_history、dev 後台

## 資料夾結構
```
mutta_merged/
├── app/
│   ├── main.py             # 合併版後端入口（含 shipper 分流）
│   ├── core_logic.py       # 共用核心（不變）
│   ├── v4_engine.py        # 7-11 引擎：擷取撕線區
│   └── v4_engine_sf.py     # 順豐引擎：上半部留白
├── templates/
│   ├── index.html          # 合併前端，左上角切換鈕
│   └── dev.html            # 開發者後台（不變）
├── static/                 # 你原本的靜態資源（複製過來即可）
├── sf_template_clean.xlsx  # 順豐批量下單 Excel 模板（必須放在專案根）
├── config.json             # 自動產生
├── purchase_history/       # 自動產生
└── uploads/                # 自動產生
```

## 部署步驟

1. **覆蓋程式碼檔**：
   - 把 `app/` 整個資料夾覆蓋到你現有的 `app/`
   - 把 `templates/index.html` 覆蓋掉舊的（dev.html 不變）

2. **保留原有檔案**：
   - `sf_template_clean.xlsx` 必須放在專案根（與 `app/` 同層）
   - `config.json`、`purchase_history/`、`static/` 完全不動

3. **啟動指令** 不變：
   ```
   uvicorn main:app --app-dir app --host 0.0.0.0 --port 8000
   ```

## 順豐寄件方資訊（sf_fixed）

預設值已寫入 `DEFAULT_CONFIG`（沐塔/林口/931753889 等）。
如需修改，可：
- 直接編輯 `config.json` 中的 `sf_fixed` 欄位
- 或透過 `POST /api/dev/sf-fixed` API 修改

⚠️ 開發者後台 dev.html **暫未加上** sf_fixed 設定 UI，
   後端 API 已就緒，待之後再加 UI（不影響系統運作）。

## 主要改動摘要

### 後端（app/main.py）
- `/api/upload` 新增 `shipper` 參數（seven|sf）與 `csv_files`（順豐必填）
- `_run_job()` 變成 dispatcher，按 shipper 分流到 `_run_job_seven()` / `_run_job_sf()`
- 新增 `_build_sf_excel()`、`_parse_csv_recipients()`、`_scan_pdfs_v4()`（共用）
- 新增 `/api/dev/sf-fixed` GET/POST API
- 順豐流程的 PDF 只用於訂單號驗證（不擷取、不抓收件資訊）
- 順豐流程末尾產出順豐批量下單 Excel

### 前端（templates/index.html）
- 左上角加切換鈕（7-11 ⇄ 順豐）
- 切換時動態變更：主色、Logo 文字、副標、上傳卡欄數
- 順豐模式自動顯示第三張 CSV 卡
- 上傳時帶 `shipper` 與 `csv_files`

### 引擎
- `v4_engine.py`：7-11 用，不變
- `v4_engine_sf.py`：順豐用，與 7-11 唯一差異 = `generate_page()` 跳過 `extract_tear_image()`，上半部留白
