"""
V4.1 出貨單 - 黑白版
新增：
  - 贈品欄（商品數量表正下方，同寬黑框）
  - 右上角紙箱欄位（預留位置，資料來源待接）
  - 右下整合資訊框（VIP / 購物金 / 已購買次數 / 均價）
沿用 V4 修正：
  1. 無備註 → 不顯示備註框（完全空白，不占位）
  2. 有備註 → 顯示淺灰底框（同前）
  3. 金額右側：label 與數字拉近（同欄位緊靠）
"""

import io, os, subprocess, tempfile

from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.lib.colors import HexColor
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from PIL import Image

# ─────────────────────────────────────────
# 座標常數
# ─────────────────────────────────────────
A4_W, A4_H = 595, 842

TEAR_PDF_TOP    = 30.0
TEAR_PDF_BOTTOM = 365.7
TEAR_X0 = 160.6
TEAR_X1 = 434.4
TEAR_RL_TOP    = A4_H - TEAR_PDF_TOP
TEAR_RL_BOTTOM = A4_H - TEAR_PDF_BOTTOM
TEAR_W = TEAR_X1 - TEAR_X0
TEAR_H = TEAR_RL_TOP - TEAR_RL_BOTTOM

RDZ_TOP    = TEAR_RL_BOTTOM
RDZ_BOTTOM = A4_H - 808
RDZ_X0     = 24
RDZ_X1     = 572
RDZ_W      = RDZ_X1 - RDZ_X0

BADGE_Y_RL = A4_H - 58

# ─────────────────────────────────────────
# 12 格子
# ─────────────────────────────────────────
GRID_ROWS = [
    ["C",  "F",  "C潤", "F潤", "抗痘沐", "水光沐"],
    ["橘", "綠", "痘乳", "白乳", "早C",   "晚A"],
]

# 分組統計定義
SUMMARY_GROUPS = [
    ("大",   ["C", "F", "C潤", "F潤"]),
    ("沐",   ["抗痘沐", "水光沐"]),
    ("小",   ["橘", "綠"]),
    ("乳",   ["痘乳", "白乳"]),
    ("幕斯", ["早C", "晚A"]),
]

def calc_summary(qty):
    parts = []
    for name, cats in SUMMARY_GROUPS:
        total = sum(qty.get(c, 0) for c in cats)
        if total > 0:
            parts.append((total, name))
    return parts

# ─────────────────────────────────────────
# 色盤（黑白）
# ─────────────────────────────────────────
BW_BLACK     = colors.black
BW_WHITE     = colors.white
BW_GRAY_MID  = HexColor("#666666")
BW_GRAY_LT   = HexColor("#aaaaaa")
BW_GRAY_BG   = HexColor("#efefef")
BW_GRAY_DARK = HexColor("#333333")

# ─────────────────────────────────────────
# 字型
# ─────────────────────────────────────────
_LOADED = False

def _find_cjk_font() -> str:
    """自動尋找可用的中文字型（Windows / Linux / macOS）"""
    import sys as _sys
    candidates = []

    # === 部署環境優先：先檢查 repo 根目錄的字型 ===
    # 注意：reportlab 不支援 NotoSansCJK 的 OTF/TTC（postscript outlines）
    # 必須用 TTF 格式（TrueType outlines / Variable Font）
    project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    for fname in ("NotoSansTC.ttf", "NotoSansTC-Bold.ttf", "NotoSansTC-Regular.ttf", "NotoSansCJK-Bold.ttc"):
        p = os.path.join(project_root, fname)
        if os.path.exists(p):
            return p

    if _sys.platform == "win32":
        # Windows：優先使用微軟正黑體 / 新細明體 / 標楷體
        win_root = os.environ.get("SystemRoot", "C:\\Windows")
        candidates = [
            os.path.join(win_root, "Fonts", "msjh.ttc"),       # 微軟正黑體
            os.path.join(win_root, "Fonts", "msjhbd.ttc"),
            os.path.join(win_root, "Fonts", "mingliu.ttc"),     # 新細明體
            os.path.join(win_root, "Fonts", "kaiu.ttf"),        # 標楷體
            os.path.join(win_root, "Fonts", "msyh.ttc"),        # 微軟雅黑
            os.path.join(win_root, "Fonts", "simsun.ttc"),      # 宋體
        ]
    elif _sys.platform == "darwin":
        candidates = [
            "/System/Library/Fonts/PingFang.ttc",
            "/Library/Fonts/Arial Unicode.ttf",
        ]
    else:
        # Linux
        candidates = [
            "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc",
            "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
        ]

    for path in candidates:
        if os.path.exists(path):
            return path
    return ""

def setup_font():
    global _LOADED
    if _LOADED: return "CJK", "CJK-Bold"
    font_path = _find_cjk_font()
    if font_path:
        try:
            pdfmetrics.registerFont(TTFont("CJK",      font_path, subfontIndex=0))
            pdfmetrics.registerFont(TTFont("CJK-Bold", font_path, subfontIndex=0))
            _LOADED = True
            return "CJK", "CJK-Bold"
        except Exception:
            pass
        # 部分字型不支援 subfontIndex，嘗試不帶 index
        try:
            pdfmetrics.registerFont(TTFont("CJK",      font_path))
            pdfmetrics.registerFont(TTFont("CJK-Bold", font_path))
            _LOADED = True
            return "CJK", "CJK-Bold"
        except Exception as e:
            print(f"⚠ 字型載入失敗（{font_path}）：{e}")
    print("⚠ 未找到中文字型，使用 Helvetica（中文可能無法顯示）")
    return "Helvetica", "Helvetica-Bold"

# ─────────────────────────────────────────
# 繪圖工具
# ─────────────────────────────────────────
def fill_rect(c, x, y, w, h, fc, sc=None, sw=0.5):
    c.saveState()
    c.setFillColor(fc)
    if sc:
        c.setStrokeColor(sc); c.setLineWidth(sw)
        c.rect(x, y, w, h, fill=1, stroke=1)
    else:
        c.rect(x, y, w, h, fill=1, stroke=0)
    c.restoreState()

def txt(c, s, x, y, f, sz, col=None, align="left"):
    col = col or BW_BLACK
    c.saveState()
    c.setFont(f, sz); c.setFillColor(col)
    {"left": c.drawString, "right": c.drawRightString,
     "center": c.drawCentredString}[align](x, y, s)
    c.restoreState()

def hline(c, x0, x1, y, w=0.5, col=None, dash=None):
    col = col or BW_BLACK
    c.saveState()
    c.setStrokeColor(col); c.setLineWidth(w)
    if dash: c.setDash(*dash)
    c.line(x0, y, x1, y)
    c.restoreState()

def rect_out(c, x, y, w, h, lw=0.5, col=None, dash=None):
    col = col or BW_BLACK
    c.saveState()
    c.setStrokeColor(col); c.setLineWidth(lw)
    if dash: c.setDash(*dash)
    c.rect(x, y, w, h, fill=0, stroke=1)
    c.restoreState()

# ─────────────────────────────────────────
# 撕線區 rasterize
# ─────────────────────────────────────────
def extract_tear_image(pdf_path, page_index=0, dpi=200):
    """
    Rasterize 原 PDF 撕線區。
    使用 PyMuPDF（pip install pymupdf）—— 跨平台，Windows / Linux / macOS / 網路版皆可用。
    不需要外部工具（pdftoppm、poppler 等）。
    新版 pymupdf 的 import 名稱因平台不同：
      Windows 新版：import pymupdf
      Linux / 舊版：import fitz
    """
    # 相容新舊版 pymupdf
    try:
        import pymupdf as fitz
    except ImportError:
        import fitz

    pt2px = dpi / 72.0
    doc  = fitz.open(pdf_path)
    page = doc[page_index]
    mat  = fitz.Matrix(pt2px, pt2px)
    pix  = page.get_pixmap(matrix=mat, colorspace=fitz.csRGB)
    img  = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    doc.close()
    return img.crop((
        int(TEAR_X0 * pt2px), int(TEAR_PDF_TOP * pt2px),
        int(TEAR_X1 * pt2px), int(TEAR_PDF_BOTTOM * pt2px),
    ))

# draw_badge 已整合進 draw_amount（左側框格區塊）
def draw_badge(c, FB, count):
    pass  # no-op: badge is now drawn inside draw_amount

# ─────────────────────────────────────────
# 2.2.1 訂單資訊區
# Fix 1: 無備註不顯示框；有備註才顯示灰底框
# ─────────────────────────────────────────
def draw_order_info(c, F, FB, order, top_y, zx, zw):
    pad = 8
    lx = zx + pad
    rx = zx + zw - pad

    # 頂線
    hline(c, zx, zx + zw, top_y, w=1.2)

    # 訂單編號（左）+ 訂購日期（右），同一行
    txt(c, f"#{order['order_id']}", lx, top_y - 22, FB, 14)
    txt(c, f"訂購日期：{order['order_time']}", rx, top_y - 22, FB, 11, align="right")

    # 分隔線（往上提，因為移除了品牌行）
    sep = top_y - 36
    hline(c, zx, zx + zw, sep, w=0.4, col=BW_GRAY_LT)

    # ── 智慧型版面 ──
    # 同名 → 一行：收件人 + 電話 + 物流
    # 不同名 → 兩行：(行1) 訂購人 + 收貨人 + 電話　　(行2) 物流
    buyer     = (order.get('buyer')     or '').strip()
    recipient = (order.get('recipient') or '').strip()
    phone     = (order.get('phone')     or '').strip()
    shipping_method = (order.get('shipping_method') or '').strip()

    if buyer and recipient and buyer != recipient:
        line1 = f"訂購人:{buyer}　　收貨人:{recipient}　　電話:{phone}"
        txt(c, line1, lx, sep - 17, FB, 11)
        if shipping_method:
            txt(c, f"物流：{shipping_method}", lx, sep - 35, FB, 11)
            info_bot = sep - 35
        else:
            info_bot = sep - 17
    else:
        name_only = recipient or buyer
        parts = [f"收件人：{name_only}", f"電話：{phone}"]
        if shipping_method:
            parts.append(f"物流：{shipping_method}")
        txt(c, "　　".join(parts), lx, sep - 17, FB, 11)
        info_bot = sep - 17

    # ── 備註框：固定高度（可容納 3 行），無備註也佔位留白 ──
    REMARK_MAX_LINES = 3
    LINE_H   = 18      # 行高
    PAD_TOP  = 10      # 框內上邊距
    PAD_BOT  = 8       # 框內下邊距
    rem_h    = PAD_TOP + REMARK_MAX_LINES * LINE_H + PAD_BOT   # = 64
    rem_y    = info_bot - 15   # 框頂部（reportlab y），相對於最後一行

    remark = (order.get('remark') or '').strip()

    if remark:
        # 截斷到最多 3 行，每行約 38 字
        max_ch = 38
        if len(remark) > max_ch * REMARK_MAX_LINES:
            remark = remark[:max_ch * REMARK_MAX_LINES - 3] + "..."
        lines = []
        while len(remark) > max_ch:
            lines.append(remark[:max_ch]); remark = remark[max_ch:]
        if remark: lines.append(remark)
        lines = lines[:REMARK_MAX_LINES]

        fill_rect(c, zx, rem_y - rem_h, zw, rem_h,
                  BW_GRAY_BG, sc=BW_GRAY_LT, sw=0.4)
        txt(c, "訂單備註：", lx, rem_y - PAD_TOP - 2, F, 10, col=BW_GRAY_MID)
        for i, ln in enumerate(lines):
            txt(c, ln, lx + 66, rem_y - PAD_TOP - 2 - i * LINE_H, F, 11)
    else:
        # 無備註：只畫空框（或不畫，空白佔位）
        pass   # 不畫框，空間留著，下方不上移

    return rem_y - rem_h   # 固定返回同樣底部位置

# ─────────────────────────────────────────
# 2.2.0.5 紙箱欄位（V4.1 預留）
# 位置：右上角，QR 碼右側
# 規格：資料來源待接（目前只畫格式，內容從 order.get('box_type') 來）
# ─────────────────────────────────────────
def draw_box_type_section(c, F, FB, order):
    """右上角紙箱顯示。資料未接時不畫任何東西。"""
    box_type = (order.get('box_type') or '').strip()
    if not box_type:
        return  # 暫未接資料，不畫
    # 位置常數（等接資料時再微調）
    BOX_X = 470
    BOX_Y = A4_H - 200    # 撕線區右側
    txt(c, "紙箱:", BOX_X, BOX_Y, F, 11)
    txt(c, box_type, BOX_X, BOX_Y - 22, FB, 18)

# ─────────────────────────────────────────
# 2.2.2 商品格子 + 贈品區
# Fix 3: 贈品區白底
# ─────────────────────────────────────────
def draw_grid(c, F, FB, order, top_y, zx, zw):
    qty     = order['qty']
    grid_w  = zw          # V4.1: 商品格佔滿整寬
    cell_w  = grid_w / 6
    from reportlab.pdfbase.pdfmetrics import stringWidth

    name_h   = 26     # 12pt 名稱格高
    num_h    = 40     # 數字格高
    TITLE_H  = 32     # 標題列高（加高，容納統計文字）
    grid_h   = TITLE_H + 2 * (name_h + num_h)

    # 分組統計（先算，用於標題列右側）
    parts   = calc_summary(qty)
    NUM_SZ  = 20      # 統計數字大小（V4.1 加大）
    NAME_SZ = 13      # 統計組名大小（V4.1 加大）
    SEP     = 12      # token 間距

    # 商品格子外框
    rect_out(c, zx, top_y - grid_h, grid_w, grid_h, lw=1.0)
    title_bot = top_y - TITLE_H

    # 標題列：「商品數量」靠左，統計靠右，垂直置中於標題列
    title_mid_y = title_bot + (TITLE_H - 13) / 2   # 「商品數量」13pt 垂直置中
    txt(c, "商品數量", zx + 10, title_mid_y, FB, 13)

    if parts:
        # 從右側往左排，先算總寬
        total_w = 0
        for i, (n, nm) in enumerate(parts):
            total_w += stringWidth(str(n), FB, NUM_SZ)
            total_w += stringWidth(nm,     F,  NAME_SZ)
            if i < len(parts) - 1:
                total_w += SEP

        cur_x = zx + grid_w - 10 - total_w   # 右邊距 10pt
        # 數字與組名垂直置中
        num_y  = title_bot + (TITLE_H - NUM_SZ)  / 2
        name_y = title_bot + (TITLE_H - NAME_SZ) / 2 + 1

        for i, (n, nm) in enumerate(parts):
            ns  = str(n)
            nw  = stringWidth(ns, FB, NUM_SZ)
            nmw = stringWidth(nm, F,  NAME_SZ)
            txt(c, ns, cur_x,           num_y,  FB, NUM_SZ)
            txt(c, nm, cur_x + nw + 4, name_y, F,  NAME_SZ)  # +4 間距
            cur_x += nw + nmw + SEP

    hline(c, zx, zx + grid_w, title_bot, w=0.8, col=BW_GRAY_LT)

    for ri, row_cats in enumerate(GRID_ROWS):
        ry = title_bot - ri * (name_h + num_h)
        for ci, cat in enumerate(row_cats):
            cx = zx + ci * cell_w
            fill_rect(c, cx, ry - name_h, cell_w, name_h,
                      BW_GRAY_BG, sc=BW_GRAY_LT, sw=0.5)
            txt(c, cat, cx + cell_w/2,
                ry - name_h + (name_h - 13)//2, FB, 13, align="center")
            fill_rect(c, cx, ry - name_h - num_h, cell_w, num_h,
                      BW_WHITE, sc=BW_GRAY_LT, sw=0.5)
            n = qty.get(cat, 0)
            if n > 0:
                txt(c, str(n), cx + cell_w/2,
                    ry - name_h - num_h + 10, FB, 22, align="center")

    return top_y - grid_h

# ─────────────────────────────────────────
# 2.2.2.5 贈品欄（V4.1 新增）
# 規格：
#   - 商品數量表正下方
#   - 同 zw 寬、黑框、無填色
#   - 左上角「贈品:」標題
#   - 內容由 main.py 透過 gift_parser 預先組好，放在 order['gift_display']
#   - 沒贈品時仍畫外框（內容空白）— 維持版面一致
# ─────────────────────────────────────────
def draw_gift_section(c, F, FB, order, top_y, zx, zw):
    """贈品欄：有黑色外框，左上角「贈品:」標題，下方顯示贈品內容"""
    GIFT_H     = 56          # V4.1 略加高，容納加大的字
    PAD_X      = 10
    PAD_TOP    = 8
    LABEL_SZ   = 13          # V4.1 加大
    CONTENT_SZ = 15          # V4.1 加大

    # 黑色外框（同寬黑框）
    rect_out(c, zx, top_y - GIFT_H, zw, GIFT_H, lw=1.0)

    # 左上角「贈品:」標題（加粗加大）
    txt(c, "贈品:", zx + PAD_X, top_y - PAD_TOP - LABEL_SZ + 2, FB, LABEL_SZ)

    # 內容（加粗加大）
    content = (order.get('gift_display') or '').strip()
    if content:
        txt(c, content, zx + PAD_X,
            top_y - PAD_TOP - LABEL_SZ - 18,
            FB, CONTENT_SZ)

    return top_y - GIFT_H

# ─────────────────────────────────────────
# 2.2.3 金額區 + VIP 整合框（V4.1.1）
# 規格：
#   - VIP 框底邊錨點對齊頁面底部（bottom_y 參數）
#   - 金額表從 VIP 框上方往上堆疊
#   - 整合框：第一行 左 ★VIP / 右 購物金；第二行 左 均價 / 右 已購買次數
# ─────────────────────────────────────────
def draw_amount(c, F, FB, order, bottom_y, zx, zw):
    pad = 8

    def fmt_val(v):
        if isinstance(v, float) and v == int(v):
            v = int(v)
        if isinstance(v, (int, float)):
            return (f"-${abs(v):,}" if v < 0 else f"${v:,}")
        return str(v)

    val_x  = zx + zw - pad
    NUM_W  = 80
    lbl_rx = val_x - NUM_W - 4

    rows = []
    rows.append(("總計",     order['subtotal'],         False))
    if order.get('product_discount', 0) != 0:
        rows.append(("商品折扣", order['product_discount'], False))
    if order.get('order_discount', 0) != 0:
        rows.append(("訂單折扣", order['order_discount'],   False))
    rows.append(("運費",     order['shipping'],          False))
    if order.get('adjustment', 0) != 0:
        rows.append(("調整",   order['adjustment'],        False))
    rows.append(None)
    rows.append(("訂單金額", order['total'],              True))

    # V4.1.1: 縮回 V4 原值,讓 4~5 行金額表 + VIP 框塞得下頁面底部
    row_h    = 16
    sep_h    = 12
    BOX_H    = 40
    BOX_GAP  = 4

    # 反推 top_y:bottom_y 是 VIP 框底邊
    n_data  = sum(1 for r in rows if r is not None)
    n_sep   = sum(1 for r in rows if r is None)
    total_h = pad + n_data * row_h + n_sep * sep_h + BOX_GAP + BOX_H
    top_y   = bottom_y + total_h

    # 繪製金額表
    cur_y = top_y - pad
    for row in rows:
        if row is None:
            hline(c, lbl_rx - 60, val_x, cur_y - 4, w=0.7, col=BW_GRAY_LT)
            cur_y -= sep_h
            continue
        label, value, bold = row
        if bold:
            fn_l, fs_l = FB, 13
            fn_v, fs_v = FB, 13
        else:
            fn_l, fs_l = F,  11
            fn_v, fs_v = FB, 11
        txt(c, label,         lbl_rx, cur_y - row_h + 4, fn_l, fs_l, align="right")
        txt(c, fmt_val(value), val_x, cur_y - row_h + 4, fn_v, fs_v, align="right")
        cur_y -= row_h

    # VIP 整合框（獨立規則:縮到右下角,寬度依「最滿情況」設計）
    # 最滿情況:★★ SVIP / 均價 $1,234.5 / 購物金 $9,999 / 已購買次數 99
    # 左半上下行、右半上下行皆「頭部對齊」(都用 align=left,固定 head x)
    line_size = 13
    from reportlab.pdfbase.pdfmetrics import stringWidth
    LEFT_MAX  = max(stringWidth("★★ SVIP",       FB, line_size),
                    stringWidth("均價 $1,234.5",  FB, line_size))
    RIGHT_MAX = max(stringWidth("購物金 $9,999",  FB, line_size),
                    stringWidth("已購買次數 99",  FB, line_size))
    MID_GAP     = 12
    inner_pad_x = 8
    BOX_W = inner_pad_x*2 + LEFT_MAX + MID_GAP + RIGHT_MAX
    BOX_X = zx + zw - BOX_W
    box_y = bottom_y

    c.saveState()
    c.setStrokeColor(BW_BLACK)
    c.setLineWidth(1.2)
    c.rect(BOX_X, box_y, BOX_W, BOX_H, fill=0, stroke=1)
    c.restoreState()

    left_head_x  = BOX_X + inner_pad_x
    right_head_x = BOX_X + BOX_W - inner_pad_x - RIGHT_MAX

    line1_y = box_y + BOX_H - 6 - line_size
    line2_y = box_y + 5

    vip    = order.get('vip_level', '')
    wallet = order.get('wallet')
    avg    = order.get('avg_price')
    cnt    = order.get('purchase_count', 0)

    # 左半上下行皆 align=left 於 left_head_x
    if vip in ("VIP", "SVIP"):
        stars = "★" if vip == "VIP" else "★★"
        txt(c, f"{stars} {vip}", left_head_x, line1_y, FB, line_size)
    if avg is not None:
        txt(c, f"均價 ${avg:.1f}", left_head_x, line2_y, FB, line_size)

    # 右半上下行皆 align=left 於 right_head_x(頭部對齊)
    if wallet:
        txt(c, f"購物金 ${wallet:,}", right_head_x, line1_y, FB, line_size)
    if cnt and cnt > 0:
        txt(c, f"已購買次數 {cnt}", right_head_x, line2_y, FB, line_size)

# ─────────────────────────────────────────
# 主繪製
# ─────────────────────────────────────────
def generate_page(pdf_path, page_index, order, output_path):
    F, FB = setup_font()
    tear_img = extract_tear_image(pdf_path, page_index=page_index, dpi=200)

    c = canvas.Canvas(output_path, pagesize=(A4_W, A4_H))
    fill_rect(c, 0, 0, A4_W, A4_H, BW_WHITE)

    c.drawImage(ImageReader(tear_img),
                TEAR_X0, TEAR_RL_BOTTOM,
                width=TEAR_W, height=TEAR_H,
                preserveAspectRatio=False, mask="auto")

    if order.get('purchase_count', 0) > 0:
        draw_badge(c, FB, order['purchase_count'])

    zx, zw = RDZ_X0, RDZ_W
    GAP = 10

    # V4.1.1 版面：上方左半 70% 放訂單編號/商品表/贈品欄;
    # 金額表 + VIP 框移到頁面下方全寬(右側上半留白)。
    zw_left = zw * 0.70

    # 左半 70%:訂單編號 → 商品表 → 贈品欄
    info_bot = draw_order_info(c, F, FB, order, RDZ_TOP, zx, zw_left)
    grid_bot = draw_grid(c, F, FB, order, info_bot - GAP, zx, zw_left)
    gift_bot = draw_gift_section(c, F, FB, order, grid_bot - GAP, zx, zw_left)

    # 下方全寬:金額表 + VIP 整合框(VIP 框底邊釘在頁面最下方,金額表往上堆疊)
    draw_amount(c, F, FB, order, RDZ_BOTTOM + 2, zx, zw)

    # 右上角紙箱(預留位置不動)
    draw_box_type_section(c, F, FB, order)

    c.showPage()
    c.save()
    print(f"✅ {output_path}")

# ─────────────────────────────────────────
# 測試資料
# ─────────────────────────────────────────

# 無備註版（使用購物金 + 均價 + SVIP）
ORDER_NO_REMARK = {
    "order_id": "WVEEGJ65B", "print_time": "2026/05/07 13:53:15",
    "order_time": "2026/05/07 13:08:16", "invoice_time": None,
    "recipient": "鄭又禎", "phone": "0976-207-902",
    "remark": "",
    "vip_level": "SVIP", "purchase_count": 3,
    "wallet": 200, "avg_price": 828.0,
    "subtotal": 1576, "product_discount": -200, "order_discount": 0,
    "shipping": 80, "adjustment": 0, "total": 1456,
    "qty": {"C":2,"F":0,"C潤":0,"F潤":1,"抗痘沐":1,"水光沐":0,
            "橘":0,"綠":0,"痘乳":0,"白乳":0,"早C":0,"晚A":0},
}

# 有備註版（使用購物金 + 均價 + VIP）
ORDER_WITH_REMARK = {
    "order_id": "W6AAJ6BVK", "print_time": "2026/05/07 13:53:15",
    "order_time": "2026/05/06 22:13:43", "invoice_time": "2026/05/06 23:00:00",
    "recipient": "陳惠珠", "phone": "0933-325-409",
    "remark": "請5/11前送達，因5/13出國，怕會耽誤到取貨時間，謝謝您",
    "vip_level": "VIP", "purchase_count": 7,
    "wallet": 150, "avg_price": 912.5,
    "subtotal": 5500, "product_discount": -2248, "order_discount": -93,
    "shipping": 0, "adjustment": 0, "total": 3009,
    "qty": {"C":0,"F":0,"C潤":0,"F潤":0,"抗痘沐":0,"水光沐":0,
            "橘":2,"綠":2,"痘乳":0,"白乳":0,"早C":0,"晚A":2},
}

if __name__ == "__main__":
    import os
    os.makedirs("/mnt/user-data/outputs", exist_ok=True)
    pdf_path = "/mnt/user-data/uploads/050701new.pdf"

    generate_page(pdf_path, 0,  ORDER_NO_REMARK,  "/mnt/user-data/outputs/v4_no_remark.pdf")
    generate_page(pdf_path, 97, ORDER_WITH_REMARK, "/mnt/user-data/outputs/v4_with_remark.pdf")
