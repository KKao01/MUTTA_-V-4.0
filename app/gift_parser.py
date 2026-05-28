"""
贈品標籤解析器 - MUTTA 分單系統 V4.1

職責:
  1. 從訂單備註抽出 #MG: 贈品標籤,解析成 [(gift_id, qty), ...]
  2. 清除備註中的「贈品自動填入段落」(人類段落 + #MG: 標籤行)
  3. 根據 dev 後台維護的「ID → 顯示片段」對應表,組裝贈品顯示文字

來源:Gift Builder (mutta_gift_builder) 在客人結帳時把以下兩段塞進備註:
  - 人類段落(夾在兩條 ────── 之間)
  - 機器標籤(獨立一行,前綴 #MG:)
"""
import re


# 人類段落的固定上下分隔線
SEPARATOR_LINE = "──────────────────"


def extract_gifts(remark: str) -> list:
    """
    從訂單備註抽出 [(gift_id, qty), ...]

    範例:
        remark 含 "#MG:gift1*3,gift2*2"
        → [('gift1', 3), ('gift2', 2)]

    沒贈品(沒有 #MG: 或格式錯)時回傳 []。
    """
    if not remark:
        return []

    # #MG: 後接逗號分隔的「gift_id*qty」直到遇到空白或行尾
    m = re.search(r'#MG:\s*([^\s]+)', remark)
    if not m:
        return []

    gifts = []
    for token in m.group(1).split(','):
        token = token.strip()
        if not token or '*' not in token:
            continue
        gid, qty_str = token.split('*', 1)
        gid = gid.strip()
        if not gid:
            continue
        try:
            qty = int(qty_str.strip())
            if qty <= 0:
                continue
        except ValueError:
            continue
        gifts.append((gid, qty))
    return gifts


def remove_gift_block(remark: str) -> str:
    """
    移除整段贈品資訊(人類段落 + #MG 標籤),回傳乾淨備註

    1. 刪除「────── 🎁 自選贈品 ... ──────」整段
    2. 刪除「#MG:xxx」那一行
    3. 收頭去尾、把連續空行壓成單一空行

    如果客人原本沒寫備註(只剩贈品段落),清完後會是空字串。
    """
    if not remark:
        return ""

    # 1. 刪除人類友善段落(兩條分隔線之間 + 🎁 自選贈品 標題)
    pattern_block = (
        r'\n*' + re.escape(SEPARATOR_LINE) +
        r'\s*\n[^\n]*🎁[^\n]*\n.*?' +
        re.escape(SEPARATOR_LINE) +
        r'\n*'
    )
    remark = re.sub(pattern_block, '\n', remark, flags=re.DOTALL)

    # 2. 刪除 #MG: 標籤行(整行,包含可能的前後空白)
    remark = re.sub(r'\n*[ \t]*#MG:\S*[ \t]*\n*', '\n', remark)

    # 3. 把連續多個空行壓成單一空行,並 strip
    remark = re.sub(r'\n{2,}', '\n', remark)
    return remark.strip()


def format_gift_display(
    gifts: list,
    display_map: dict,
    fixed_suffix: str = "",
    fallback_map: dict | None = None,
) -> str:
    """
    把 [(gift_id, qty), ...] 轉成顯示文字。

    串接規則(依使用者規格):
      - 每個 gift 用 + 串接(無空格)
      - qty == 1 不顯示「×N」,qty > 1 才顯示「×N」
      - 全部串完後,若 fixed_suffix 非空,再加「+ fixed_suffix」
      - gift_id 在 display_map 找不到 → fallback 到 fallback_map[gift_id]
        若 fallback_map 也沒有 → 顯示「⚠️gift_id」(出貨人員看就會去設定)
      - gifts 為空 → 回傳空字串(贈品欄不畫內容,但外框保留)

    範例:
      gifts=[("gift1", 3), ("gift2", 1)], display_map={gift1: "洗髮包", gift2: "梳"}, fixed_suffix="卡"
      → "洗髮包×3+梳+卡"

      gifts=[("gift2", 1), ("gift4", 1)], display_map={gift2: "梳", gift4: "球"}, fixed_suffix="卡"
      → "梳+球+卡"

      gifts=[("gift1", 1)], display_map={gift1: "洗髮包"}, fixed_suffix=""
      → "洗髮包"

      gifts=[("unknown", 2)], display_map={}, fallback_map={"unknown": "未命名贈品 A"}, fixed_suffix="卡"
      → "未命名贈品 A×2+卡"
    """
    if not gifts:
        return ""

    fallback_map = fallback_map or {}
    parts = []
    for gid, qty in gifts:
        # 顯示片段:對應表 > fallback 表 > ⚠️ID
        seg = display_map.get(gid) or fallback_map.get(gid) or f"⚠️{gid}"
        if qty > 1:
            seg = f"{seg}×{qty}"
        parts.append(seg)

    text = "+".join(parts)

    # 固定後綴(只在有贈品時才加)
    suffix = (fixed_suffix or "").strip()
    if suffix:
        text = f"{text}+{suffix}"

    return text


def process_order_remark(
    remark: str,
    display_map: dict,
    fixed_suffix: str = "",
    fallback_map: dict | None = None,
) -> tuple:
    """
    一站式處理:
      輸入:原始備註
      輸出:(clean_remark, gift_display, has_gift)

      clean_remark : 移除贈品段落後的備註(可能為空)
      gift_display : 贈品欄要顯示的文字(可能為空)
      has_gift     : 是否含有 #MG: 標籤(用來判斷要不要畫贈品欄外框)
    """
    gifts = extract_gifts(remark)
    has_gift = bool(gifts)

    if has_gift:
        clean_remark = remove_gift_block(remark)
        gift_display = format_gift_display(
            gifts, display_map, fixed_suffix, fallback_map
        )
    else:
        clean_remark = remark or ""
        gift_display = ""

    return clean_remark, gift_display, has_gift


# ─────────────────────────────────────────
# 自我測試
# ─────────────────────────────────────────
if __name__ == "__main__":
    test_remark = """請幫我中午前到貨

──────────────────
🎁 自選贈品(系統自動填入,請勿修改)
• 洗髮精組 × 3
• 按摩梳 × 2
──────────────────
#MG:gift1*3,gift2*2"""

    print("=== Test 1: 完整解析 ===")
    gifts = extract_gifts(test_remark)
    print(f"extract_gifts: {gifts}")

    clean = remove_gift_block(test_remark)
    print(f"clean_remark: {repr(clean)}")

    display = format_gift_display(
        gifts,
        display_map={"gift1": "洗髮包", "gift2": "梳"},
        fixed_suffix="卡"
    )
    print(f"display: {display}")

    print("\n=== Test 2: 只有贈品 ===")
    only_gift = test_remark.replace("請幫我中午前到貨", "").strip()
    clean2 = remove_gift_block(only_gift)
    print(f"clean_remark (應為空): {repr(clean2)}")

    print("\n=== Test 3: 沒贈品 ===")
    no_gift = "麻煩晚上配送"
    print(f"extract_gifts: {extract_gifts(no_gift)}")
    print(f"clean_remark: {repr(remove_gift_block(no_gift))}")

    print("\n=== Test 4: 數量=1 不加×1 ===")
    g = [("gift2", 1), ("gift4", 1)]
    print(format_gift_display(g, {"gift2": "梳", "gift4": "球"}, "卡"))
    # 預期:梳+球+卡

    print("\n=== Test 5: 未知 gift_id fallback ===")
    g = [("unknown", 2)]
    print(format_gift_display(
        g, {}, "卡",
        fallback_map={"unknown": "絲絨雲感沐浴球"}
    ))
    # 預期:絲絨雲感沐浴球×2+卡

    print("\n=== Test 6: 完全未知 ===")
    g = [("xxx", 1)]
    print(format_gift_display(g, {}, "卡"))
    # 預期:⚠️xxx+卡

    print("\n=== Test 7: process_order_remark 一站式 ===")
    r, d, h = process_order_remark(
        test_remark,
        {"gift1": "洗髮包", "gift2": "梳"},
        "卡"
    )
    print(f"clean={repr(r)}, display={d}, has_gift={h}")
