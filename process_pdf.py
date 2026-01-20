#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
处理PDF文件：从拣货单提取SKU信息并插入到物流面单
"""

import os
import re
from typing import List, Tuple, Dict, Set

import pdfplumber
import fitz  # PyMuPDF


# ========== 表格位置配置 ==========
# 手动设置表格的Y坐标（表格顶部位置）
# 单位：点（points），PDF坐标系从页面底部开始，Y值越大越靠上
# 如果设置为None，将自动搜索"RDC 01"位置
TABLE_Y_POSITION = 148  # 例如: 500 表示距离页面底部500点，或者设置为None使用自动定位

# ========== 表格尺寸配置 ==========
# 表格行数
TABLE_ROWS = 3  # 表格行数（默认3行）
# 每行高度（单位：点）
TABLE_ROW_HEIGHT = 16  # 每行高度，单位：点（points）
# ===================================

# ========== SKU 名称配置 ==========
# 从 sku_name.csv 中加载所有合法的 SKU 名称，用于辅助识别，避免把描述性文本误当成 SKU
SKU_NAME_LIST: List[str] = []
SKU_NAME_SET: Set[str] = set()


def _load_sku_name_list() -> None:
    """
    从当前目录下的 sku_name.csv 中加载 SKU 名称列表。
    每一行一个 SKU 名称。
    """
    global SKU_NAME_LIST, SKU_NAME_SET

    try:
        base_dir = os.path.dirname(os.path.abspath(__file__))
        csv_path = os.path.join(base_dir, "sku_name.csv")
        if not os.path.exists(csv_path):
            # 没有配置文件时，保持为空列表，后续逻辑会退回到旧的规则
            return

        with open(csv_path, "r", encoding="utf-8") as f:
            names = [line.strip() for line in f if line.strip()]

        # 去重，保留顺序
        seen: Set[str] = set()
        unique_names: List[str] = []
        for name in names:
            if name not in seen:
                seen.add(name)
                unique_names.append(name)

        SKU_NAME_LIST = unique_names
        SKU_NAME_SET = set(unique_names)
    except Exception:
        # 任何异常都不应中断主流程，只是放弃 SKU 白名单能力
        SKU_NAME_LIST = []
        SKU_NAME_SET = set()


# 在模块导入时尝试加载一次 SKU 名称列表
_load_sku_name_list()


def _match_complex_candidate(sku_raw: str, candidate: str) -> bool:
    """
    判断一个带括号的复杂 SKU（例如 Ink-pack-Y2K-(beige53+pink53+green53+Ins50)）
    是否可以与当前抽取的 sku_raw 对应。
    规则：
    - candidate 形如 前缀(部件1+部件2+...)
    - 要求前缀和每个部件的关键部分都能在 sku_raw 中找到：
      - 前缀：直接作为子串出现
      - 每个部件：字母部分取前3个字符，数字部分整体；二者都要能在 sku_raw 中找到
    这样既能较好地识别被换行/拆词的复杂 SKU，又能避免与其它 SKU 发生混淆。
    """
    sku_lower = sku_raw.lower()
    if "(" not in candidate or not candidate.endswith(")"):
        return False

    base, rest = candidate.split("(", 1)
    rest = rest[:-1]  # 去掉结尾的 ')'
    parts = [p for p in rest.split("+") if p]

    if base and base.lower() not in sku_lower:
        return False

    for part in parts:
        letters = "".join(ch for ch in part if ch.isalpha())
        digits = "".join(ch for ch in part if ch.isdigit())

        if letters:
            key = letters[:3].lower()  # 取前3个字母作为关键片段
            if key and key not in sku_lower:
                return False
        if digits:
            if digits not in sku_raw:
                return False

    return True


def _normalize_sku_by_whitelist(sku_raw: str) -> str:
    """
    使用 sku_name.csv 中的白名单对提取到的 SKU 名进行归一化：
    1. 完全匹配优先
    2. 其次是白名单项出现在 sku_raw 中
    3. 再其次，对复杂 SKU 使用 _match_complex_candidate 做严格匹配
    如果有多个候选同时满足，则保持原样以避免混淆。
    """
    if not SKU_NAME_LIST:
        return sku_raw

    if sku_raw in SKU_NAME_SET:
        return sku_raw

    candidates: List[str] = []

    # 子串匹配
    for name in SKU_NAME_LIST:
        if name in sku_raw or sku_raw in name:
            candidates.append(name)

    # 复杂 SKU 匹配
    for name in SKU_NAME_LIST:
        if name not in candidates and _match_complex_candidate(sku_raw, name):
            candidates.append(name)

    if len(candidates) == 1:
        return candidates[0]

    # 多个候选或无候选时，为避免误判，保留原始字符串
    return sku_raw


def extract_sku_from_packing_slip(text: str) -> List[Tuple[str, int]]:
    """
    从拣货单文本中提取SKU和数量
    返回: [(SKU名称, 数量), ...]
    """
    skus: List[Tuple[str, int]] = []
    lines = text.split('\n')

    found_header = False
    header_index = None  # 表头所在行索引
    qty_total_index = None  # "Qty Total" 所在行索引
    current_sku_parts = []
    current_qty = None
    
    for i, line in enumerate(lines):
        line = line.strip()
        if not line:
            # 空行可能表示一个SKU记录的结束
            if current_sku_parts and current_qty is not None:
                sku_name = ' '.join(current_sku_parts)
                skus.append((sku_name, current_qty))
                current_sku_parts = []
                current_qty = None
            continue
        
        # 查找表头
        if 'SKU' in line.upper() and ('QTY' in line.upper() or 'QUANTITY' in line.upper()):
            found_header = True
            header_index = i
            continue
        
        # 标记 \"Qty Total\" 行，用于后续限定 SKU 搜索区域
        if 'QTY TOTAL' in line.upper() and qty_total_index is None:
            qty_total_index = i

        # 在找到表头后，查找SKU行
        if found_header:
            # 先尝试使用 SKU 白名单做精确匹配：在本行中查找已知的 SKU 名称
            matched_in_line: List[Tuple[str, int]] = []
            if SKU_NAME_LIST:
                for sku_name_candidate in SKU_NAME_LIST:
                    if sku_name_candidate in line:
                        # 在同一行中查找该 SKU 后面的数量（如果有）
                        pattern = re.escape(sku_name_candidate) + r".*?(\d+)\s*$"
                        m = re.search(pattern, line)
                        qty_val: int = 1
                        if m:
                            try:
                                qty_val = int(m.group(1))
                            except ValueError:
                                qty_val = 1
                        matched_in_line.append((sku_name_candidate, qty_val))

            if matched_in_line:
                # 如果通过白名单匹配到了 SKU，则优先使用这些结果，并跳过后续的启发式解析
                for sku_name_candidate, qty_val in matched_in_line:
                    skus.append((sku_name_candidate, qty_val))
                continue

            # 白名单没有命中时，退回到基于文本结构的启发式解析逻辑
            # 查找包含"Ink-"的行（这是常见的SKU格式）
            if 'Ink-' in line:
                # 提取SKU和数量
                parts = line.split()

                # 找到Ink-开头的部分
                sku_start_idx = None
                for j, part in enumerate(parts):
                    if 'Ink-' in part:
                        sku_start_idx = j
                        break

                if sku_start_idx is not None:
                    # 收集SKU的所有部分（直到遇到数字）
                    sku_parts: List[str] = []
                    qty = None

                    for j in range(sku_start_idx, len(parts)):
                        part = parts[j]
                        # 检查是否是数字（数量）
                        try:
                            qty = int(part)
                            break
                        except ValueError:
                            # 不是数字，是SKU的一部分
                            sku_parts.append(part)

                    if sku_parts:
                        if qty is not None:
                            # 找到SKU和数量，但检查下一行是否还有SKU的延续部分
                            # 检查下一行（如果存在）
                            if i + 1 < len(lines):
                                next_line = lines[i + 1].strip()
                                next_parts = next_line.split()

                                # 检查下一行末尾是否有可能是SKU延续的短字符串
                                # 通常SKU延续部分在行尾，是2-5个字符的字母数字组合
                                if next_parts:
                                    # 从行尾开始检查
                                    for part in reversed(next_parts):
                                        # 如果是短字符串（2-5字符）且是纯字母或字母数字组合
                                        if 2 <= len(part) <= 5 and part.replace('-', '').replace('_', '').isalnum():
                                            # 检查是否包含常见产品描述关键词（如果是，则不是SKU延续）
                                            if not any(keyword in part.lower() for keyword in ['for', 'and', 'the', 'card', 'chip']):
                                                # 可能是SKU的延续，添加到SKU
                                                sku_parts.append(part)
                                                break
                                        # 如果遇到明显是产品描述的部分，停止
                                        elif any(keyword in part.lower() for keyword in ['sticker', 'card', 'pack', 'sheet', 'for', 'and', 'the', 'chip', 'key', 'metro']):
                                            break

                            # 找到完整的SKU和数量
                            # 将SKU部分连接，如果都是连字符分隔的格式，则用连字符连接；否则用空格
                            sku_name = '-'.join(sku_parts) if all('-' in part or part.isalnum() for part in sku_parts) else ' '.join(sku_parts)
                            # 清理多余的连字符
                            sku_name = re.sub(r'-+', '-', sku_name).strip('-')
                            skus.append((sku_name, qty))
                        else:
                            # SKU可能跨多行，保存当前部分
                            current_sku_parts = sku_parts
            else:
                # 可能是产品名称的延续行，检查是否包含数量
                parts = line.split()
                for part in parts:
                    try:
                        qty = int(part)
                        if current_sku_parts:
                            # 找到数量，完成当前SKU
                            sku_name = ' '.join(current_sku_parts)
                            skus.append((sku_name, qty))
                            current_sku_parts = []
                            current_qty = None
                        break
                    except ValueError:
                        continue
    
    # 处理最后一个SKU（如果没有遇到空行）
    if current_sku_parts and current_qty is not None:
        sku_name = ' '.join(current_sku_parts)
        skus.append((sku_name, current_qty))

    # 使用 SKU 白名单在表格区域内做一次全局兜底匹配，弥补跨行拆分导致的漏识别
    # 区域定义为：表头行(header_index) 到 \"Qty Total\" 行(qty_total_index) 之间的文本
    if SKU_NAME_LIST and header_index is not None:
        end_index = qty_total_index if qty_total_index is not None else len(lines)
        region_lines = lines[header_index + 1:end_index]
        region_text = ' '.join(l.strip() for l in region_lines if l.strip())

        existing_names = {name for name, _ in skus}
        for sku_name_candidate in SKU_NAME_LIST:
            if sku_name_candidate in existing_names:
                continue

            matched = False
            replacement_base: str = ""

            # 情况1：SKU 名在文本中连续出现
            if sku_name_candidate in region_text:
                matched = True
            else:
                # 情况2：像 Ink-pack-Y2K-(beige53+pink53+green53+Ins50) 这类，
                # 在 PDF 文本中可能被拆成多行，导致整体字符串不存在。
                if "(" in sku_name_candidate and sku_name_candidate.endswith(")"):
                    base, rest = sku_name_candidate.split("(", 1)
                    rest = rest[:-1]  # 去掉结尾的 ')'
                    parts = [p for p in rest.split("+") if p]
                    # 要求基础前缀和所有部件都在文本中出现
                    if base and base in region_text and all(p in region_text for p in parts):
                        matched = True
                        replacement_base = base

            if matched:
                # 如果这是一个更精确的 SKU（例如复杂组合 SKU），尝试替换掉之前用启发式得到的粗略 SKU
                if replacement_base:
                    skus = [
                        (name, qty)
                        for (name, qty) in skus
                        if not (name.startswith(replacement_base) and name != sku_name_candidate)
                    ]

                # 如果未找到对应数量，则默认数量为1；对于当前业务场景这是安全的
                skus.append((sku_name_candidate, 1))

    # 去重并合并相同SKU的数量
    seen = {}
    for sku, qty in skus:
        # 清理SKU名称（移除多余空格）
        sku_clean = ' '.join(sku.split())
        if sku_clean in seen:
            seen[sku_clean] += qty
        else:
            seen[sku_clean] = qty
    
    return [(sku, seen[sku]) for sku in seen]


def extract_sku_from_page(page: pdfplumber.page.Page) -> List[Tuple[str, int]]:
    """
    基于坐标从拣货单页面中提取 SKU 和数量。
    仅识别位于 \"Seller\" 左端点 x 坐标与 \"Qty\" 左端点 x 坐标之间的文字作为 SKU 名称列，
    以排除其它文字信息的干扰。
    """
    words = page.extract_words()
    if not words:
        return []

    # 1. 找到表头行中的 \"Seller\" 和 \"Qty\"
    seller_word = None
    qty_header_word = None

    for w in words:
        txt = w.get("text", "").strip()
        if txt.lower() == "seller" and seller_word is None:
            seller_word = w
        # 只把表头行的 Qty 作为列起点；后面的 \"Qty Total\" 会在不同的 y 位置
        if txt.lower().startswith("qty"):
            if qty_header_word is None:
                qty_header_word = w
            else:
                # 选择 y 值较小的那个作为表头（更靠上的那个）
                if w.get("top", 0) < qty_header_word.get("top", 0):
                    qty_header_word = w

    if not seller_word or not qty_header_word:
        # 如果无法可靠找到列边界，则退回到基于文本的解析逻辑
        text = page.extract_text() or ""
        return extract_sku_from_packing_slip(text)

    seller_x0 = float(seller_word["x0"])
    qty_x0 = float(qty_header_word["x0"])
    if qty_x0 <= seller_x0:
        # 列边界异常时，同样退回文本解析
        text = page.extract_text() or ""
        return extract_sku_from_packing_slip(text)

    # SKU 列的水平范围：Seller 左端到 Qty 左端
    sku_x_min = seller_x0
    sku_x_max = qty_x0

    # 2. 确定纵向范围：从表头行下方，到 \"Qty Total\" 行上方
    header_top = float(seller_word["top"])

    qty_total_top = None
    for w in words:
        txt = w.get("text", "").strip().lower()
        if "qty" in txt and "total" in txt:
            # 选择 y 值较大的那个，作为 \"Qty Total\" 行
            if qty_total_top is None or w["top"] > qty_total_top:
                qty_total_top = float(w["top"])

    # 纵向范围
    y_min = header_top + 1.0  # 略微避开表头行
    y_max = qty_total_top - 1.0 if qty_total_top is not None else max(w["bottom"] for w in words)

    # 3. 将每一行中位于 SKU 列范围内的单词视为 SKU 名的一部分
    #    按 top 聚类成行，再在每行内按 x0 排序
    row_tolerance = 2.0  # y 方向聚类容忍度
    # 只考虑纵向在 y_min~y_max 之间的单词
    candidate_words = [
        w for w in words
        if y_min <= (float(w["top"]) + float(w["bottom"])) / 2.0 <= y_max
    ]

    # 先按 top 排序
    candidate_words.sort(key=lambda w: w["top"])

    rows: List[List[Dict]] = []
    current_row: List[Dict] = []
    current_top: float = None  # type: ignore

    for w in candidate_words:
        top = float(w["top"])
        if current_row and current_top is not None and abs(top - current_top) > row_tolerance:
            rows.append(current_row)
            current_row = [w]
            current_top = top
        else:
            if not current_row:
                current_row = [w]
                current_top = top
            else:
                current_row.append(w)

    if current_row:
        rows.append(current_row)

    # 为每一行构建基本信息：sku_tokens 和 数量候选
    rows_info: List[Dict] = []
    for row_words in rows:
        sku_tokens: List[str] = []
        qty_candidates: List[int] = []

        for w in sorted(row_words, key=lambda ww: ww["x0"]):
            txt = w.get("text", "").strip()
            if not txt:
                continue
            # 避免将列标题或标签性的 \"Qty\" 误认为 SKU 名的一部分
            if txt.lower() == "qty":
                # 但如果这是数量列中的纯数字，在下方会单独处理
                continue
            x_center = (float(w["x0"]) + float(w["x1"])) / 2.0
            if sku_x_min <= x_center < sku_x_max:
                sku_tokens.append(txt)
            if x_center >= qty_x0 and txt.isdigit():
                try:
                    qty_candidates.append(int(txt))
                except ValueError:
                    pass

        if not sku_tokens:
            continue

        rows_info.append(
            {
                "sku_tokens": sku_tokens,
                "qty_candidates": qty_candidates,
            }
        )

    # 合并由于分行导致拆开的长 SKU 行：
    # 以包含 "Ink-" 的行为起点，向下吸收连续的非新 SKU 行，
    # 得到更完整的长 SKU 文本，避免与其它 SKU 混淆。
    merged_rows: List[Dict] = []
    i = 0
    while i < len(rows_info):
        cur = rows_info[i]
        cur_tokens = list(cur["sku_tokens"])
        qty_candidates = list(cur["qty_candidates"])
        cur_text = " ".join(cur_tokens)

        # 判断是否是一个 SKU 的起始行：包含 "Ink-" 或命中白名单前缀
        is_start = "Ink-" in cur_text
        if not is_start and SKU_NAME_LIST:
            for name in SKU_NAME_LIST:
                if name.startswith(cur_tokens[0]):
                    is_start = True
                    break

        if not is_start:
            i += 1
            continue

        j = i + 1
        while j < len(rows_info):
            nxt = rows_info[j]
            nxt_tokens = nxt["sku_tokens"]
            nxt_text = " ".join(nxt_tokens)
            lower = nxt_text.lower()

            # 碰到新 SKU 起点或明显不是 SKU 的行就停止合并
            if "Ink-" in nxt_text:
                break
            if lower.startswith(("qty total", "order id", "package id", "buyer id", "product name")):
                break

            # 否则视为当前 SKU 的续行
            cur_tokens.extend(nxt_tokens)
            j += 1

        merged_rows.append(
            {
                "sku_tokens": cur_tokens,
                "qty_candidates": qty_candidates,
            }
        )
        i = j

    result: List[Tuple[str, int]] = []

    for row in merged_rows:
        sku_tokens = row["sku_tokens"]
        qty_candidates = row["qty_candidates"]
        if not sku_tokens:
            continue

        sku_raw = " ".join(sku_tokens)

        # 使用白名单归一化复杂 SKU，避免与其它 SKU 混淆
        normalized_sku = _normalize_sku_by_whitelist(sku_raw)

        # 只保留看起来像 SKU 的行：包含 "Ink-" 或在白名单中
        if "Ink-" not in normalized_sku and normalized_sku not in SKU_NAME_SET:
            continue

        qty_val = qty_candidates[-1] if qty_candidates else 1
        result.append((normalized_sku, qty_val))

    # 如果坐标解析失败或结果为空，则退回基于文本的逻辑
    if not result:
        text = page.extract_text() or ""
        return extract_sku_from_packing_slip(text)

    return result


def find_shipping_label_position(page: fitz.Page) -> float:
    """
    获取表格应该放置的Y坐标（表格顶部位置）
    如果TABLE_Y_POSITION已设置，直接使用该值
    否则自动搜索"RDC 01"位置
    """
    # 如果手动设置了Y坐标，直接使用
    if TABLE_Y_POSITION is not None:
        return TABLE_Y_POSITION
    
    # 否则自动搜索"RDC 01"位置
    search_terms = [
        "RDC 01",
        "RDC 01",
        "RDC",
        "rdc 01",
        "rdc"
    ]
    
    for term in search_terms:
        text_instances = page.search_for(term, flags=fitz.TEXT_DEHYPHENATE)
        if text_instances:
            # 找到文本位置，返回其下方的Y坐标
            rect = text_instances[0]
            # 在文本下方留出一些间距（约20-30点）
            table_top = rect.y1 + 25
            return table_top
    
    # 如果没找到，使用页面中部偏上（约55%位置）
    return page.rect.height * 0.55


def create_sku_table(doc: fitz.Document, page_num: int, skus: List[Tuple[str, int]]):
    """
    在指定页面上创建SKU表格
    """
    page = doc[page_num]
    
    # 限制SKU数量为6个
    display_skus = skus[:6]
    if len(skus) > 6:
        display_skus.append(("Check More", ""))
    
    # 找到表格位置
    table_y = find_shipping_label_position(page)
    
    # 获取页面尺寸
    page_width = page.rect.width
    page_height = page.rect.height
    
    # 计算表格位置和尺寸
    # 表格不超出两边的黑线（假设黑线在页面边缘10%处）
    margin_left = page_width * 0.1
    margin_right = page_width * 0.1
    table_width = page_width - margin_left - margin_right
    table_x = margin_left
    
    # 表格高度（使用配置的行数和行高）
    row_height = TABLE_ROW_HEIGHT  # 每行高度
    table_height = TABLE_ROWS * row_height  # 总高度 = 行数 × 每行高度
    
    # 确保表格不覆盖地址（地址通常在页面下部）
    # 地址通常3-4行，每行约15点，所以需要预留约60-80点
    address_area = page_height * 0.15  # 底部15%用于地址
    if table_y + table_height > page_height - address_area:
        # 如果表格会覆盖地址，向上调整位置
        table_y = page_height - address_area - table_height - 10
    
    # 列宽：SKU名称占更多空间，数量列较窄
    # 每行2个SKU，每个SKU占2列（名称+数量）
    col_width_sku = table_width * 0.4  # SKU名称列宽40%
    col_width_qty = table_width * 0.1  # 数量列宽10%
    
    # 创建表格
    y_pos = table_y
    
    # 绘制表格外边框（内部分隔线在计算布局后再画）
    rect = fitz.Rect(table_x, y_pos, table_x + table_width, y_pos + table_height)
    page.draw_rect(rect, color=(0, 0, 0), width=1)
    
    # 填充数据
    # 字号范围
    base_font_size = 9  # 初始字号
    max_font_size = 11
    min_font_size = 6   # 最小字号（当仍溢出时，启用“独占一行并合并三格”规则）

    # 预判哪些 SKU 在单元格内会溢出（即使使用最小字号）
    def _will_overflow_single_cell(name: str) -> bool:
        display = name[4:] if name.startswith("Ink-") else name
        width = fitz.get_text_length(display, fontname="helv", fontsize=min_font_size)
        available = col_width_sku - 6  # 左右留出边距
        return width > available

    long_flags = [_will_overflow_single_cell(name) for name, _ in display_skus]

    # 计算文本垂直居中偏移
    text_y_offset = row_height / 2 + base_font_size / 3

    # 根据长/短 SKU 布局行：
    # - 短 SKU：一行最多放 2 个（左右各一个）
    # - 长 SKU：独占一行，合并左侧 3 个单元格显示名称，数量放在最右数量列
    rows_layout: List[List[Tuple[str, int]]] = []  # 每行是 [(name, qty), ...]，长度可为1或2
    row_types: List[str] = []  # "merged" or "normal"
    i = 0
    while i < len(display_skus) and len(rows_layout) < TABLE_ROWS:
        name, qty = display_skus[i]
        is_long = long_flags[i]

        if is_long:
            rows_layout.append([(name, qty)])
            row_types.append("merged")
            i += 1
        else:
            # 普通行，尽量放两个短 SKU
            row: List[Tuple[str, int]] = [(name, qty)]
            i += 1
            if i < len(display_skus) and not long_flags[i]:
                row.append(display_skus[i])
                i += 1
            rows_layout.append(row)
            row_types.append("normal")

    # 补齐行类型列表长度到 TABLE_ROWS，未使用的行按普通行处理，用于绘制完整网格
    full_row_types: List[str] = []
    for idx in range(TABLE_ROWS):
        if idx < len(row_types):
            full_row_types.append(row_types[idx])
        else:
            full_row_types.append("normal")

    # 根据行类型绘制内部网格线：
    # - 水平线：始终画，分隔各行
    # - 垂直线：
    #   - 外边框两根已由 draw_rect 完成
    #   - 内部三根：
    #       * 对于合并行(merged)，在该行高度范围内，跳过前两根（合并前三个单元格）
    #       * 最后一根始终绘制，用于分隔合并单元格与数量列
    # 列边界 x 坐标
    v_x1 = table_x + col_width_sku                       # 第1条内部竖线（name1|qty1）
    v_x2 = table_x + col_width_sku + col_width_qty       # 第2条内部竖线（qty1|name2）
    v_x3 = table_x + col_width_sku + col_width_qty + col_width_sku  # 第3条内部竖线（name2|qty2）

    # 水平线
    for i in range(1, TABLE_ROWS):  # TABLE_ROWS-1条水平线
        y_line = table_y + i * row_height
        page.draw_line(
            fitz.Point(table_x, y_line),
            fitz.Point(table_x + table_width, y_line),
            color=(0, 0, 0),
            width=0.5,
        )

    # 垂直线按行分段绘制
    for row_idx in range(TABLE_ROWS):
        seg_top = table_y + row_idx * row_height
        seg_bot = seg_top + row_height
        is_merged = full_row_types[row_idx] == "merged"

        # 第1条内部竖线：如果该行为合并行，则跳过
        if not is_merged:
            page.draw_line(
                fitz.Point(v_x1, seg_top),
                fitz.Point(v_x1, seg_bot),
                color=(0, 0, 0),
                width=0.5,
            )

        # 第2条内部竖线：如果该行为合并行，则跳过
        if not is_merged:
            page.draw_line(
                fitz.Point(v_x2, seg_top),
                fitz.Point(v_x2, seg_bot),
                color=(0, 0, 0),
                width=0.5,
            )

        # 第3条内部竖线：始终绘制，用于分隔第二个 SKU 名与数量列（或合并行与数量列）
        page.draw_line(
            fitz.Point(v_x3, seg_top),
            fitz.Point(v_x3, seg_bot),
            color=(0, 0, 0),
            width=0.5,
        )

    # 绘制每一行
    for row_idx, row in enumerate(rows_layout):
        y_pos = table_y + row_idx * row_height + text_y_offset

        if row_types[row_idx] == "merged":
            # 独占一行的长 SKU：名称占用前三个单元格，数量在最右数量列
            name, qty = row[0]
            if name:
                display_name = name[4:] if name.startswith("Ink-") else name
                merged_width = col_width_sku * 2 + col_width_qty  # 前三格总宽度
                available_width = merged_width - 6

                # 从最大字号往下试，直到不溢出或达到最小字号
                optimal_font_size = base_font_size
                for test_size in range(max_font_size, min_font_size - 1, -1):
                    text_width = fitz.get_text_length(display_name, fontname="helv", fontsize=test_size)
                    if text_width <= available_width:
                        optimal_font_size = test_size
                        break

                # 如果在最小字号仍然溢出，则截断并加省略号
                if optimal_font_size == min_font_size:
                    test_width = 0
                    max_chars = 0
                    for ch in display_name:
                        ch_width = fitz.get_text_length(ch, fontname="helv", fontsize=min_font_size)
                        if test_width + ch_width <= available_width:
                            test_width += ch_width
                            max_chars += 1
                        else:
                            break
                    if max_chars < len(display_name):
                        display_name = (display_name[:max_chars-3] + "...") if max_chars > 3 else display_name[:max_chars]

                # 名称起点为表格左边 + 左边距
                x_name = table_x + 3
                try:
                    page.insert_text(
                        fitz.Point(x_name, y_pos),
                        display_name,
                        fontsize=optimal_font_size,
                        color=(0, 0, 0),
                        fontname="helv",
                    )
                except Exception:
                    pass

            # 数量画在最右一列下面
            if qty:
                qty_text = str(qty)
                # 最右数量列左端 x：table_x + col_width_sku + col_width_qty + col_width_sku
                x_qty = table_x + col_width_sku + col_width_qty + col_width_sku
                try:
                    page.insert_text(
                        fitz.Point(x_qty, y_pos),
                        qty_text,
                        fontsize=base_font_size,
                        color=(0, 0, 0),
                        fontname="helv",
                    )
                except Exception:
                    pass
        else:
            # 普通行：最多两个 SKU，左右各一个
            x_start = table_x + 3
            for col_idx, (name, qty) in enumerate(row):
                x_pos = x_start + col_idx * (col_width_sku + col_width_qty)

                if name:
                    display_name = name[4:] if name.startswith("Ink-") else name
                    available_width = col_width_sku - 6

                    # 自动调整字号以适应单元格宽度
                    optimal_font_size = base_font_size
                    for test_size in range(max_font_size, min_font_size - 1, -1):
                        text_width = fitz.get_text_length(display_name, fontname="helv", fontsize=test_size)
                        if text_width <= available_width:
                            optimal_font_size = test_size
                            break

                    # 如果仍溢出，在普通行中依然采用截断策略（不再强制独占一行）
                    if optimal_font_size == min_font_size:
                        test_width = 0
                        max_chars = 0
                        for ch in display_name:
                            ch_width = fitz.get_text_length(ch, fontname="helv", fontsize=min_font_size)
                            if test_width + ch_width <= available_width:
                                test_width += ch_width
                                max_chars += 1
                            else:
                                break
                        if max_chars < len(display_name):
                            display_name = (display_name[:max_chars-3] + "...") if max_chars > 3 else display_name[:max_chars]

                    try:
                        page.insert_text(
                            fitz.Point(x_pos, y_pos),
                            display_name,
                            fontsize=optimal_font_size,
                            color=(0, 0, 0),
                            fontname="helv",
                        )
                    except Exception:
                        pass

                # 数量列紧跟在名称列右侧
                x_qty = x_pos + col_width_sku
                if qty:
                    qty_text = str(qty)
                    try:
                        page.insert_text(
                            fitz.Point(x_qty, y_pos),
                            qty_text,
                            fontsize=base_font_size,
                            color=(0, 0, 0),
                            fontname="helv",
                        )
                    except Exception:
                        pass


def process_pdf(input_path: str, output_path: str):
    """
    处理PDF文件：提取SKU信息并插入到物流面单
    只输出物流面单页面，不包含拣货单
    """
    print(f"正在处理PDF文件: {input_path}")
    
    # 使用PyMuPDF打开PDF
    doc = fitz.open(input_path)
    
    # 使用pdfplumber提取文本
    packing_slip_data = {}
    
    with pdfplumber.open(input_path) as pdf:
        for page_num in range(len(pdf.pages)):
            # 偶数页（索引从0开始，所以1, 3, 5...）是拣货单
            if page_num % 2 == 1:
                page = pdf.pages[page_num]
                # 优先使用基于坐标的解析，以 Seller~Qty 区域为准识别 SKU
                skus = extract_sku_from_page(page)
                if skus:
                    # 对应的物流面单是前一页（奇数页，索引从0开始，所以0, 2, 4...）
                    shipping_label_page = page_num - 1
                    packing_slip_data[shipping_label_page] = skus
                    print(f"页面 {page_num + 1} (拣货单) 提取到 {len(skus)} 个SKU: {skus}")
                    print(f"  将插入到页面 {shipping_label_page + 1} (物流面单)")
    
    # 在物流面单上添加表格
    for shipping_label_page, skus in packing_slip_data.items():
        print(f"\n在页面 {shipping_label_page + 1} 上创建表格...")
        create_sku_table(doc, shipping_label_page, skus)
    
    # 创建新的PDF文档，只包含物流面单页面（奇数页）
    output_doc = fitz.open()
    
    shipping_label_count = 0
    for page_num in range(len(doc)):
        # 只保留奇数页（索引从0开始，所以0, 2, 4...是物流面单）
        if page_num % 2 == 0:
            # 复制页面到新文档
            output_doc.insert_pdf(doc, from_page=page_num, to_page=page_num)
            shipping_label_count += 1
    
    # 保存输出文件
    output_doc.save(output_path)
    output_doc.close()
    doc.close()
    
    print(f"\n处理完成，输出文件: {output_path}")
    print(f"输出文件包含 {shipping_label_count} 个物流面单页面（已移除拣货单）")


if __name__ == "__main__":
    import sys
    
    # 如果命令行提供了文件名，使用命令行参数
    if len(sys.argv) > 1:
        input_file = sys.argv[1]
        # 自动生成输出文件名
        if input_file.endswith('.pdf'):
            output_file = input_file.replace('.pdf', '_processed.pdf')
        else:
            output_file = input_file + '_processed.pdf'
    else:
        # 默认处理第二个PDF文件
        input_file = "10-06_05-29-05_Shippinglabel+Packingslip.pdf"
        output_file = "10-06_05-29-05_Shippinglabel+Packingslip_processed.pdf"
    
    process_pdf(input_file, output_file)
