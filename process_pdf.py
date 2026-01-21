#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
处理PDF文件：从拣货单提取SKU信息并插入到物流面单
内存优化版本 - 适用于低内存服务器(512MB)
"""

import pdfplumber
import fitz  # PyMuPDF
import re
import gc
import os
import sys
import tempfile

from typing import List, Tuple, Dict


def get_memory_mb():
    """获取当前进程内存使用量(MB)"""
    try:
        # Linux系统：读取/proc/self/status
        with open('/proc/self/status', 'r') as f:
            for line in f:
                if line.startswith('VmRSS:'):
                    # VmRSS是实际使用的物理内存
                    return int(line.split()[1]) / 1024  # KB转MB
    except:
        pass
    
    try:
        # 备用方案：使用resource模块
        import resource
        usage = resource.getrusage(resource.RUSAGE_SELF)
        # macOS返回字节，Linux返回KB
        if sys.platform == 'darwin':
            return usage.ru_maxrss / 1024 / 1024  # 字节转MB
        else:
            return usage.ru_maxrss / 1024  # KB转MB
    except:
        pass
    
    return 0


def log_memory(msg: str):
    """打印带内存信息的日志"""
    mem = get_memory_mb()
    if mem > 0:
        print(f"[{mem:.1f}MB] {msg}", flush=True)
    else:
        print(msg, flush=True)

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

def extract_sku_from_packing_slip(text: str) -> List[Tuple[str, int]]:
    """
    从拣货单文本中提取SKU和数量
    返回: [(SKU名称, 数量), ...]
    """
    skus = []
    lines = text.split('\n')
    
    found_header = False
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
            continue
        
        # 在找到表头后，查找SKU行
        if found_header:
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
                    sku_parts = []
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
    
    # 绘制表格边框
    rect = fitz.Rect(table_x, y_pos, table_x + table_width, y_pos + table_height)
    
    # 绘制外边框
    page.draw_rect(rect, color=(0, 0, 0), width=1)
    
    # 绘制列分隔线（3条垂直线，将表格分成4列）
    for i in range(1, 4):  # 3条垂直线
        if i == 2:
            # 中间分隔线（将两个SKU分开）
            x = table_x + col_width_sku + col_width_qty
        elif i == 1:
            # 第一个SKU的数量列分隔线
            x = table_x + col_width_sku
        else:  # i == 3
            # 第二个SKU的数量列分隔线
            x = table_x + col_width_sku + col_width_qty + col_width_sku
        
        page.draw_line(
            fitz.Point(x, y_pos),
            fitz.Point(x, y_pos + table_height),
            color=(0, 0, 0),
            width=0.5
        )
    
    # 绘制行分隔线（将表格分成多行）
    for i in range(1, TABLE_ROWS):  # TABLE_ROWS-1条水平线
        y = y_pos + i * row_height
        page.draw_line(
            fitz.Point(table_x, y),
            fitz.Point(table_x + table_width, y),
            color=(0, 0, 0),
            width=0.5
        )
    
    # 填充数据
    # 计算合适的字号
    font_size = 9  # 初始字号
    max_font_size = 11
    
    # 测试字号是否合适
    test_text = "A" * 30
    for test_size in range(max_font_size, 6, -1):
        text_width = fitz.get_text_length(test_text, fontname="helv", fontsize=test_size)
        if text_width <= col_width_sku - 6:
            font_size = test_size
            break
    
    # 填充数据
    # 计算文本垂直居中位置（每行高度相等）
    text_y_offset = row_height / 2 + font_size / 3  # 垂直居中
    
    for row_idx in range(TABLE_ROWS):
        y_pos = table_y + row_idx * row_height + text_y_offset  # 每行使用相同的计算方式
        x_pos = table_x + 3  # 左边距
        
        for col_idx in range(2):  # 每行2个SKU
            sku_idx = row_idx * 2 + col_idx
            
            if sku_idx < len(display_skus):
                sku_name, qty = display_skus[sku_idx]
                
                # 绘制SKU名称
                if sku_name:
                    # 删除"Ink-"前缀
                    display_name = sku_name
                    if display_name.startswith("Ink-"):
                        display_name = display_name[4:]  # 删除"Ink-"
                    
                    # 自动调整字号以适应单元格宽度
                    # 可用宽度 = 列宽 - 左右边距
                    available_width = col_width_sku - 6
                    min_font_size = 6  # 最小字号
                    optimal_font_size = font_size
                    
                    # 从最大字号开始，逐步减小直到文本能放入单元格
                    for test_size in range(font_size, min_font_size - 1, -1):
                        text_width = fitz.get_text_length(display_name, fontname="helv", fontsize=test_size)
                        if text_width <= available_width:
                            optimal_font_size = test_size
                            break
                    
                    # 如果即使是最小字号也放不下，则截断文本
                    if optimal_font_size == min_font_size:
                        # 使用最小字号，计算能放下的最大字符数
                        test_width = 0
                        max_chars = 0
                        for i in range(len(display_name)):
                            char_width = fitz.get_text_length(display_name[i], fontname="helv", fontsize=min_font_size)
                            if test_width + char_width <= available_width:
                                test_width += char_width
                                max_chars = i + 1
                            else:
                                break
                        
                        if max_chars < len(display_name):
                            # 截断并添加省略号
                            if max_chars > 3:
                                display_name = display_name[:max_chars-3] + "..."
                            else:
                                display_name = display_name[:max_chars]
                    
                    # 插入文本
                    try:
                        page.insert_text(
                            fitz.Point(x_pos, y_pos),
                            display_name,
                            fontsize=optimal_font_size,
                            color=(0, 0, 0),
                            fontname="helv"
                        )
                    except:
                        pass
                
                # 绘制数量
                x_pos += col_width_sku
                if qty:
                    qty_text = str(qty)
                    try:
                        page.insert_text(
                            fitz.Point(x_pos, y_pos),
                            qty_text,
                            fontsize=font_size,
                            color=(0, 0, 0),
                            fontname="helv"
                        )
                    except:
                        pass
                
                x_pos += col_width_qty
            else:
                # 空单元格
                x_pos += col_width_sku + col_width_qty


def process_pdf(input_path: str, output_path: str):
    """
    处理PDF文件：提取SKU信息并插入到物流面单
    只输出物流面单页面，不包含拣货单
    
    内存优化版本：
    - 分步骤处理，避免同时打开多个大对象
    - 使用临时文件减少内存占用
    - 定期强制垃圾回收
    - 实时内存监控
    """
    log_memory(f"开始处理: {input_path}")
    
    # 第一步：获取总页数
    with fitz.open(input_path) as doc:
        total_pages = len(doc)
    log_memory(f"PDF共 {total_pages} 页, {total_pages // 2} 个面单")
    
    # 第二步：分批提取SKU数据（避免pdfplumber一次加载太多页面）
    packing_slip_data = {}
    log_memory("开始提取SKU...")
    
    # 每次只打开PDF提取一部分页面的SKU，然后关闭释放内存
    EXTRACT_BATCH = 20  # 每次提取20页的SKU
    for batch_start in range(0, total_pages, EXTRACT_BATCH):
        batch_end = min(batch_start + EXTRACT_BATCH, total_pages)
        
        with pdfplumber.open(input_path) as pdf:
            for page_num in range(batch_start, batch_end):
                if page_num % 2 == 1:  # 拣货单页面
                    page = pdf.pages[page_num]
                    text = page.extract_text()
                    if text:
                        skus = extract_sku_from_packing_slip(text)
                        if skus:
                            shipping_label_page = page_num - 1
                            packing_slip_data[shipping_label_page] = skus
        
        # 每批提取后强制GC
        gc.collect()
    
    log_memory(f"SKU提取完成, 共 {len(packing_slip_data)} 个面单有SKU")
    
    # 第三步：处理并输出（使用临时文件分批处理）
    # 每批处理的面单数量 - 减小批次大小以降低内存峰值
    BATCH_SIZE = 5  # 每批5个面单（极低内存模式）
    
    temp_files = []
    shipping_label_count = 0
    
    try:
        # 分批处理
        label_pages = [p for p in range(total_pages) if p % 2 == 0]  # 所有面单页码
        num_batches = (len(label_pages) + BATCH_SIZE - 1) // BATCH_SIZE
        
        log_memory(f"分 {num_batches} 批处理, 每批 {BATCH_SIZE} 个面单")
        
        for batch_idx in range(num_batches):
            start_idx = batch_idx * BATCH_SIZE
            end_idx = min(start_idx + BATCH_SIZE, len(label_pages))
            batch_pages = label_pages[start_idx:end_idx]
            
            # 打开源文档
            doc = fitz.open(input_path)
            
            # 在这批页面上添加表格
            for label_page in batch_pages:
                if label_page in packing_slip_data:
                    create_sku_table(doc, label_page, packing_slip_data[label_page])
            
            # 创建临时输出文档，只包含这批面单页面
            batch_doc = fitz.open()
            for label_page in batch_pages:
                batch_doc.insert_pdf(doc, from_page=label_page, to_page=label_page)
                shipping_label_count += 1
            
            # 保存到临时文件
            temp_file = tempfile.NamedTemporaryFile(suffix='.pdf', delete=False)
            batch_doc.save(temp_file.name, garbage=4, deflate=True)
            temp_files.append(temp_file.name)
            
            # 关闭并释放内存
            batch_doc.close()
            doc.close()
            gc.collect()
            
            log_memory(f"批次 {batch_idx + 1}/{num_batches} 完成 ({len(batch_pages)} 个)")
        
        # 合并所有临时文件 - 逐个合并并立即删除以节省内存和磁盘
        log_memory("开始合并输出文件...")
        output_doc = fitz.open()
        for i, temp_file in enumerate(temp_files):
            temp_pdf = fitz.open(temp_file)
            output_doc.insert_pdf(temp_pdf)
            temp_pdf.close()
            # 合并后立即删除临时文件释放磁盘空间
            try:
                os.unlink(temp_file)
            except:
                pass
            # 每合并2个就gc一次
            if (i + 1) % 2 == 0:
                gc.collect()
                log_memory(f"已合并 {i + 1}/{len(temp_files)} 批")
        
        # 清空临时文件列表（已删除）
        temp_files = []
        
        log_memory("保存最终文件...")
        output_doc.save(output_path, garbage=4, deflate=True)
        output_doc.close()
        
    finally:
        # 清理临时文件
        for temp_file in temp_files:
            try:
                os.unlink(temp_file)
            except:
                pass
        gc.collect()
    
    log_memory(f"处理完成! 输出 {shipping_label_count} 个面单")
    
    return shipping_label_count


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
