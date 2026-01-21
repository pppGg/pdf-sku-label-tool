#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF处理Web应用 - 低内存优化版本 v0.9.1
针对512MB RAM / 0.1 CPU环境优化
提供文件上传和处理后的文件下载功能
"""

from flask import Flask, render_template, request, send_file, flash, redirect, url_for, jsonify
from werkzeug.utils import secure_filename
import os
import gc
import json
import tempfile
import shutil
from process_pdf import process_pdf, process_pdf_streaming, BATCH_SIZE
import traceback

app = Flask(__name__)
app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY', 'your-secret-key-here')
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['OUTPUT_FOLDER'] = 'outputs'

# 不限制文件大小（移除限制）
app.config['MAX_CONTENT_LENGTH'] = None

# 是否使用流式处理模式（极低内存）
USE_STREAMING = os.environ.get('USE_STREAMING', 'false').lower() == 'true'

# 历史记录文件
HISTORY_FILE = 'processing_history.json'

# 确保上传和输出文件夹存在
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)

ALLOWED_EXTENSIONS = {'pdf'}


def get_memory_usage():
    """获取当前内存使用情况（如果可用）"""
    try:
        import resource
        usage = resource.getrusage(resource.RUSAGE_SELF)
        return f"{usage.ru_maxrss / 1024:.1f}MB"
    except:
        return "N/A"


def load_history():
    """加载处理历史记录"""
    try:
        if os.path.exists(HISTORY_FILE):
            with open(HISTORY_FILE, 'r') as f:
                return json.load(f)
    except:
        pass
    return {'max_file_size_mb': 0, 'max_pages': 0}


def save_history(history):
    """保存处理历史记录"""
    try:
        with open(HISTORY_FILE, 'w') as f:
            json.dump(history, f)
    except:
        pass


def update_history(file_size_mb, page_count):
    """更新历史记录（如果有新的最大值）"""
    history = load_history()
    updated = False
    
    if file_size_mb > history.get('max_file_size_mb', 0):
        history['max_file_size_mb'] = round(file_size_mb, 2)
        updated = True
    
    if page_count > history.get('max_pages', 0):
        history['max_pages'] = page_count
        updated = True
    
    if updated:
        save_history(history)
    
    return history


def cleanup_old_files():
    """清理旧文件以释放磁盘空间"""
    try:
        for folder in [app.config['UPLOAD_FOLDER'], app.config['OUTPUT_FOLDER']]:
            files = []
            for f in os.listdir(folder):
                path = os.path.join(folder, f)
                if os.path.isfile(path):
                    files.append((path, os.path.getmtime(path)))
            
            # 按时间排序，只保留最近5个文件
            files.sort(key=lambda x: x[1], reverse=True)
            for path, _ in files[5:]:
                try:
                    os.remove(path)
                except:
                    pass
    except:
        pass
    gc.collect()

def allowed_file(filename):
    """检查文件扩展名是否允许"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def get_pdf_page_count(filepath):
    """获取PDF页数"""
    try:
        import fitz
        with fitz.open(filepath) as doc:
            return len(doc)
    except:
        return 0


@app.route('/')
def index():
    """主页"""
    # 加载历史记录
    history = load_history()
    return render_template('index.html', history=history)


@app.route('/upload', methods=['POST'])
def upload_file():
    """处理文件上传和PDF处理"""
    # 先清理旧文件释放空间
    cleanup_old_files()
    
    if 'file' not in request.files:
        flash('没有选择文件')
        return redirect(request.url)
    
    file = request.files['file']
    
    if file.filename == '':
        flash('没有选择文件')
        return redirect(request.url)
    
    if file and allowed_file(file.filename):
        input_path = None
        output_path = None
        
        try:
            # 保存上传的文件
            filename = secure_filename(file.filename)
            input_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(input_path)
            
            # 检查文件大小
            file_size_mb = os.path.getsize(input_path) / (1024 * 1024)
            print(f"上传文件大小: {file_size_mb:.1f}MB")
            
            # 检查页数
            page_count = get_pdf_page_count(input_path)
            print(f"PDF页数: {page_count}")
            
            # 生成输出文件名
            base_name = os.path.splitext(filename)[0]
            output_filename = f"{base_name}_processed.pdf"
            output_path = os.path.join(app.config['OUTPUT_FOLDER'], output_filename)
            
            # 删除可能存在的旧输出文件
            if os.path.exists(output_path):
                os.remove(output_path)
            
            print(f"开始处理... 内存使用: {get_memory_usage()}")
            
            # 根据配置选择处理模式
            if USE_STREAMING:
                # 流式模式：极低内存，但较慢
                print("使用流式处理模式")
                label_count = process_pdf_streaming(input_path, output_path)
            else:
                # 批处理模式：平衡性能和内存
                print("使用批处理模式")
                label_count = process_pdf(input_path, output_path)
            
            print(f"处理完成，内存使用: {get_memory_usage()}")
            
            # 更新历史记录
            history = update_history(file_size_mb, page_count)
            
            # 清理上传的原文件以释放空间
            if input_path and os.path.exists(input_path):
                try:
                    os.remove(input_path)
                except:
                    pass
            
            gc.collect()
            
            # 检查输出文件是否存在
            if os.path.exists(output_path):
                flash(f'处理成功！共处理 {label_count} 个物流面单')
                return render_template('index.html', 
                                      download_file=output_filename,
                                      success=True,
                                      label_count=label_count,
                                      history=history)
            else:
                flash('处理失败：输出文件未生成')
                return redirect(url_for('index'))
                
        except MemoryError:
            error_msg = '内存不足！请尝试上传更小的PDF文件'
            flash(error_msg)
            print(f"内存错误: {traceback.format_exc()}")
            gc.collect()
            history = load_history()
            return render_template('index.html', history=history)
            
        except Exception as e:
            error_msg = f'处理文件时出错: {str(e)}'
            flash(error_msg)
            print(f"错误详情: {traceback.format_exc()}")
            gc.collect()
            history = load_history()
            return render_template('index.html', history=history)
            
        finally:
            # 确保清理临时文件
            gc.collect()
    else:
        flash('不允许的文件类型，请上传PDF文件')
        return redirect(url_for('index'))

@app.route('/download/<filename>')
def download_file(filename):
    """下载处理后的文件"""
    try:
        file_path = os.path.join(app.config['OUTPUT_FOLDER'], secure_filename(filename))
        if os.path.exists(file_path):
            return send_file(file_path, 
                           as_attachment=True,
                           download_name=filename,
                           mimetype='application/pdf')
        else:
            flash('文件不存在')
            return redirect(url_for('index'))
    except Exception as e:
        flash(f'下载文件时出错: {str(e)}')
        return redirect(url_for('index'))

@app.route('/cleanup', methods=['POST'])
def cleanup():
    """清理临时文件"""
    try:
        count = 0
        for folder in [app.config['UPLOAD_FOLDER'], app.config['OUTPUT_FOLDER']]:
            for f in os.listdir(folder):
                path = os.path.join(folder, f)
                if os.path.isfile(path):
                    os.remove(path)
                    count += 1
        gc.collect()
        flash(f'清理完成，删除了 {count} 个文件')
    except Exception as e:
        flash(f'清理时出错: {str(e)}')
    return redirect(url_for('index'))


@app.route('/health')
def health():
    """健康检查端点"""
    history = load_history()
    return jsonify({
        'status': 'ok',
        'memory': get_memory_usage(),
        'history': history,
        'config': {
            'batch_size': BATCH_SIZE,
            'streaming': USE_STREAMING
        }
    })


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    history = load_history()
    print("=" * 60)
    print("PDF处理Web应用 v0.9.1 (低内存优化版)")
    print("=" * 60)
    print(f"配置信息:")
    print(f"  - 批处理大小: {BATCH_SIZE}对/批")
    print(f"  - 流式模式: {'启用' if USE_STREAMING else '禁用'}")
    print(f"  - 上传文件夹: {os.path.abspath(app.config['UPLOAD_FOLDER'])}")
    print(f"  - 输出文件夹: {os.path.abspath(app.config['OUTPUT_FOLDER'])}")
    print(f"历史记录:")
    print(f"  - 最大文件: {history.get('max_file_size_mb', 0)}MB")
    print(f"  - 最大页数: {history.get('max_pages', 0)}页")
    print("=" * 60)
    print(f"请在浏览器中访问: http://localhost:{port}")
    print("按 Ctrl+C 停止服务器")
    print("=" * 60)
    
    app.run(debug=False, host='0.0.0.0', port=port)
