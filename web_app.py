#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF处理Web应用 - 内存优化版本
提供文件上传和处理后的文件下载功能
"""

from flask import Flask, render_template, request, send_file, flash, redirect, url_for
from werkzeug.utils import secure_filename
import os
import gc
import tempfile
import shutil
from process_pdf import process_pdf
import traceback

app = Flask(__name__)
app.config['SECRET_KEY'] = 'your-secret-key-here'
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['OUTPUT_FOLDER'] = 'outputs'
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 最大100MB

# 确保上传和输出文件夹存在
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)

ALLOWED_EXTENSIONS = {'pdf'}


def cleanup_old_files():
    """处理前清理旧文件释放磁盘空间"""
    try:
        for folder in [app.config['UPLOAD_FOLDER'], app.config['OUTPUT_FOLDER']]:
            for f in os.listdir(folder):
                try:
                    os.remove(os.path.join(folder, f))
                except:
                    pass
    except:
        pass
    gc.collect()

def allowed_file(filename):
    """检查文件扩展名是否允许"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    """主页"""
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    """处理文件上传和PDF处理"""
    # 先清理旧文件
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
        try:
            # 保存上传的文件
            filename = secure_filename(file.filename)
            input_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(input_path)
            
            # 生成输出文件名
            base_name = os.path.splitext(filename)[0]
            output_filename = f"{base_name}_processed.pdf"
            output_path = os.path.join(app.config['OUTPUT_FOLDER'], output_filename)
            
            # 处理PDF
            label_count = process_pdf(input_path, output_path)
            
            # 处理完成后删除上传的原文件释放空间
            try:
                if input_path and os.path.exists(input_path):
                    os.remove(input_path)
            except:
                pass
            
            gc.collect()
            
            # 检查输出文件是否存在
            if os.path.exists(output_path):
                flash(f'处理成功！共 {label_count} 个物流面单')
                return render_template('index.html', 
                                      download_file=output_filename,
                                      success=True)
            else:
                flash('处理失败：输出文件未生成')
                return redirect(url_for('index'))
                
        except Exception as e:
            error_msg = f'处理文件时出错: {str(e)}'
            flash(error_msg)
            print(f"错误详情: {traceback.format_exc()}")
            # 清理
            try:
                if input_path and os.path.exists(input_path):
                    os.remove(input_path)
            except:
                pass
            gc.collect()
            return redirect(url_for('index'))
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
        # 清理上传文件夹（保留最近的文件）
        for folder in [app.config['UPLOAD_FOLDER'], app.config['OUTPUT_FOLDER']]:
            files = os.listdir(folder)
            # 只保留最近10个文件
            if len(files) > 10:
                files.sort(key=lambda x: os.path.getmtime(os.path.join(folder, x)))
                for file in files[:-10]:
                    os.remove(os.path.join(folder, file))
        flash('清理完成')
    except Exception as e:
        flash(f'清理时出错: {str(e)}')
    return redirect(url_for('index'))

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    print("=" * 60)
    print("PDF处理Web应用启动中...")
    print("=" * 60)
    print(f"上传文件夹: {os.path.abspath(app.config['UPLOAD_FOLDER'])}")
    print(f"输出文件夹: {os.path.abspath(app.config['OUTPUT_FOLDER'])}")
    print("=" * 60)
    print(f"请在浏览器中访问: http://localhost:{port}")
    print("按 Ctrl+C 停止服务器")
    print("=" * 60)
    
    app.run(debug=False, host='0.0.0.0', port=port)
