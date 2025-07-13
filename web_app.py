#!/usr/bin/env python3

import os
import tempfile
from flask import Flask, render_template, request, send_file, flash, redirect, url_for
from werkzeug.utils import secure_filename
from core.markdown_to_word import MarkdownToWordConverter
from core.complete_converter import FixedCompleteMarkdownConverter as CompleteMarkdownConverter
from templates_config import list_templates
import shutil
from datetime import datetime

app = Flask(__name__)
app.config['SECRET_KEY'] = 'your-secret-key-here'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 限制上传文件大小为16MB
app.config['UPLOAD_FOLDER'] = tempfile.mkdtemp()
app.config['SEND_FILE_MAX_AGE_DEFAULT'] = 0  # 禁用缓存

# 允许的文件扩展名
ALLOWED_EXTENSIONS = {'md', 'markdown', 'txt'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    templates = list_templates()
    print(f"DEBUG: Templates loaded: {templates}")  # 调试信息
    return render_template('index.html', templates=templates)

@app.route('/convert', methods=['POST'])
def convert():
    # 检查是否有文件上传
    if 'file' not in request.files:
        flash('没有选择文件', 'error')
        return redirect(url_for('index'))
    
    file = request.files['file']
    
    # 检查文件名是否为空
    if file.filename == '':
        flash('没有选择文件', 'error')
        return redirect(url_for('index'))
    
    if file and allowed_file(file.filename):
        # 保存上传的文件
        filename = secure_filename(file.filename)
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        input_filename = f"{timestamp}_{filename}"
        input_path = os.path.join(app.config['UPLOAD_FOLDER'], input_filename)
        file.save(input_path)
        
        # 获取输出格式
        output_format = request.form.get('format', 'docx')
        format_extensions = {
            'docx': '.docx',
            'pdf': '.pdf',
            'html': '.html',
            'txt': '.txt'
        }
        
        # 创建输出文件路径
        output_filename = input_filename.rsplit('.', 1)[0] + format_extensions.get(output_format, '.docx')
        output_path = os.path.join(app.config['UPLOAD_FOLDER'], output_filename)
        
        # 转换文件
        template_name = request.form.get('template', 'default')
        enable_mermaid = request.form.get('enable_mermaid', 'off') == 'on'
        
        # 获取转换方法
        method = request.form.get('method', 'python-docx')
        
        try:
            success = False
            
            # 如果启用了Mermaid支持，使用完整转换器
            if enable_mermaid:
                converter = CompleteMarkdownConverter(mermaid_method='web')
                if output_format == 'docx':
                    success = converter.convert(input_path, output_path)
                else:
                    flash('Mermaid支持目前仅适用于DOCX格式', 'warning')
                    converter = MarkdownToWordConverter(template_name=template_name)
            else:
                converter = MarkdownToWordConverter(template_name=template_name)
            
            if not enable_mermaid:
                if method == 'pandoc':
                    success = converter.convert_with_pandoc(input_path, output_path, output_format)
                else:
                    # 使用原生Python方法
                    if output_format == 'docx':
                        success = converter.convert_with_python_docx(input_path, output_path)
                    elif output_format == 'html':
                        success = converter.convert_to_html(input_path, output_path)
                    elif output_format == 'txt':
                        success = converter.convert_to_txt(input_path, output_path)
                    else:
                        flash(f'格式 {output_format.upper()} 仅支持Pandoc转换方法，请选择Pandoc或切换到其他格式', 'warning')
                        return redirect(url_for('index'))
            
            if success:
                # 清理输入文件
                os.remove(input_path)
                
                # 设置MIME类型
                mime_types = {
                    'docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                    'pdf': 'application/pdf',
                    'html': 'text/html',
                    'txt': 'text/plain'
                }
                
                # 返回转换后的文件
                return send_file(
                    output_path,
                    as_attachment=True,
                    download_name=output_filename,
                    mimetype=mime_types.get(output_format, 'application/octet-stream')
                )
            else:
                flash('转换失败，请检查文件格式或尝试其他转换方法', 'error')
                return redirect(url_for('index'))
                
        except Exception as e:
            flash(f'转换出错: {str(e)}', 'error')
            return redirect(url_for('index'))
    
    else:
        flash('不支持的文件格式，请上传 .md、.markdown 或 .txt 文件', 'error')
        return redirect(url_for('index'))

@app.route('/convert_text', methods=['POST'])
def convert_text():
    # 获取文本内容
    markdown_text = request.form.get('markdown_text', '')
    
    if not markdown_text.strip():
        flash('请输入Markdown内容', 'warning')
        return redirect(url_for('index'))
    
    # 获取输出格式
    output_format = request.form.get('format', 'docx')
    format_extensions = {
        'docx': '.docx',
        'pdf': '.pdf',
        'html': '.html',
        'txt': '.txt'
    }
    
    # 创建临时文件
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    input_filename = f"{timestamp}_text_input.md"
    input_path = os.path.join(app.config['UPLOAD_FOLDER'], input_filename)
    
    # 保存文本到文件
    with open(input_path, 'w', encoding='utf-8') as f:
        f.write(markdown_text)
    
    # 创建输出文件路径
    output_filename = f"{timestamp}_converted{format_extensions.get(output_format, '.docx')}"
    output_path = os.path.join(app.config['UPLOAD_FOLDER'], output_filename)
    
    # 转换文件
    template_name = request.form.get('template', 'default')
    enable_mermaid = request.form.get('enable_mermaid', 'off') == 'on'
    
    # 获取转换方法
    method = request.form.get('method', 'python-docx')
    
    try:
        success = False
        
        # 如果启用了Mermaid支持，使用完整转换器
        if enable_mermaid:
            converter = CompleteMarkdownConverter(mermaid_method='web')
            if output_format == 'docx':
                success = converter.convert(input_path, output_path)
            else:
                flash('Mermaid支持目前仅适用于DOCX格式', 'warning')
                converter = MarkdownToWordConverter(template_name=template_name)
        else:
            converter = MarkdownToWordConverter(template_name=template_name)
        
        if not enable_mermaid:
            if method == 'pandoc':
                success = converter.convert_with_pandoc(input_path, output_path, output_format)
            else:
                # 使用原生Python方法
                if output_format == 'docx':
                    success = converter.convert_with_python_docx(input_path, output_path)
                elif output_format == 'html':
                    success = converter.convert_to_html(input_path, output_path)
                elif output_format == 'txt':
                    success = converter.convert_to_txt(input_path, output_path)
                else:
                    flash(f'格式 {output_format.upper()} 仅支持Pandoc转换方法，请选择Pandoc或切换到其他格式', 'warning')
                    return redirect(url_for('index'))
        
        if success:
            # 清理输入文件
            os.remove(input_path)
            
            # 设置MIME类型
            mime_types = {
                'docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                'pdf': 'application/pdf',
                'html': 'text/html',
                'txt': 'text/plain'
            }
            
            # 返回转换后的文件
            return send_file(
                output_path,
                as_attachment=True,
                download_name=output_filename,
                mimetype=mime_types.get(output_format, 'application/octet-stream')
            )
        else:
            flash('转换失败，请尝试其他转换方法', 'error')
            return redirect(url_for('index'))
            
    except Exception as e:
        flash(f'转换出错: {str(e)}', 'error')
        return redirect(url_for('index'))

@app.teardown_appcontext
def cleanup(exc):
    # 清理临时文件
    temp_dir = app.config['UPLOAD_FOLDER']
    if os.path.exists(temp_dir):
        # 删除超过1小时的文件
        current_time = datetime.now()
        for filename in os.listdir(temp_dir):
            file_path = os.path.join(temp_dir, filename)
            if os.path.isfile(file_path):
                file_time = datetime.fromtimestamp(os.path.getmtime(file_path))
                if (current_time - file_time).seconds > 3600:
                    try:
                        os.remove(file_path)
                    except:
                        pass

if __name__ == '__main__':
    # 确保上传目录存在
    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
    
    # 运行应用
    app.run(debug=True, host='127.0.0.1', port=8080)