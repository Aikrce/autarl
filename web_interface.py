#!/usr/bin/env python3
"""
Web Interface for Template-Based Markdown Converter
基于Web的模板上传和Markdown转换界面
"""

import os
import json
import logging
from pathlib import Path
from typing import Dict, List, Optional, Any
import tempfile
import shutil
from datetime import datetime

from flask import Flask, request, jsonify, render_template, send_file, redirect, url_for, flash
from werkzeug.utils import secure_filename
from werkzeug.exceptions import RequestEntityTooLarge

from word_template_analyzer import TemplateLibrary, analyze_word_template
from template_based_converter import AdvancedTemplateConverter, ConversionResult
from markdown_style_mapper import analyze_markdown_for_mapping
from enhanced_document_analyzer import analyze_markdown_document

# 配置日志
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Flask应用配置
app = Flask(__name__)
app.config['SECRET_KEY'] = 'your-secret-key-here'
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB最大文件大小
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['OUTPUT_FOLDER'] = 'outputs'
app.config['TEMPLATE_LIBRARY'] = 'template_library'

# 创建必要的目录
for folder in [app.config['UPLOAD_FOLDER'], app.config['OUTPUT_FOLDER'], app.config['TEMPLATE_LIBRARY']]:
    Path(folder).mkdir(exist_ok=True)

# 全局变量
converter = None
template_library = None

# 允许的文件扩展名
ALLOWED_WORD_EXTENSIONS = {'docx', 'dotx'}
ALLOWED_MARKDOWN_EXTENSIONS = {'md', 'markdown', 'txt'}


def allowed_file(filename: str, allowed_extensions: set) -> bool:
    """检查文件扩展名是否允许"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in allowed_extensions


def init_converter():
    """初始化转换器"""
    global converter, template_library
    
    try:
        template_library = TemplateLibrary(app.config['TEMPLATE_LIBRARY'])
        converter = AdvancedTemplateConverter(app.config['TEMPLATE_LIBRARY'])
        logger.info("转换器初始化成功")
    except Exception as e:
        logger.error(f"转换器初始化失败: {e}")


# 初始化转换器
init_converter()


@app.route('/')
def index():
    """主页"""
    try:
        templates = converter.get_available_templates() if converter else {}
        return render_template('index.html', templates=templates)
    except Exception as e:
        logger.error(f"加载主页失败: {e}")
        return render_template('error.html', error=str(e)), 500


@app.route('/upload_template', methods=['GET', 'POST'])
def upload_template():
    """上传Word模板"""
    if request.method == 'GET':
        return render_template('upload_template.html')
    
    try:
        # 检查文件
        if 'template_file' not in request.files:
            flash('没有选择文件', 'error')
            return redirect(request.url)
        
        file = request.files['template_file']
        if file.filename == '':
            flash('没有选择文件', 'error')
            return redirect(request.url)
        
        if not allowed_file(file.filename, ALLOWED_WORD_EXTENSIONS):
            flash('只支持 .docx 和 .dotx 文件', 'error')
            return redirect(request.url)
        
        # 获取表单数据
        template_name = request.form.get('template_name', '').strip()
        description = request.form.get('description', '').strip()
        tags = request.form.get('tags', '').strip().split(',')
        tags = [tag.strip() for tag in tags if tag.strip()]
        
        if not template_name:
            flash('请输入模板名称', 'error')
            return redirect(request.url)
        
        # 保存上传的文件
        filename = secure_filename(file.filename)
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        unique_filename = f"{timestamp}_{filename}"
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], unique_filename)
        file.save(file_path)
        
        # 分析模板
        try:
            template_info = analyze_word_template(file_path)
            
            # 添加到模板库
            template_id = template_library.add_template(
                file_path, template_name, description, tags
            )
            
            # 更新转换器
            converter.converters[template_id] = converter.template_library
            
            flash(f'模板 "{template_name}" 上传成功！', 'success')
            
            return render_template('template_analysis.html', 
                                 template_info=template_info.to_dict(),
                                 template_id=template_id,
                                 template_name=template_name)
            
        except Exception as e:
            # 删除上传的文件
            if os.path.exists(file_path):
                os.remove(file_path)
            flash(f'模板分析失败: {str(e)}', 'error')
            return redirect(request.url)
    
    except RequestEntityTooLarge:
        flash('文件太大，请选择小于50MB的文件', 'error')
        return redirect(request.url)
    except Exception as e:
        logger.error(f"上传模板失败: {e}")
        flash(f'上传失败: {str(e)}', 'error')
        return redirect(request.url)


@app.route('/templates')
def list_templates():
    """模板列表页面"""
    try:
        templates = converter.get_available_templates() if converter else {}
        return render_template('templates.html', templates=templates)
    except Exception as e:
        logger.error(f"获取模板列表失败: {e}")
        return render_template('error.html', error=str(e)), 500


@app.route('/template/<template_id>')
def template_detail(template_id):
    """模板详情页面"""
    try:
        template_info = converter.get_template_info(template_id) if converter else None
        
        if not template_info:
            flash('模板不存在', 'error')
            return redirect(url_for('list_templates'))
        
        return render_template('template_detail.html', 
                             template_info=template_info,
                             template_id=template_id)
    except Exception as e:
        logger.error(f"获取模板详情失败: {e}")
        return render_template('error.html', error=str(e)), 500


@app.route('/batch_convert', methods=['GET', 'POST'])
def batch_convert():
    """批量转换"""
    if request.method == 'GET':
        templates = converter.get_available_templates() if converter else {}
        return render_template('batch_convert.html', templates=templates)
    
    try:
        # 检查上传的文件
        if 'markdown_files' not in request.files:
            flash('没有选择文件', 'error')
            return redirect(request.url)
        
        files = request.files.getlist('markdown_files')
        
        if not files or all(f.filename == '' for f in files):
            flash('请选择要转换的文件', 'error')
            return redirect(request.url)
        
        # 验证文件类型
        valid_files = []
        for file in files:
            if file.filename and allowed_file(file.filename, ALLOWED_MARKDOWN_EXTENSIONS):
                valid_files.append(file)
        
        if not valid_files:
            flash('没有有效的Markdown文件', 'error')
            return redirect(request.url)
        
        # 获取模板选择
        template_id = request.form.get('template_id', '').strip()
        
        # 创建临时目录保存上传的文件
        temp_dir = tempfile.mkdtemp()
        uploaded_files = []
        
        try:
            # 保存上传的文件
            for file in valid_files:
                filename = secure_filename(file.filename)
                file_path = os.path.join(temp_dir, filename)
                file.save(file_path)
                uploaded_files.append(file_path)
            
            # 创建输出目录
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            output_dir = os.path.join(app.config['OUTPUT_FOLDER'], f"batch_{timestamp}")
            
            # 执行批量转换
            results = converter.batch_convert(
                uploaded_files, 
                output_dir, 
                template_id if template_id else None
            )
            
            # 统计结果
            total_files = len(results)
            successful_files = sum(1 for r in results.values() if r.success)
            failed_files = total_files - successful_files
            
            # 创建下载包
            if successful_files > 0:
                zip_path = self._create_download_zip(output_dir, f"batch_converted_{timestamp}.zip")
                zip_filename = os.path.basename(zip_path)
            else:
                zip_filename = None
            
            return render_template('batch_result.html',
                                 results=results,
                                 total_files=total_files,
                                 successful_files=successful_files,
                                 failed_files=failed_files,
                                 download_url=url_for('download_file', filename=zip_filename) if zip_filename else None)
        
        finally:
            # 清理临时文件
            shutil.rmtree(temp_dir, ignore_errors=True)
    
    except Exception as e:
        logger.error(f"批量转换失败: {e}")
        flash(f'批量转换过程中发生错误: {str(e)}', 'error')
        return redirect(request.url)


@app.route('/download/<filename>')
def download_file(filename):
    """下载文件"""
    try:
        file_path = os.path.join(app.config['OUTPUT_FOLDER'], filename)
        
        if not os.path.exists(file_path):
            flash('文件不存在', 'error')
            return redirect(url_for('index'))
        
        return send_file(file_path, as_attachment=True)
    
    except Exception as e:
        logger.error(f"下载文件失败: {e}")
        flash(f'下载失败: {str(e)}', 'error')
        return redirect(url_for('index'))


@app.route('/api/analyze_markdown', methods=['POST'])
def api_analyze_markdown():
    """API: 分析Markdown文档"""
    try:
        data = request.get_json()
        
        if not data or 'content' not in data:
            return jsonify({'error': '缺少Markdown内容'}), 400
        
        markdown_content = data['content']
        
        # 分析文档
        doc_analysis = analyze_markdown_document(markdown_content)
        doc_context = analyze_markdown_for_mapping(markdown_content)
        
        # 建议模板
        suggested_template = converter.auto_select_template(markdown_content) if converter else None
        
        return jsonify({
            'document_analysis': {
                'document_type': doc_analysis.document_type.value,
                'sections_count': len(doc_analysis.sections),
                'detected_components': list(doc_analysis.detected_components),
                'confidence_score': doc_analysis.confidence_score
            },
            'document_context': doc_context,
            'suggested_template': suggested_template,
            'available_templates': converter.get_available_templates() if converter else {}
        })
    
    except Exception as e:
        logger.error(f"分析Markdown失败: {e}")
        return jsonify({'error': str(e)}), 500


@app.route('/api/templates', methods=['GET'])
def api_list_templates():
    """API: 获取模板列表"""
    try:
        templates = converter.get_available_templates() if converter else {}
        return jsonify({'templates': templates})
    except Exception as e:
        logger.error(f"获取模板列表失败: {e}")
        return jsonify({'error': str(e)}), 500


@app.route('/api/template/<template_id>', methods=['GET'])
def api_template_detail(template_id):
    """API: 获取模板详情"""
    try:
        template_info = converter.get_template_info(template_id) if converter else None
        
        if not template_info:
            return jsonify({'error': '模板不存在'}), 404
        
        return jsonify({'template_info': template_info})
    except Exception as e:
        logger.error(f"获取模板详情失败: {e}")
        return jsonify({'error': str(e)}), 500


@app.route('/api/convert', methods=['POST'])
def api_convert():
    """API: 转换Markdown"""
    try:
        data = request.get_json()
        
        if not data or 'content' not in data:
            return jsonify({'error': '缺少Markdown内容'}), 400
        
        markdown_content = data['content']
        template_id = data.get('template_id')
        
        # 生成输出文件
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_filename = f"api_converted_{timestamp}.docx"
        output_path = os.path.join(app.config['OUTPUT_FOLDER'], output_filename)
        
        # 执行转换
        if template_id:
            result = converter.convert_with_template(markdown_content, template_id, output_path)
        else:
            result = converter.convert_with_auto_template(markdown_content, output_path)
        
        if result.success:
            download_url = url_for('download_file', filename=output_filename, _external=True)
            
            response_data = result.to_dict()
            response_data['download_url'] = download_url
            
            return jsonify(response_data)
        else:
            return jsonify({
                'error': '转换失败',
                'details': result.errors
            }), 500
    
    except Exception as e:
        logger.error(f"API转换失败: {e}")
        return jsonify({'error': str(e)}), 500


@app.route('/help')
def help_page():
    """帮助页面"""
    return render_template('help.html')


@app.route('/about')
def about():
    """关于页面"""
    return render_template('about.html')


def _create_download_zip(source_dir: str, zip_filename: str) -> str:
    """创建下载用的ZIP文件"""
    import zipfile
    
    zip_path = os.path.join(app.config['OUTPUT_FOLDER'], zip_filename)
    
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk(source_dir):
            for file in files:
                if file.endswith('.docx'):
                    file_path = os.path.join(root, file)
                    archive_name = os.path.relpath(file_path, source_dir)
                    zipf.write(file_path, archive_name)
    
    return zip_path


@app.route('/convert', methods=['POST'])
def convert_file():
    """处理文件上传转换（兼容当前HTML表单）"""
    try:
        # 检查是否有文件上传
        if 'file' not in request.files:
            flash('没有选择文件', 'error')
            return redirect('/')
        
        file = request.files['file']
        if file.filename == '':
            flash('没有选择文件', 'error')
            return redirect('/')
        
        if not allowed_file(file.filename, ALLOWED_MARKDOWN_EXTENSIONS):
            flash('只支持 .md, .markdown, .txt 文件', 'error')
            return redirect('/')
        
        # 读取文件内容
        markdown_content = file.read().decode('utf-8')
        
        # 获取模板选择
        template_id = request.form.get('template', 'default')
        output_format = request.form.get('format', 'docx')
        method = request.form.get('method', 'python-docx')
        
        logger.info(f"转换请求 - 模板: {template_id}, 格式: {output_format}, 方法: {method}")
        
        # 生成输出文件名
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        base_name = os.path.splitext(file.filename)[0]
        output_filename = f"{base_name}_{timestamp}.{output_format}"
        output_path = os.path.join(app.config['OUTPUT_FOLDER'], output_filename)
        
        # 根据选择的方法进行转换
        if method == 'pandoc':
            # 使用pandoc转换（如果可用）
            success = convert_with_pandoc(markdown_content, output_path, output_format)
        else:
            # 使用我们的增强转换器
            if template_id in ['default', 'nenu_thesis', 'business', 'technical', 'simple_report']:
                # 内置模板 - 使用简化转换器
                success = convert_with_enhanced_converter(markdown_content, output_path, template_id)
            else:
                # 自定义模板 - 使用高级转换器
                if converter:
                    result = converter.convert_with_template(markdown_content, template_id, output_path)
                    success = result.success
                    if not success:
                        flash(f'转换失败: {"; ".join(result.errors)}', 'error')
                        return redirect('/')
                else:
                    flash('转换器未初始化', 'error')
                    return redirect('/')
        
        if success:
            flash(f'转换成功！文件已保存为 {output_filename}', 'success')
            return send_file(output_path, as_attachment=True, download_name=output_filename)
        else:
            flash('转换失败', 'error')
            return redirect('/')
            
    except Exception as e:
        logger.error(f"转换失败: {e}")
        flash(f'转换过程中发生错误: {str(e)}', 'error')
        return redirect('/')


@app.route('/convert_text', methods=['POST'])
def convert_text():
    """处理文本输入转换"""
    try:
        # 获取文本内容
        markdown_content = request.form.get('markdown_text', '').strip()
        
        if not markdown_content:
            flash('请输入Markdown内容', 'error')
            return redirect('/')
        
        # 获取模板选择
        template_id = request.form.get('template', 'default')
        output_format = request.form.get('format', 'docx')
        method = request.form.get('method', 'python-docx')
        
        logger.info(f"文本转换请求 - 模板: {template_id}, 格式: {output_format}")
        
        # 生成输出文件名
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_filename = f"text_converted_{timestamp}.{output_format}"
        output_path = os.path.join(app.config['OUTPUT_FOLDER'], output_filename)
        
        # 根据选择的方法进行转换
        if method == 'pandoc':
            success = convert_with_pandoc(markdown_content, output_path, output_format)
        else:
            if template_id in ['default', 'nenu_thesis', 'business', 'technical', 'simple_report']:
                success = convert_with_enhanced_converter(markdown_content, output_path, template_id)
            else:
                if converter:
                    result = converter.convert_with_template(markdown_content, template_id, output_path)
                    success = result.success
                    if not success:
                        flash(f'转换失败: {"; ".join(result.errors)}', 'error')
                        return redirect('/')
                else:
                    flash('转换器未初始化', 'error')
                    return redirect('/')
        
        if success:
            flash(f'转换成功！文件已保存为 {output_filename}', 'success')
            return send_file(output_path, as_attachment=True, download_name=output_filename)
        else:
            flash('转换失败', 'error')
            return redirect('/')
            
    except Exception as e:
        logger.error(f"文本转换失败: {e}")
        flash(f'转换过程中发生错误: {str(e)}', 'error')
        return redirect('/')


def convert_with_pandoc(content, output_path, output_format):
    """使用pandoc进行转换"""
    try:
        import pypandoc
        
        # 创建临时输入文件
        temp_input = tempfile.NamedTemporaryFile(mode='w', suffix='.md', delete=False, encoding='utf-8')
        temp_input.write(content)
        temp_input.close()
        
        # 转换
        pypandoc.convert_file(temp_input.name, output_format, outputfile=output_path)
        
        # 清理临时文件
        os.unlink(temp_input.name)
        
        return True
    except Exception as e:
        logger.error(f"Pandoc转换失败: {e}")
        return False


def convert_with_enhanced_converter(content, output_path, template_name):
    """使用增强转换器进行转换"""
    try:
        # 导入我们的增强转换器
        from table_enhanced_converter import enhanced_markdown_to_docx
        
        # 创建临时输入文件
        temp_input = tempfile.NamedTemporaryFile(mode='w', suffix='.md', delete=False, encoding='utf-8')
        temp_input.write(content)
        temp_input.close()
        
        # 转换
        success = enhanced_markdown_to_docx(temp_input.name, output_path)
        
        # 清理临时文件
        os.unlink(temp_input.name)
        
        return success
    except Exception as e:
        logger.error(f"增强转换器失败: {e}")
        return False


# 错误处理
@app.errorhandler(404)
def not_found_error(error):
    return render_template('error.html', error='页面不存在'), 404


@app.errorhandler(500)
def internal_error(error):
    return render_template('error.html', error='服务器内部错误'), 500


@app.errorhandler(RequestEntityTooLarge)
def too_large_error(error):
    flash('文件太大，请选择小于50MB的文件', 'error')
    return redirect(request.url)


if __name__ == '__main__':
    # 创建必要的目录
    for folder in [app.config['UPLOAD_FOLDER'], app.config['OUTPUT_FOLDER']]:
        Path(folder).mkdir(exist_ok=True)
    
    # 运行应用
    app.run(debug=True, host='0.0.0.0', port=8080)