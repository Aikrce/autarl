#!/usr/bin/env python3
"""
Enhanced Output Formats Module
支持多种输出格式的增强模块，包括PDF、HTML等
"""

import os
import logging
from typing import Dict, List, Optional, Any, Union
from pathlib import Path
from dataclasses import dataclass
from enum import Enum
import markdown2
import weasyprint
from weasyprint import HTML, CSS
from jinja2 import Template
import json
import base64

from enhanced_templates_config import TemplateConfig
from enhanced_document_analyzer import DocumentStructure, ContentSection, SectionType

logger = logging.getLogger(__name__)


class OutputFormat(Enum):
    """输出格式枚举"""
    DOCX = "docx"
    PDF = "pdf"
    HTML = "html"
    EPUB = "epub"
    LATEX = "latex"


@dataclass
class OutputConfig:
    """输出配置"""
    format_type: OutputFormat
    quality: str = "high"  # low, medium, high
    include_toc: bool = True
    include_page_numbers: bool = True
    include_headers: bool = True
    include_footers: bool = True
    custom_css: Optional[str] = None
    custom_options: Dict[str, Any] = None
    
    def __post_init__(self):
        if self.custom_options is None:
            self.custom_options = {}


class EnhancedOutputGenerator:
    """增强的输出生成器"""
    
    def __init__(self, template_config: TemplateConfig):
        self.template_config = template_config
        self.markdown_extensions = [
            'tables', 'fenced-code-blocks', 'header-ids', 'toc',
            'footnotes', 'smartypants', 'code-friendly'
        ]
        
        # HTML模板
        self.html_template = self._load_html_template()
        
        # CSS样式
        self.base_css = self._generate_base_css()
        
        logger.info(f"输出生成器初始化完成，模板: {template_config.name}")
    
    def _load_html_template(self) -> Template:
        """加载HTML模板"""
        template_str = """
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{{ title }}</title>
    <style>
        {{ css_content }}
    </style>
    {% if custom_css %}
    <style>
        {{ custom_css }}
    </style>
    {% endif %}
</head>
<body class="{{ body_class }}">
    {% if include_header %}
    <header class="document-header">
        <h1>{{ header_title }}</h1>
        <div class="header-meta">{{ header_meta }}</div>
    </header>
    {% endif %}
    
    {% if include_toc %}
    <nav class="table-of-contents">
        <h2>目录</h2>
        {{ toc_content }}
    </nav>
    {% endif %}
    
    <main class="document-content">
        {{ content }}
    </main>
    
    {% if include_footer %}
    <footer class="document-footer">
        <div class="footer-content">{{ footer_content }}</div>
        {% if include_page_numbers %}
        <div class="page-number">第 <span class="page-current"></span> 页</div>
        {% endif %}
    </footer>
    {% endif %}
    
    <script>
        {{ javascript }}
    </script>
</body>
</html>
        """
        return Template(template_str)
    
    def _generate_base_css(self) -> str:
        """生成基础CSS样式"""
        css = """
        /* 基础样式重置 */
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        /* 页面布局 */
        @page {
            size: A4;
            margin: 2.5cm;
        }
        
        body {
            font-family: 'Source Han Sans CN', 'Microsoft YaHei', '微软雅黑', sans-serif;
            line-height: 1.6;
            color: #333;
            background: #fff;
        }
        
        /* 标题样式 */
        h1, h2, h3, h4, h5, h6 {
            margin: 1.5em 0 0.5em 0;
            color: #2c3e50;
            page-break-after: avoid;
        }
        
        h1 {
            font-size: 2.2em;
            text-align: center;
            border-bottom: 3px solid #3498db;
            padding-bottom: 0.3em;
            margin-bottom: 1em;
        }
        
        h2 {
            font-size: 1.8em;
            border-left: 4px solid #3498db;
            padding-left: 0.5em;
        }
        
        h3 {
            font-size: 1.4em;
            color: #34495e;
        }
        
        h4 {
            font-size: 1.2em;
            color: #7f8c8d;
        }
        
        /* 段落样式 */
        p {
            margin: 0.8em 0;
            text-align: justify;
            text-indent: 2em;
        }
        
        /* 列表样式 */
        ul, ol {
            margin: 1em 0;
            padding-left: 2em;
        }
        
        li {
            margin: 0.3em 0;
        }
        
        /* 代码样式 */
        code {
            background: #f8f9fa;
            padding: 0.2em 0.4em;
            border-radius: 3px;
            font-family: 'JetBrains Mono', 'Consolas', monospace;
            font-size: 0.9em;
            color: #e74c3c;
        }
        
        pre {
            background: #2c3e50;
            color: #ecf0f1;
            padding: 1em;
            border-radius: 5px;
            overflow-x: auto;
            margin: 1em 0;
            page-break-inside: avoid;
        }
        
        pre code {
            background: none;
            color: inherit;
            padding: 0;
        }
        
        /* 表格样式 */
        table {
            width: 100%;
            border-collapse: collapse;
            margin: 1em 0;
            page-break-inside: avoid;
        }
        
        th, td {
            border: 1px solid #bdc3c7;
            padding: 0.8em;
            text-align: left;
        }
        
        th {
            background: #ecf0f1;
            font-weight: bold;
        }
        
        /* 引用样式 */
        blockquote {
            border-left: 4px solid #3498db;
            margin: 1em 0;
            padding: 0.5em 1em;
            background: #f8f9fa;
            color: #555;
            font-style: italic;
        }
        
        /* 链接样式 */
        a {
            color: #3498db;
            text-decoration: none;
        }
        
        a:hover {
            text-decoration: underline;
        }
        
        /* 图片样式 */
        img {
            max-width: 100%;
            height: auto;
            display: block;
            margin: 1em auto;
            page-break-inside: avoid;
        }
        
        /* 页眉页脚 */
        .document-header {
            text-align: center;
            border-bottom: 2px solid #3498db;
            padding-bottom: 1em;
            margin-bottom: 2em;
        }
        
        .document-footer {
            margin-top: 2em;
            padding-top: 1em;
            border-top: 1px solid #bdc3c7;
            text-align: center;
            color: #7f8c8d;
        }
        
        /* 目录样式 */
        .table-of-contents {
            margin: 2em 0;
            padding: 1em;
            background: #f8f9fa;
            border-radius: 5px;
            page-break-after: always;
        }
        
        .table-of-contents h2 {
            text-align: center;
            margin-bottom: 1em;
            border: none;
            padding: 0;
        }
        
        .toc-level-1 {
            font-weight: bold;
            margin: 0.5em 0;
        }
        
        .toc-level-2 {
            margin-left: 1em;
            margin: 0.3em 0;
        }
        
        .toc-level-3 {
            margin-left: 2em;
            margin: 0.2em 0;
            font-size: 0.9em;
        }
        
        /* 学术论文特定样式 */
        .academic-abstract {
            margin: 2em 0;
            padding: 1.5em;
            background: #f8f9fa;
            border-radius: 5px;
        }
        
        .academic-keywords {
            margin: 1em 0;
            font-weight: bold;
        }
        
        .academic-references {
            page-break-before: always;
        }
        
        .academic-references .reference-item {
            margin: 0.5em 0;
            padding-left: 2em;
            text-indent: -2em;
        }
        
        /* 打印样式 */
        @media print {
            body {
                font-size: 12pt;
            }
            
            h1 {
                page-break-before: always;
            }
            
            .page-break-before {
                page-break-before: always;
            }
            
            .page-break-after {
                page-break-after: always;
            }
            
            .no-print {
                display: none;
            }
        }
        
        /* 响应式设计 */
        @media screen and (max-width: 768px) {
            body {
                font-size: 14px;
                margin: 1em;
            }
            
            .table-of-contents {
                margin: 1em 0;
                padding: 0.5em;
            }
        }
        """
        
        # 根据模板类型添加特定样式
        if self.template_config.name == 'nenu_thesis':
            css += self._get_nenu_thesis_css()
        elif self.template_config.name == 'business_report':
            css += self._get_business_report_css()
        elif self.template_config.name == 'technical_doc':
            css += self._get_technical_doc_css()
        
        return css
    
    def _get_nenu_thesis_css(self) -> str:
        """获取东北师大论文样式"""
        return """
        /* 东北师大论文特定样式 */
        body {
            font-family: '宋体', SimSun, serif;
        }
        
        h1 {
            font-family: '黑体', SimHei, sans-serif;
            font-size: 16pt;
            text-align: center;
        }
        
        h2 {
            font-family: '黑体', SimHei, sans-serif;
            font-size: 14pt;
            border-left: none;
            padding-left: 0;
        }
        
        h3 {
            font-family: '宋体', SimSun, serif;
            font-weight: bold;
            font-size: 12pt;
        }
        
        p {
            font-size: 12pt;
            line-height: 1.5;
            text-indent: 2em;
        }
        
        .thesis-cover {
            text-align: center;
            page-break-after: always;
        }
        
        .thesis-title {
            font-size: 18pt;
            font-weight: bold;
            margin: 2em 0;
        }
        
        .thesis-info {
            margin: 1em 0;
            font-size: 12pt;
        }
        
        .abstract-title {
            font-family: '黑体', SimHei, sans-serif;
            font-size: 16pt;
            text-align: center;
            margin: 2em 0 1em 0;
        }
        
        .keywords {
            margin: 1em 0;
            font-weight: bold;
        }
        """
    
    def _get_business_report_css(self) -> str:
        """获取商务报告样式"""
        return """
        /* 商务报告特定样式 */
        body {
            font-family: '微软雅黑', 'Microsoft YaHei', sans-serif;
        }
        
        h1 {
            color: #003366;
            font-size: 20pt;
        }
        
        h2 {
            color: #336699;
            font-size: 16pt;
        }
        
        .business-summary {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 2em;
            border-radius: 10px;
            margin: 2em 0;
        }
        
        .chart-container {
            text-align: center;
            margin: 2em 0;
            page-break-inside: avoid;
        }
        """
    
    def _get_technical_doc_css(self) -> str:
        """获取技术文档样式"""
        return """
        /* 技术文档特定样式 */
        body {
            font-family: 'Source Han Sans CN', sans-serif;
        }
        
        pre {
            background: #1e1e1e;
            color: #d4d4d4;
        }
        
        .api-endpoint {
            background: #f1c40f;
            color: #2c3e50;
            padding: 0.5em 1em;
            border-radius: 5px;
            margin: 1em 0;
            font-family: monospace;
        }
        
        .code-example {
            border: 1px solid #e1e5e9;
            border-radius: 6px;
            margin: 1em 0;
        }
        
        .code-example-header {
            background: #f6f8fa;
            padding: 0.5em 1em;
            border-bottom: 1px solid #e1e5e9;
            font-weight: bold;
        }
        """
    
    def generate_html(self, content: str, document_structure: DocumentStructure,
                     output_config: OutputConfig, output_path: str) -> bool:
        """生成HTML输出"""
        try:
            # 转换Markdown为HTML
            html_content = markdown2.markdown(
                content,
                extras=self.markdown_extensions
            )
            
            # 生成目录
            toc_content = self._generate_html_toc(document_structure) if output_config.include_toc else ""
            
            # 准备模板变量
            template_vars = {
                'title': self._extract_title(document_structure),
                'content': html_content,
                'css_content': self.base_css,
                'custom_css': output_config.custom_css or "",
                'toc_content': toc_content,
                'include_header': output_config.include_headers,
                'include_footer': output_config.include_footers,
                'include_toc': output_config.include_toc,
                'include_page_numbers': output_config.include_page_numbers,
                'body_class': f"template-{self.template_config.name}",
                'header_title': self.template_config.description,
                'header_meta': self._get_document_meta(document_structure),
                'footer_content': self._get_footer_content(),
                'javascript': self._get_javascript()
            }
            
            # 渲染HTML
            final_html = self.html_template.render(**template_vars)
            
            # 保存文件
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(final_html)
            
            logger.info(f"HTML文件生成成功: {output_path}")
            return True
            
        except Exception as e:
            logger.error(f"生成HTML失败: {e}")
            return False
    
    def generate_pdf(self, content: str, document_structure: DocumentStructure,
                    output_config: OutputConfig, output_path: str) -> bool:
        """生成PDF输出"""
        try:
            # 首先生成HTML
            temp_html_path = output_path.replace('.pdf', '_temp.html')
            
            if not self.generate_html(content, document_structure, output_config, temp_html_path):
                return False
            
            # 使用WeasyPrint转换为PDF
            html_doc = HTML(filename=temp_html_path)
            
            # PDF特定CSS
            pdf_css = CSS(string=self._get_pdf_css())
            
            # 生成PDF
            html_doc.write_pdf(
                output_path,
                stylesheets=[pdf_css]
            )
            
            # 清理临时文件
            if os.path.exists(temp_html_path):
                os.remove(temp_html_path)
            
            logger.info(f"PDF文件生成成功: {output_path}")
            return True
            
        except Exception as e:
            logger.error(f"生成PDF失败: {e}")
            # 清理临时文件
            temp_html_path = output_path.replace('.pdf', '_temp.html')
            if os.path.exists(temp_html_path):
                os.remove(temp_html_path)
            return False
    
    def _get_pdf_css(self) -> str:
        """获取PDF特定的CSS样式"""
        return """
        @page {
            size: A4;
            margin: 2.5cm;
            
            @top-center {
                content: "东北师范大学硕士学位论文";
                font-family: '黑体', SimHei, sans-serif;
                font-size: 12pt;
            }
            
            @bottom-center {
                content: counter(page);
                font-family: 'Times New Roman', serif;
                font-size: 10.5pt;
            }
        }
        
        /* 章节分页 */
        h1 {
            page-break-before: always;
        }
        
        /* 避免孤行寡行 */
        p {
            orphans: 3;
            widows: 3;
        }
        
        /* 表格和图片的分页控制 */
        table, img, pre {
            page-break-inside: avoid;
        }
        
        /* 目录页分页 */
        .table-of-contents {
            page-break-after: always;
        }
        
        /* 摘要页分页 */
        .academic-abstract {
            page-break-before: always;
            page-break-after: always;
        }
        """
    
    def generate_epub(self, content: str, document_structure: DocumentStructure,
                     output_config: OutputConfig, output_path: str) -> bool:
        """生成EPUB输出"""
        try:
            # EPUB需要专门的库，这里提供基本实现框架
            logger.warning("EPUB格式输出正在开发中")
            return False
        except Exception as e:
            logger.error(f"生成EPUB失败: {e}")
            return False
    
    def generate_latex(self, content: str, document_structure: DocumentStructure,
                      output_config: OutputConfig, output_path: str) -> bool:
        """生成LaTeX输出"""
        try:
            # LaTeX模板
            latex_template = self._get_latex_template()
            
            # 转换内容
            latex_content = self._markdown_to_latex(content, document_structure)
            
            # 填充模板
            final_latex = latex_template.format(
                title=self._extract_title(document_structure),
                author="作者",
                content=latex_content,
                documentclass=self._get_latex_documentclass()
            )
            
            # 保存文件
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(final_latex)
            
            logger.info(f"LaTeX文件生成成功: {output_path}")
            return True
            
        except Exception as e:
            logger.error(f"生成LaTeX失败: {e}")
            return False
    
    def _generate_html_toc(self, document_structure: DocumentStructure) -> str:
        """生成HTML目录"""
        toc_html = "<ul class='toc-list'>"
        
        for section in document_structure.sections:
            if section.section_type in [SectionType.CHAPTER, SectionType.SECTION, SectionType.SUBSECTION]:
                # 根据级别设置CSS类
                css_class = f"toc-level-{section.level}"
                anchor = self._generate_anchor(section.name)
                
                toc_html += f"""
                <li class="{css_class}">
                    <a href="#{anchor}">{section.name}</a>
                </li>
                """
        
        toc_html += "</ul>"
        return toc_html
    
    def _extract_title(self, document_structure: DocumentStructure) -> str:
        """提取文档标题"""
        for section in document_structure.sections:
            if section.level == 1 or section.section_type == SectionType.CHAPTER:
                return section.name
        return "文档"
    
    def _generate_anchor(self, text: str) -> str:
        """生成锚点ID"""
        import re
        # 移除特殊字符，转换为适合的ID
        anchor = re.sub(r'[^\w\u4e00-\u9fff]', '-', text)
        return anchor.strip('-')
    
    def _get_document_meta(self, document_structure: DocumentStructure) -> str:
        """获取文档元信息"""
        stats = document_structure.statistics
        return f"共 {stats['total_sections']} 章节 | {stats['total_words']} 字"
    
    def _get_footer_content(self) -> str:
        """获取页脚内容"""
        return "Generated by Enhanced Markdown Converter"
    
    def _get_javascript(self) -> str:
        """获取JavaScript代码"""
        return """
        // 页码显示
        function updatePageNumbers() {
            const pageNumbers = document.querySelectorAll('.page-current');
            pageNumbers.forEach(el => {
                el.textContent = '1';
            });
        }
        
        // 平滑滚动
        document.querySelectorAll('a[href^="#"]').forEach(anchor => {
            anchor.addEventListener('click', function (e) {
                e.preventDefault();
                const target = document.querySelector(this.getAttribute('href'));
                if (target) {
                    target.scrollIntoView({
                        behavior: 'smooth'
                    });
                }
            });
        });
        
        // 页面加载完成后执行
        document.addEventListener('DOMContentLoaded', function() {
            updatePageNumbers();
        });
        """
    
    def _get_latex_template(self) -> str:
        """获取LaTeX模板"""
        return r"""
\documentclass{{{documentclass}}}
\usepackage[UTF8]{{ctex}}
\usepackage{{geometry}}
\usepackage{{graphicx}}
\usepackage{{hyperref}}
\usepackage{{listings}}
\usepackage{{xcolor}}

\geometry{{a4paper, margin=2.5cm}}

\title{{{title}}}
\author{{{author}}}
\date{{\today}}

\begin{{document}}

\maketitle
\tableofcontents
\newpage

{content}

\end{{document}}
        """
    
    def _get_latex_documentclass(self) -> str:
        """获取LaTeX文档类"""
        if self.template_config.name == 'nenu_thesis':
            return "article"
        elif self.template_config.name == 'business_report':
            return "report"
        else:
            return "article"
    
    def _markdown_to_latex(self, content: str, document_structure: DocumentStructure) -> str:
        """将Markdown转换为LaTeX"""
        # 简单的Markdown到LaTeX转换
        import re
        
        latex_content = content
        
        # 标题转换
        latex_content = re.sub(r'^# (.+)$', r'\\section{\1}', latex_content, flags=re.MULTILINE)
        latex_content = re.sub(r'^## (.+)$', r'\\subsection{\1}', latex_content, flags=re.MULTILINE)
        latex_content = re.sub(r'^### (.+)$', r'\\subsubsection{\1}', latex_content, flags=re.MULTILINE)
        
        # 粗体和斜体
        latex_content = re.sub(r'\*\*(.+?)\*\*', r'\\textbf{\1}', latex_content)
        latex_content = re.sub(r'\*(.+?)\*', r'\\textit{\1}', latex_content)
        
        # 代码
        latex_content = re.sub(r'`(.+?)`', r'\\texttt{\1}', latex_content)
        
        return latex_content
    
    def batch_convert(self, input_files: List[str], output_format: OutputFormat,
                     output_dir: str, output_config: OutputConfig) -> Dict[str, bool]:
        """批量转换文件"""
        results = {}
        output_path = Path(output_dir)
        output_path.mkdir(parents=True, exist_ok=True)
        
        for input_file in input_files:
            try:
                # 读取文件内容
                with open(input_file, 'r', encoding='utf-8') as f:
                    content = f.read()
                
                # 分析文档结构
                from enhanced_document_analyzer import analyze_markdown_document
                document_structure = analyze_markdown_document(content)
                
                # 生成输出文件名
                input_path = Path(input_file)
                output_file = output_path / f"{input_path.stem}.{output_format.value}"
                
                # 转换文件
                success = self.convert_to_format(
                    content, document_structure, output_format,
                    str(output_file), output_config
                )
                
                results[input_file] = success
                
            except Exception as e:
                logger.error(f"转换文件 {input_file} 失败: {e}")
                results[input_file] = False
        
        success_count = sum(results.values())
        logger.info(f"批量转换完成: {success_count}/{len(input_files)} 个文件成功")
        
        return results
    
    def convert_to_format(self, content: str, document_structure: DocumentStructure,
                         output_format: OutputFormat, output_path: str,
                         output_config: OutputConfig) -> bool:
        """转换到指定格式"""
        if output_format == OutputFormat.HTML:
            return self.generate_html(content, document_structure, output_config, output_path)
        elif output_format == OutputFormat.PDF:
            return self.generate_pdf(content, document_structure, output_config, output_path)
        elif output_format == OutputFormat.EPUB:
            return self.generate_epub(content, document_structure, output_config, output_path)
        elif output_format == OutputFormat.LATEX:
            return self.generate_latex(content, document_structure, output_config, output_path)
        else:
            logger.error(f"不支持的输出格式: {output_format}")
            return False
    
    def get_supported_formats(self) -> List[OutputFormat]:
        """获取支持的输出格式"""
        return [OutputFormat.HTML, OutputFormat.PDF, OutputFormat.LATEX]
    
    def validate_output_config(self, config: OutputConfig) -> bool:
        """验证输出配置"""
        if config.format_type not in self.get_supported_formats():
            logger.error(f"不支持的格式: {config.format_type}")
            return False
        
        if config.quality not in ['low', 'medium', 'high']:
            logger.warning(f"无效的质量设置: {config.quality}，使用默认值 'high'")
            config.quality = 'high'
        
        return True


class OutputManager:
    """输出管理器"""
    
    def __init__(self, template_config: TemplateConfig):
        self.template_config = template_config
        self.generator = EnhancedOutputGenerator(template_config)
    
    def convert_document(self, input_file: str, output_format: OutputFormat,
                        output_file: str, config: Optional[OutputConfig] = None) -> bool:
        """转换单个文档"""
        if config is None:
            config = OutputConfig(format_type=output_format)
        
        if not self.generator.validate_output_config(config):
            return False
        
        try:
            # 读取输入文件
            with open(input_file, 'r', encoding='utf-8') as f:
                content = f.read()
            
            # 分析文档结构
            from enhanced_document_analyzer import analyze_markdown_document
            document_structure = analyze_markdown_document(content)
            
            # 转换文档
            return self.generator.convert_to_format(
                content, document_structure, output_format, output_file, config
            )
            
        except Exception as e:
            logger.error(f"转换文档失败: {e}")
            return False
    
    def batch_convert_directory(self, input_dir: str, output_dir: str,
                               output_format: OutputFormat,
                               config: Optional[OutputConfig] = None) -> Dict[str, bool]:
        """批量转换目录中的文档"""
        input_path = Path(input_dir)
        md_files = list(input_path.glob("**/*.md"))
        
        if not md_files:
            logger.warning(f"在目录 {input_dir} 中未找到Markdown文件")
            return {}
        
        return self.generator.batch_convert(
            [str(f) for f in md_files], output_format, output_dir, config or OutputConfig(output_format)
        )


# 便捷函数
def create_output_manager(template_config: TemplateConfig) -> OutputManager:
    """创建输出管理器的便捷函数"""
    return OutputManager(template_config)


def convert_markdown_to_format(input_file: str, output_format: str,
                              output_file: str, template_name: str = "default") -> bool:
    """转换Markdown到指定格式的便捷函数"""
    from enhanced_templates_config import get_template
    
    template_config = get_template(template_name)
    manager = OutputManager(template_config)
    
    format_enum = OutputFormat(output_format.lower())
    return manager.convert_document(input_file, format_enum, output_file)