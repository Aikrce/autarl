#!/usr/bin/env python3
"""
文档模板系统
提供多种预设模板，支持学术论文、商业报告、技术文档等格式
"""

from docx.shared import Pt, Cm, RGBColor, Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_LINE_SPACING, WD_TAB_ALIGNMENT
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.section import WD_SECTION_START, WD_ORIENTATION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.shared import OxmlElement, qn
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml
import re
import logging
from typing import Dict, Any, Optional

# 配置日志
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# 字体配置常量
FONT_CONFIGS = {
    'chinese': {
        'songti': '宋体',
        'heiti': '黑体',
        'yahei': '微软雅黑',
        'kaiti': '楷体',
        'fangsong': '仿宋'
    },
    'english': {
        'times': 'Times New Roman',
        'arial': 'Arial',
        'calibri': 'Calibri',
        'consolas': 'Consolas'
    }
}

# 字号配置常量
FONT_SIZES = {
    'chuhao': Pt(42),      # 初号
    'xiaochuhao': Pt(36),  # 小初号
    'yihao': Pt(26),       # 一号
    'xiaoyihao': Pt(24),   # 小一号
    'erhao': Pt(22),       # 二号
    'xiaoerhao': Pt(18),   # 小二号
    'sanhao': Pt(16),      # 三号
    'xiaosanhao': Pt(15),  # 小三号
    'sihao': Pt(14),       # 四号
    'xiaosihao': Pt(12),   # 小四号
    'wuhao': Pt(10.5),     # 五号
    'xiaowuhao': Pt(9),    # 小五号
    'liuhao': Pt(8),       # 六号
    'qihao': Pt(5.5)       # 七号
}

class DocumentTemplate:
    """文档模板基类"""
    def __init__(self):
        self.name = "默认模板"
        self.description = "标准Markdown转Word格式"
        self.category = "通用"
        self.version = "1.0"
        self.author = "系统"
        self.config = {}
    
    def apply_to_document(self, doc):
        """应用模板到文档"""
        try:
            logger.info(f"正在应用模板: {self.name}")
            self._apply_template_specific_settings(doc)
            logger.info(f"模板应用完成: {self.name}")
        except Exception as e:
            logger.error(f"模板应用失败: {self.name}, 错误: {str(e)}")
            raise
    
    def _apply_template_specific_settings(self, doc):
        """子类需要实现的具体模板设置"""
        pass
    
    def _setup_font(self, element, font_name: str, font_size: Pt, 
                   bold: bool = False, italic: bool = False, 
                   color: Optional[RGBColor] = None):
        """统一的字体设置方法"""
        try:
            if hasattr(element, 'font'):
                font = element.font
            else:
                font = element
            
            font.name = font_name
            font.size = font_size
            font.bold = bold
            font.italic = italic
            
            if color:
                font.color.rgb = color
            
            # 设置中英文字体
            if hasattr(element, '_element'):
                element._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
                if font_name in FONT_CONFIGS['english'].values():
                    element._element.rPr.rFonts.set(qn('w:ascii'), font_name)
                    element._element.rPr.rFonts.set(qn('w:hAnsi'), font_name)
                    
        except Exception as e:
            logger.warning(f"字体设置失败: {str(e)}")
    
    def _setup_paragraph(self, paragraph_format, alignment=WD_PARAGRAPH_ALIGNMENT.JUSTIFY,
                        first_line_indent=None, space_before=None, space_after=None,
                        line_spacing=1.5, line_spacing_rule=WD_LINE_SPACING.MULTIPLE):
        """统一的段落格式设置方法"""
        try:
            paragraph_format.alignment = alignment
            paragraph_format.line_spacing_rule = line_spacing_rule
            paragraph_format.line_spacing = line_spacing
            
            if first_line_indent is not None:
                paragraph_format.first_line_indent = first_line_indent
            if space_before is not None:
                paragraph_format.space_before = space_before
            if space_after is not None:
                paragraph_format.space_after = space_after
                
        except Exception as e:
            logger.warning(f"段落格式设置失败: {str(e)}")
    
    def _create_or_get_style(self, doc, style_name: str, style_type=WD_STYLE_TYPE.PARAGRAPH):
        """创建或获取样式"""
        try:
            return doc.styles[style_name]
        except KeyError:
            try:
                return doc.styles.add_style(style_name, style_type)
            except Exception as e:
                logger.warning(f"无法创建样式 {style_name}: {str(e)}")
                return None

class DefaultTemplate(DocumentTemplate):
    """默认模板"""
    
    def __init__(self):
        super().__init__()
        self.name = "默认"
        self.description = "标准格式，适合一般文档"
        self.category = "通用"
    
    def _apply_template_specific_settings(self, doc):
        """应用默认格式"""
        # 设置Normal样式
        normal_style = doc.styles['Normal']
        self._setup_font(normal_style, 'Microsoft YaHei', Pt(11))
        
        # 设置标题样式
        for i in range(1, 7):
            try:
                heading_style = doc.styles[f'Heading {i}']
                self._setup_font(
                    heading_style, 
                    'Microsoft YaHei', 
                    Pt(20 - i * 2), 
                    bold=True
                )
            except:
                continue

class ModernBusinessTemplate(DocumentTemplate):
    """现代商业文档模板"""
    
    def __init__(self):
        super().__init__()
        self.name = "商业文档"
        self.description = "现代商业报告和文档格式"
        self.category = "商业"
        
        self.config = {
            'fonts': {
                'main': FONT_CONFIGS['chinese']['yahei'],
                'english': FONT_CONFIGS['english']['calibri'],
                'heading': FONT_CONFIGS['chinese']['yahei']
            },
            'sizes': {
                'body': Pt(11),
                'heading1': Pt(18),
                'heading2': Pt(16),
                'heading3': Pt(14)
            },
            'colors': {
                'primary': RGBColor(44, 62, 80),
                'secondary': RGBColor(52, 152, 219),
                'accent': RGBColor(231, 76, 60)
            }
        }
    
    def _apply_template_specific_settings(self, doc):
        self._setup_modern_normal_style(doc)
        self._setup_modern_heading_styles(doc)
        self._setup_business_styles(doc)
    
    def _setup_modern_normal_style(self, doc):
        normal_style = doc.styles['Normal']
        self._setup_font(
            normal_style,
            self.config['fonts']['main'],
            self.config['sizes']['body'],
            color=self.config['colors']['primary']
        )
        
        self._setup_paragraph(
            normal_style.paragraph_format,
            alignment=WD_PARAGRAPH_ALIGNMENT.JUSTIFY,
            line_spacing=1.15,
            space_after=Pt(6)
        )
    
    def _setup_modern_heading_styles(self, doc):
        # 一级标题
        heading1 = doc.styles['Heading 1']
        self._setup_font(
            heading1,
            self.config['fonts']['heading'],
            self.config['sizes']['heading1'],
            bold=True,
            color=self.config['colors']['secondary']
        )
        
        self._setup_paragraph(
            heading1.paragraph_format,
            alignment=WD_PARAGRAPH_ALIGNMENT.LEFT,
            space_before=Pt(24),
            space_after=Pt(12)
        )
        
        # 二级标题
        heading2 = doc.styles['Heading 2']
        self._setup_font(
            heading2,
            self.config['fonts']['heading'],
            self.config['sizes']['heading2'],
            bold=True,
            color=self.config['colors']['primary']
        )
        
        self._setup_paragraph(
            heading2.paragraph_format,
            alignment=WD_PARAGRAPH_ALIGNMENT.LEFT,
            space_before=Pt(18),
            space_after=Pt(6)
        )
    
    def _setup_business_styles(self, doc):
        # 重点内容样式
        highlight = self._create_or_get_style(doc, 'Highlight')
        if highlight:
            self._setup_font(
                highlight,
                self.config['fonts']['main'],
                self.config['sizes']['body'],
                bold=True,
                color=self.config['colors']['accent']
            )

class TechnicalDocumentTemplate(DocumentTemplate):
    """技术文档模板"""
    
    def __init__(self):
        super().__init__()
        self.name = "技术文档"
        self.description = "技术手册和开发文档格式"
        self.category = "技术"
        
        self.config = {
            'fonts': {
                'main': FONT_CONFIGS['chinese']['yahei'],
                'code': FONT_CONFIGS['english']['consolas'],
                'heading': FONT_CONFIGS['chinese']['yahei']
            },
            'sizes': {
                'body': Pt(10),
                'code': Pt(9),
                'heading1': Pt(16),
                'heading2': Pt(14),
                'heading3': Pt(12)
            }
        }
    
    def _apply_template_specific_settings(self, doc):
        self._setup_technical_normal_style(doc)
        self._setup_technical_heading_styles(doc)
        self._setup_code_styles(doc)
    
    def _setup_technical_normal_style(self, doc):
        normal_style = doc.styles['Normal']
        self._setup_font(
            normal_style,
            self.config['fonts']['main'],
            self.config['sizes']['body']
        )
        
        self._setup_paragraph(
            normal_style.paragraph_format,
            alignment=WD_PARAGRAPH_ALIGNMENT.LEFT,
            line_spacing=1.2,
            space_after=Pt(3)
        )
    
    def _setup_technical_heading_styles(self, doc):
        # 简洁的标题样式
        for i in range(1, 4):
            try:
                heading = doc.styles[f'Heading {i}']
                self._setup_font(
                    heading,
                    self.config['fonts']['heading'],
                    self.config['sizes'][f'heading{i}'],
                    bold=True
                )
                
                self._setup_paragraph(
                    heading.paragraph_format,
                    alignment=WD_PARAGRAPH_ALIGNMENT.LEFT,
                    space_before=Pt(12),
                    space_after=Pt(6)
                )
            except:
                continue
    
    def _setup_code_styles(self, doc):
        # 代码块样式
        code_style = self._create_or_get_style(doc, 'Code Block')
        if code_style:
            self._setup_font(
                code_style,
                self.config['fonts']['code'],
                self.config['sizes']['code']
            )
            
            self._setup_paragraph(
                code_style.paragraph_format,
                alignment=WD_PARAGRAPH_ALIGNMENT.LEFT,
                line_spacing=1.0,
                line_spacing_rule=WD_LINE_SPACING.SINGLE
            )
            
            # 手动设置左缩进
            code_style.paragraph_format.left_indent = Cm(1)

class SimpleReportTemplate(DocumentTemplate):
    """简洁报告模板"""
    
    def __init__(self):
        super().__init__()
        self.name = "简洁报告"
        self.description = "简洁清晰的报告格式"
        self.category = "报告"
        
        self.config = {
            'fonts': {
                'main': FONT_CONFIGS['chinese']['songti'],
                'heading': FONT_CONFIGS['chinese']['heiti']
            },
            'sizes': {
                'body': FONT_SIZES['xiaosihao'],
                'heading1': FONT_SIZES['sanhao'],
                'heading2': FONT_SIZES['sihao']
            }
        }
    
    def _apply_template_specific_settings(self, doc):
        self._setup_simple_normal_style(doc)
        self._setup_simple_heading_styles(doc)
    
    def _setup_simple_normal_style(self, doc):
        normal_style = doc.styles['Normal']
        self._setup_font(
            normal_style,
            self.config['fonts']['main'],
            self.config['sizes']['body']
        )
        
        self._setup_paragraph(
            normal_style.paragraph_format,
            alignment=WD_PARAGRAPH_ALIGNMENT.JUSTIFY,
            first_line_indent=Cm(0.74),
            line_spacing=1.5
        )
    
    def _setup_simple_heading_styles(self, doc):
        # 一级标题
        heading1 = doc.styles['Heading 1']
        self._setup_font(
            heading1,
            self.config['fonts']['heading'],
            self.config['sizes']['heading1'],
            bold=True
        )
        
        self._setup_paragraph(
            heading1.paragraph_format,
            alignment=WD_PARAGRAPH_ALIGNMENT.CENTER,
            space_before=Pt(24),
            space_after=Pt(12)
        )
        
        # 二级标题
        heading2 = doc.styles['Heading 2']
        self._setup_font(
            heading2,
            self.config['fonts']['heading'],
            self.config['sizes']['heading2'],
            bold=True
        )
        
        self._setup_paragraph(
            heading2.paragraph_format,
            alignment=WD_PARAGRAPH_ALIGNMENT.LEFT,
            space_before=Pt(12),
            space_after=Pt(6)
        )

# 导入东北师范大学论文模板
# from nenu_thesis_template import NENUThesisTemplate

# 优化的模板注册表
TEMPLATES = {
    'default': DefaultTemplate(),
    # 'nenu_thesis': NENUThesisTemplate(),
    'business': ModernBusinessTemplate(),
    'technical': TechnicalDocumentTemplate(),
    'simple_report': SimpleReportTemplate(),
    # 向后兼容
    # 'university_thesis': NENUThesisTemplate(),
    # 'graduation_thesis': NENUThesisTemplate(),
}

# 模板分类
TEMPLATE_CATEGORIES = {
    '通用': ['default'],
    '学术论文': [],  # ['nenu_thesis', 'university_thesis', 'graduation_thesis'],
    '商业': ['business'],
    '技术': ['technical'],
    '报告': ['simple_report']
}

class TemplateManager:
    """模板管理器"""
    
    @staticmethod
    def get_template(name: str) -> DocumentTemplate:
        """获取模板"""
        template = TEMPLATES.get(name)
        if template is None:
            logger.warning(f"未找到模板 '{name}'，使用默认模板")
            return TEMPLATES['default']
        return template
    
    @staticmethod
    def list_templates() -> Dict[str, str]:
        """列出所有可用模板"""
        return {name: template.description for name, template in TEMPLATES.items()}
    
    @staticmethod
    def list_templates_by_category() -> Dict[str, Dict[str, str]]:
        """按分类列出模板"""
        result = {}
        for category, template_names in TEMPLATE_CATEGORIES.items():
            result[category] = {}
            for name in template_names:
                if name in TEMPLATES:
                    result[category][name] = TEMPLATES[name].description
        return result
    
    @staticmethod
    def get_template_info(name: str) -> Dict[str, Any]:
        """获取模板详细信息"""
        template = TEMPLATES.get(name)
        if template is None:
            return None
        
        return {
            'name': template.name,
            'description': template.description,
            'category': getattr(template, 'category', '未分类'),
            'version': getattr(template, 'version', '1.0'),
            'author': getattr(template, 'author', '系统'),
            'config': getattr(template, 'config', {})
        }
    
    @staticmethod
    def register_template(key: str, template: DocumentTemplate, category: str = '自定义'):
        """注册新模板"""
        TEMPLATES[key] = template
        if category not in TEMPLATE_CATEGORIES:
            TEMPLATE_CATEGORIES[category] = []
        if key not in TEMPLATE_CATEGORIES[category]:
            TEMPLATE_CATEGORIES[category].append(key)
        logger.info(f"已注册模板: {key} ({template.name})")

# 向后兼容的函数
def get_template(name):
    """获取模板（向后兼容）"""
    return TemplateManager.get_template(name)

def list_templates():
    """列出所有可用模板（向后兼容）"""
    return TemplateManager.list_templates()