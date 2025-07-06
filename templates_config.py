#!/usr/bin/env python3

from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_LINE_SPACING
from docx.enum.style import WD_STYLE_TYPE

class DocumentTemplate:
    """文档模板基类"""
    def __init__(self):
        self.name = "默认模板"
        self.description = "标准Markdown转Word格式"
    
    def apply_to_document(self, doc):
        """应用模板到文档"""
        pass

class GraduationThesisTemplate(DocumentTemplate):
    """东北师范大学毕业论文模板"""
    
    def __init__(self):
        super().__init__()
        self.name = "毕业论文"
        self.description = "东北师范大学研究生学位论文格式规范"
    
    def apply_to_document(self, doc):
        """应用毕业论文格式到文档"""
        # 设置页面边距和页面设置
        self._setup_page_margins(doc)
        
        # 设置基础样式
        self._setup_base_styles(doc)
        
        # 设置标题样式
        self._setup_heading_styles(doc)
        
        # 设置段落样式
        self._setup_paragraph_styles(doc)
        
        # 设置列表样式
        self._setup_list_styles(doc)
        
        # 设置表格样式
        self._setup_table_styles(doc)
    
    def _setup_page_margins(self, doc):
        """设置页面边距"""
        section = doc.sections[0]
        
        # 页边距设置：上下2cm，左右2.5cm
        section.top_margin = Cm(2.0)
        section.bottom_margin = Cm(2.0)
        section.left_margin = Cm(2.5)
        section.right_margin = Cm(2.5)
        
        # 页眉页脚边距
        section.header_distance = Cm(1.5)
        section.footer_distance = Cm(1.75)
        
        # 设置页眉
        header = section.header
        header_para = header.paragraphs[0]
        header_para.text = "东北师范大学硕士学位论文"
        header_para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        
        # 设置页眉字体
        run = header_para.runs[0]
        run.font.name = '黑体'
        run.font.size = Pt(12)  # 小四号
        run._element.rPr.rFonts.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia', '黑体')
    
    def _setup_base_styles(self, doc):
        """设置基础样式"""
        # 修改Normal样式
        normal_style = doc.styles['Normal']
        normal_font = normal_style.font
        normal_font.name = '宋体'
        normal_font.size = Pt(12)  # 小四号
        normal_font.color.rgb = RGBColor(0, 0, 0)
        normal_style._element.rPr.rFonts.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia', '宋体')
        
        # 设置段落格式
        normal_paragraph_format = normal_style.paragraph_format
        normal_paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY  # 两端对齐
        normal_paragraph_format.first_line_indent = Cm(0.74)  # 首行缩进2字符
        normal_paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        normal_paragraph_format.line_spacing = 1.5  # 1.5倍行距
        normal_paragraph_format.space_before = Pt(0)
        normal_paragraph_format.space_after = Pt(0)
    
    def _setup_heading_styles(self, doc):
        """设置标题样式"""
        # 章标题 (Heading 1) - 三号黑体，居中
        heading1 = doc.styles['Heading 1']
        heading1_font = heading1.font
        heading1_font.name = '黑体'
        heading1_font.size = Pt(16)  # 三号
        heading1_font.bold = True
        heading1_font.color.rgb = RGBColor(0, 0, 0)
        heading1._element.rPr.rFonts.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia', '黑体')
        
        heading1_paragraph = heading1.paragraph_format
        heading1_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        heading1_paragraph.space_before = Pt(48)  # 段前48磅
        heading1_paragraph.space_after = Pt(24)   # 段后24磅
        heading1_paragraph.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        heading1_paragraph.line_spacing = 1.5
        heading1_paragraph.first_line_indent = Pt(0)
        heading1_paragraph.page_break_before = True  # 每章另起一页
        
        # 二级标题 (Heading 2) - 四号黑体，两端对齐
        heading2 = doc.styles['Heading 2']
        heading2_font = heading2.font
        heading2_font.name = '黑体'
        heading2_font.size = Pt(14)  # 四号
        heading2_font.bold = True
        heading2_font.color.rgb = RGBColor(0, 0, 0)
        heading2._element.rPr.rFonts.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia', '黑体')
        
        heading2_paragraph = heading2.paragraph_format
        heading2_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
        heading2_paragraph.space_before = Pt(6)
        heading2_paragraph.space_after = Pt(0)
        heading2_paragraph.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        heading2_paragraph.line_spacing = 1.5
        heading2_paragraph.first_line_indent = Pt(0)
        
        # 三级标题 (Heading 3) - 小四号宋体加粗，两端对齐
        heading3 = doc.styles['Heading 3']
        heading3_font = heading3.font
        heading3_font.name = '宋体'
        heading3_font.size = Pt(12)  # 小四号
        heading3_font.bold = True
        heading3_font.color.rgb = RGBColor(0, 0, 0)
        heading3._element.rPr.rFonts.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia', '宋体')
        
        heading3_paragraph = heading3.paragraph_format
        heading3_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
        heading3_paragraph.space_before = Pt(6)
        heading3_paragraph.space_after = Pt(0)
        heading3_paragraph.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        heading3_paragraph.line_spacing = 1.5
        heading3_paragraph.first_line_indent = Pt(0)
    
    def _setup_paragraph_styles(self, doc):
        """设置段落样式"""
        # 创建摘要样式
        try:
            abstract_style = doc.styles.add_style('Abstract Title', WD_STYLE_TYPE.PARAGRAPH)
        except:
            abstract_style = doc.styles['Abstract Title']
        
        abstract_font = abstract_style.font
        abstract_font.name = '黑体'
        abstract_font.size = Pt(16)  # 三号
        abstract_font.bold = True
        abstract_style._element.rPr.rFonts.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia', '黑体')
        
        abstract_paragraph = abstract_style.paragraph_format
        abstract_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        abstract_paragraph.space_before = Pt(48)
        abstract_paragraph.space_after = Pt(24)
        abstract_paragraph.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        abstract_paragraph.line_spacing = 1.5
        
        # 创建关键词样式
        try:
            keywords_style = doc.styles.add_style('Keywords', WD_STYLE_TYPE.PARAGRAPH)
        except:
            keywords_style = doc.styles['Keywords']
        
        keywords_font = keywords_style.font
        keywords_font.name = '宋体'
        keywords_font.size = Pt(12)  # 小四号
        keywords_style._element.rPr.rFonts.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia', '宋体')
        
        keywords_paragraph = keywords_style.paragraph_format
        keywords_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
        keywords_paragraph.first_line_indent = Cm(0.74)  # 首行缩进2字符
        keywords_paragraph.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        keywords_paragraph.line_spacing = 1.5
        keywords_paragraph.space_before = Pt(0)
        keywords_paragraph.space_after = Pt(0)
    
    def _setup_list_styles(self, doc):
        """设置列表样式"""
        # 无序列表样式
        try:
            bullet_style = doc.styles['List Bullet']
        except:
            return
        
        bullet_font = bullet_style.font
        bullet_font.name = '宋体'
        bullet_font.size = Pt(12)  # 小四号
        bullet_style._element.rPr.rFonts.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia', '宋体')
        
        bullet_paragraph = bullet_style.paragraph_format
        bullet_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
        bullet_paragraph.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        bullet_paragraph.line_spacing = 1.5
        bullet_paragraph.space_before = Pt(0)
        bullet_paragraph.space_after = Pt(0)
        
        # 有序列表样式
        try:
            number_style = doc.styles['List Number']
        except:
            return
        
        number_font = number_style.font
        number_font.name = '宋体'
        number_font.size = Pt(12)  # 小四号
        number_style._element.rPr.rFonts.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia', '宋体')
        
        number_paragraph = number_style.paragraph_format
        number_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
        number_paragraph.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        number_paragraph.line_spacing = 1.5
        number_paragraph.space_before = Pt(0)
        number_paragraph.space_after = Pt(0)
    
    def _setup_table_styles(self, doc):
        """设置表格样式"""
        # 表格使用三线表格式，这个会在创建表格时单独处理
        pass

class DefaultTemplate(DocumentTemplate):
    """默认模板"""
    
    def __init__(self):
        super().__init__()
        self.name = "默认"
        self.description = "标准格式，适合一般文档"
    
    def apply_to_document(self, doc):
        """应用默认格式"""
        # 设置Normal样式
        normal_style = doc.styles['Normal']
        normal_font = normal_style.font
        normal_font.name = 'Microsoft YaHei'
        normal_font.size = Pt(11)
        
        # 设置标题样式
        for i in range(1, 7):
            try:
                heading_style = doc.styles[f'Heading {i}']
                heading_style.font.name = 'Microsoft YaHei'
                heading_style.font.bold = True
                heading_style.font.size = Pt(20 - i * 2)
            except:
                pass

# 模板注册表
TEMPLATES = {
    'default': DefaultTemplate(),
    'graduation_thesis': GraduationThesisTemplate(),
}

def get_template(name):
    """获取模板"""
    return TEMPLATES.get(name, TEMPLATES['default'])

def list_templates():
    """列出所有可用模板"""
    return {name: template.description for name, template in TEMPLATES.items()}