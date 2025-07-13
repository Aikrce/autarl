#!/usr/bin/env python3
"""
Enhanced Template Configuration System
支持JSON/YAML驱动的可扩展模板配置
"""

import json
import yaml
from pathlib import Path
from typing import Dict, Any, List, Optional
from dataclasses import dataclass, asdict
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.style import WD_STYLE_TYPE
import logging

logger = logging.getLogger(__name__)


@dataclass
class FontConfig:
    """字体配置"""
    name: str = "宋体"
    size_pt: float = 12.0
    bold: bool = False
    italic: bool = False
    color_rgb: tuple = (0, 0, 0)
    
    def to_dict(self):
        return asdict(self)


@dataclass
class ParagraphConfig:
    """段落配置"""
    alignment: str = "justify"  # left, center, right, justify
    space_before_pt: float = 0.0
    space_after_pt: float = 0.0
    line_spacing: float = 1.0
    first_line_indent_cm: float = 0.0
    left_indent_cm: float = 0.0
    right_indent_cm: float = 0.0
    hanging_indent_cm: float = 0.0
    
    def to_dict(self):
        return asdict(self)


@dataclass
class StyleConfig:
    """样式配置"""
    name: str
    style_type: str = "paragraph"  # paragraph, character, table, list
    font: FontConfig = None
    paragraph: ParagraphConfig = None
    based_on: Optional[str] = None
    next_style: Optional[str] = None
    
    def __post_init__(self):
        if self.font is None:
            self.font = FontConfig()
        if self.paragraph is None:
            self.paragraph = ParagraphConfig()
    
    def to_dict(self):
        return {
            'name': self.name,
            'style_type': self.style_type,
            'font': self.font.to_dict(),
            'paragraph': self.paragraph.to_dict(),
            'based_on': self.based_on,
            'next_style': self.next_style
        }


@dataclass
class ComponentMapping:
    """组件映射配置"""
    component_type: str
    style_name: str
    requires_page_break: bool = False
    custom_formatting: Optional[Dict[str, Any]] = None
    
    def to_dict(self):
        return asdict(self)


@dataclass
class TemplateConfig:
    """模板配置"""
    name: str
    description: str
    version: str = "1.0"
    author: str = ""
    
    # 页面设置
    page_margins_cm: Dict[str, float] = None
    page_orientation: str = "portrait"  # portrait, landscape
    page_size: str = "A4"
    
    # 样式定义
    styles: List[StyleConfig] = None
    
    # 组件映射
    component_mappings: List[ComponentMapping] = None
    
    # 页眉页脚配置
    header_config: Optional[Dict[str, Any]] = None
    footer_config: Optional[Dict[str, Any]] = None
    
    # 自定义规则
    custom_rules: Optional[Dict[str, Any]] = None
    
    def __post_init__(self):
        if self.page_margins_cm is None:
            self.page_margins_cm = {"top": 2.5, "bottom": 2.5, "left": 2.5, "right": 2.5}
        if self.styles is None:
            self.styles = []
        if self.component_mappings is None:
            self.component_mappings = []
    
    def to_dict(self):
        return {
            'name': self.name,
            'description': self.description,
            'version': self.version,
            'author': self.author,
            'page_margins_cm': self.page_margins_cm,
            'page_orientation': self.page_orientation,
            'page_size': self.page_size,
            'styles': [style.to_dict() for style in self.styles],
            'component_mappings': [mapping.to_dict() for mapping in self.component_mappings],
            'header_config': self.header_config,
            'footer_config': self.footer_config,
            'custom_rules': self.custom_rules
        }


class EnhancedTemplateManager:
    """增强的模板管理器"""
    
    def __init__(self, templates_dir: str = "templates"):
        self.templates_dir = Path(templates_dir)
        self.templates_dir.mkdir(exist_ok=True)
        self._templates_cache: Dict[str, TemplateConfig] = {}
        self._load_builtin_templates()
    
    def _load_builtin_templates(self):
        """加载内置模板"""
        # 默认模板
        default_template = self._create_default_template()
        self._templates_cache['default'] = default_template
        
        # 东北师大论文模板
        nenu_template = self._create_nenu_thesis_template()
        self._templates_cache['nenu_thesis'] = nenu_template
        
        # 商务报告模板
        business_template = self._create_business_report_template()
        self._templates_cache['business_report'] = business_template
        
        # 技术文档模板
        technical_template = self._create_technical_doc_template()
        self._templates_cache['technical_doc'] = technical_template
    
    def _create_default_template(self) -> TemplateConfig:
        """创建默认模板"""
        template = TemplateConfig(
            name="default",
            description="默认模板 - 简洁的文档格式",
            author="Enhanced Converter"
        )
        
        # 定义基础样式
        template.styles = [
            StyleConfig(
                name="Normal",
                font=FontConfig(name="宋体", size_pt=12),
                paragraph=ParagraphConfig(
                    alignment="justify",
                    line_spacing=1.5,
                    first_line_indent_cm=0.74
                )
            ),
            StyleConfig(
                name="Heading 1",
                font=FontConfig(name="黑体", size_pt=16, bold=True),
                paragraph=ParagraphConfig(
                    alignment="center",
                    space_before_pt=24,
                    space_after_pt=18,
                    line_spacing=1.5
                )
            ),
            StyleConfig(
                name="Heading 2",
                font=FontConfig(name="黑体", size_pt=14, bold=True),
                paragraph=ParagraphConfig(
                    alignment="justify",
                    space_before_pt=12,
                    space_after_pt=6,
                    line_spacing=1.5
                )
            ),
            StyleConfig(
                name="Code Block",
                font=FontConfig(name="Consolas", size_pt=10),
                paragraph=ParagraphConfig(
                    alignment="left",
                    left_indent_cm=1.0,
                    space_before_pt=6,
                    space_after_pt=6
                )
            )
        ]
        
        # 组件映射
        template.component_mappings = [
            ComponentMapping("abstract_cn", "Abstract CN"),
            ComponentMapping("abstract_en", "Abstract EN"),
            ComponentMapping("references", "Reference", requires_page_break=True),
            ComponentMapping("appendix", "Appendix", requires_page_break=True)
        ]
        
        return template
    
    def _create_nenu_thesis_template(self) -> TemplateConfig:
        """创建东北师大论文模板"""
        template = TemplateConfig(
            name="nenu_thesis",
            description="东北师范大学硕士学位论文模板",
            author="NENU Graduate School"
        )
        
        # 页面设置
        template.page_margins_cm = {"top": 2.5, "bottom": 2.0, "left": 2.5, "right": 2.0}
        
        # 样式定义
        template.styles = [
            # 正文样式
            StyleConfig(
                name="Normal",
                font=FontConfig(name="宋体", size_pt=12),
                paragraph=ParagraphConfig(
                    alignment="justify",
                    line_spacing=1.5,
                    first_line_indent_cm=0.74
                )
            ),
            
            # 章标题
            StyleConfig(
                name="Chapter Title",
                font=FontConfig(name="黑体", size_pt=16, bold=True),
                paragraph=ParagraphConfig(
                    alignment="center",
                    space_before_pt=48,
                    space_after_pt=24,
                    line_spacing=1.5
                )
            ),
            
            # 二级标题
            StyleConfig(
                name="Section Title",
                font=FontConfig(name="黑体", size_pt=14, bold=True),
                paragraph=ParagraphConfig(
                    alignment="justify",
                    space_before_pt=6,
                    space_after_pt=0,
                    line_spacing=1.5
                )
            ),
            
            # 三级标题
            StyleConfig(
                name="Subsection Title",
                font=FontConfig(name="宋体", size_pt=12, bold=True),
                paragraph=ParagraphConfig(
                    alignment="justify",
                    space_before_pt=6,
                    space_after_pt=0,
                    line_spacing=1.5
                )
            ),
            
            # 摘要标题
            StyleConfig(
                name="Abstract Title CN",
                font=FontConfig(name="黑体", size_pt=16, bold=True),
                paragraph=ParagraphConfig(
                    alignment="center",
                    space_before_pt=48,
                    space_after_pt=24,
                    line_spacing=1.5
                )
            ),
            
            StyleConfig(
                name="Abstract Title EN",
                font=FontConfig(name="Times New Roman", size_pt=16, bold=True),
                paragraph=ParagraphConfig(
                    alignment="center",
                    space_before_pt=48,
                    space_after_pt=24,
                    line_spacing=1.5
                )
            ),
            
            # 摘要正文
            StyleConfig(
                name="Abstract Body CN",
                font=FontConfig(name="宋体", size_pt=12),
                paragraph=ParagraphConfig(
                    alignment="justify",
                    first_line_indent_cm=0.74,
                    line_spacing=1.5
                )
            ),
            
            StyleConfig(
                name="Abstract Body EN",
                font=FontConfig(name="Times New Roman", size_pt=12),
                paragraph=ParagraphConfig(
                    alignment="justify",
                    first_line_indent_cm=0.74,
                    line_spacing=1.5
                )
            ),
            
            # 关键词
            StyleConfig(
                name="Keywords CN",
                font=FontConfig(name="宋体", size_pt=12, bold=True),
                paragraph=ParagraphConfig(
                    alignment="justify",
                    first_line_indent_cm=0.74,
                    line_spacing=1.5
                )
            ),
            
            StyleConfig(
                name="Keywords EN",
                font=FontConfig(name="Times New Roman", size_pt=12, bold=True),
                paragraph=ParagraphConfig(
                    alignment="justify",
                    hanging_indent_cm=2.1,
                    line_spacing=1.5
                )
            ),
            
            # 目录样式
            StyleConfig(
                name="TOC Title",
                font=FontConfig(name="黑体", size_pt=16, bold=True),
                paragraph=ParagraphConfig(
                    alignment="center",
                    space_before_pt=48,
                    space_after_pt=24,
                    line_spacing=1.5
                )
            ),
            
            StyleConfig(
                name="TOC Level 1",
                font=FontConfig(name="黑体", size_pt=12),
                paragraph=ParagraphConfig(
                    alignment="justify",
                    line_spacing=1.5
                )
            ),
            
            StyleConfig(
                name="TOC Level 2",
                font=FontConfig(name="宋体", size_pt=12),
                paragraph=ParagraphConfig(
                    alignment="justify",
                    left_indent_cm=0.37,
                    line_spacing=1.5
                )
            ),
            
            StyleConfig(
                name="TOC Level 3",
                font=FontConfig(name="宋体", size_pt=12),
                paragraph=ParagraphConfig(
                    alignment="justify",
                    left_indent_cm=0.74,
                    line_spacing=1.5
                )
            ),
            
            # 参考文献
            StyleConfig(
                name="Reference Title",
                font=FontConfig(name="黑体", size_pt=16, bold=True),
                paragraph=ParagraphConfig(
                    alignment="center",
                    space_before_pt=48,
                    space_after_pt=24,
                    line_spacing=1.5
                )
            ),
            
            StyleConfig(
                name="Reference Content",
                font=FontConfig(name="宋体", size_pt=12),
                paragraph=ParagraphConfig(
                    alignment="justify",
                    hanging_indent_cm=0.6,
                    line_spacing=1.0
                )
            )
        ]
        
        # 组件映射
        template.component_mappings = [
            ComponentMapping("cover_page", "Cover Page", requires_page_break=True),
            ComponentMapping("english_cover", "English Cover", requires_page_break=True),
            ComponentMapping("declaration", "Declaration", requires_page_break=True),
            ComponentMapping("authorization", "Authorization", requires_page_break=True),
            ComponentMapping("abstract_cn", "Abstract Title CN", requires_page_break=True),
            ComponentMapping("abstract_en", "Abstract Title EN", requires_page_break=True),
            ComponentMapping("toc", "TOC Title", requires_page_break=True),
            ComponentMapping("introduction", "Chapter Title", requires_page_break=True),
            ComponentMapping("literature_review", "Chapter Title", requires_page_break=True),
            ComponentMapping("methodology", "Chapter Title", requires_page_break=True),
            ComponentMapping("results", "Chapter Title", requires_page_break=True),
            ComponentMapping("discussion", "Chapter Title", requires_page_break=True),
            ComponentMapping("conclusion", "Chapter Title", requires_page_break=True),
            ComponentMapping("references", "Reference Title", requires_page_break=True),
            ComponentMapping("appendix", "Appendix Title", requires_page_break=True),
            ComponentMapping("acknowledgments", "Chapter Title", requires_page_break=True)
        ]
        
        # 页眉页脚配置
        template.header_config = {
            "text": "东北师范大学硕士学位论文",
            "font": {"name": "黑体", "size_pt": 12},
            "alignment": "center"
        }
        
        template.footer_config = {
            "show_page_number": True,
            "font": {"name": "Times New Roman", "size_pt": 10.5},
            "alignment": "center"
        }
        
        return template
    
    def _create_business_report_template(self) -> TemplateConfig:
        """创建商务报告模板"""
        template = TemplateConfig(
            name="business_report",
            description="商务报告模板 - 专业商务文档格式",
            author="Business Template"
        )
        
        template.styles = [
            StyleConfig(
                name="Normal",
                font=FontConfig(name="微软雅黑", size_pt=11),
                paragraph=ParagraphConfig(
                    alignment="justify",
                    line_spacing=1.15,
                    space_after_pt=6
                )
            ),
            StyleConfig(
                name="Title",
                font=FontConfig(name="微软雅黑", size_pt=20, bold=True, color_rgb=(0, 51, 102)),
                paragraph=ParagraphConfig(
                    alignment="center",
                    space_before_pt=36,
                    space_after_pt=24
                )
            ),
            StyleConfig(
                name="Heading 1",
                font=FontConfig(name="微软雅黑", size_pt=16, bold=True, color_rgb=(0, 51, 102)),
                paragraph=ParagraphConfig(
                    alignment="left",
                    space_before_pt=18,
                    space_after_pt=12
                )
            ),
            StyleConfig(
                name="Heading 2",
                font=FontConfig(name="微软雅黑", size_pt=14, bold=True, color_rgb=(51, 102, 153)),
                paragraph=ParagraphConfig(
                    alignment="left",
                    space_before_pt=12,
                    space_after_pt=6
                )
            )
        ]
        
        return template
    
    def _create_technical_doc_template(self) -> TemplateConfig:
        """创建技术文档模板"""
        template = TemplateConfig(
            name="technical_doc",
            description="技术文档模板 - IT技术文档格式",
            author="Tech Template"
        )
        
        template.styles = [
            StyleConfig(
                name="Normal",
                font=FontConfig(name="Source Han Sans CN", size_pt=11),
                paragraph=ParagraphConfig(
                    alignment="justify",
                    line_spacing=1.3
                )
            ),
            StyleConfig(
                name="Code Block",
                font=FontConfig(name="JetBrains Mono", size_pt=9),
                paragraph=ParagraphConfig(
                    alignment="left",
                    left_indent_cm=1.0,
                    space_before_pt=6,
                    space_after_pt=6
                )
            ),
            StyleConfig(
                name="API Reference",
                font=FontConfig(name="Consolas", size_pt=10),
                paragraph=ParagraphConfig(
                    alignment="left",
                    hanging_indent_cm=1.5
                )
            )
        ]
        
        return template
    
    def get_template(self, name: str) -> TemplateConfig:
        """获取模板配置"""
        if name in self._templates_cache:
            return self._templates_cache[name]
        
        # 尝试从文件加载
        template_file = self.templates_dir / f"{name}.json"
        if template_file.exists():
            with open(template_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
                template = self._dict_to_template_config(data)
                self._templates_cache[name] = template
                return template
        
        # 尝试从YAML文件加载
        yaml_file = self.templates_dir / f"{name}.yaml"
        if yaml_file.exists():
            with open(yaml_file, 'r', encoding='utf-8') as f:
                data = yaml.safe_load(f)
                template = self._dict_to_template_config(data)
                self._templates_cache[name] = template
                return template
        
        logger.warning(f"Template '{name}' not found, using default")
        return self._templates_cache['default']
    
    def save_template(self, template: TemplateConfig, format_type: str = "json"):
        """保存模板配置到文件"""
        if format_type == "json":
            file_path = self.templates_dir / f"{template.name}.json"
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(template.to_dict(), f, indent=2, ensure_ascii=False)
        elif format_type == "yaml":
            file_path = self.templates_dir / f"{template.name}.yaml"
            with open(file_path, 'w', encoding='utf-8') as f:
                yaml.dump(template.to_dict(), f, default_flow_style=False, allow_unicode=True)
        
        # 更新缓存
        self._templates_cache[template.name] = template
        logger.info(f"Template '{template.name}' saved to {file_path}")
    
    def list_templates(self) -> Dict[str, str]:
        """列出所有可用模板"""
        templates = {}
        
        # 添加缓存中的模板
        for name, template in self._templates_cache.items():
            templates[name] = template.description
        
        # 扫描模板目录
        for file_path in self.templates_dir.glob("*.json"):
            name = file_path.stem
            if name not in templates:
                try:
                    with open(file_path, 'r', encoding='utf-8') as f:
                        data = json.load(f)
                        templates[name] = data.get('description', 'External template')
                except Exception as e:
                    logger.warning(f"Failed to load template {name}: {e}")
        
        for file_path in self.templates_dir.glob("*.yaml"):
            name = file_path.stem
            if name not in templates:
                try:
                    with open(file_path, 'r', encoding='utf-8') as f:
                        data = yaml.safe_load(f)
                        templates[name] = data.get('description', 'External YAML template')
                except Exception as e:
                    logger.warning(f"Failed to load YAML template {name}: {e}")
        
        return templates
    
    def _dict_to_template_config(self, data: Dict[str, Any]) -> TemplateConfig:
        """将字典转换为TemplateConfig对象"""
        # 转换样式配置
        styles = []
        for style_data in data.get('styles', []):
            font_data = style_data.get('font', {})
            paragraph_data = style_data.get('paragraph', {})
            
            font = FontConfig(
                name=font_data.get('name', '宋体'),
                size_pt=font_data.get('size_pt', 12),
                bold=font_data.get('bold', False),
                italic=font_data.get('italic', False),
                color_rgb=tuple(font_data.get('color_rgb', [0, 0, 0]))
            )
            
            paragraph = ParagraphConfig(
                alignment=paragraph_data.get('alignment', 'justify'),
                space_before_pt=paragraph_data.get('space_before_pt', 0),
                space_after_pt=paragraph_data.get('space_after_pt', 0),
                line_spacing=paragraph_data.get('line_spacing', 1.0),
                first_line_indent_cm=paragraph_data.get('first_line_indent_cm', 0),
                left_indent_cm=paragraph_data.get('left_indent_cm', 0),
                right_indent_cm=paragraph_data.get('right_indent_cm', 0),
                hanging_indent_cm=paragraph_data.get('hanging_indent_cm', 0)
            )
            
            style = StyleConfig(
                name=style_data['name'],
                style_type=style_data.get('style_type', 'paragraph'),
                font=font,
                paragraph=paragraph,
                based_on=style_data.get('based_on'),
                next_style=style_data.get('next_style')
            )
            styles.append(style)
        
        # 转换组件映射
        mappings = []
        for mapping_data in data.get('component_mappings', []):
            mapping = ComponentMapping(
                component_type=mapping_data['component_type'],
                style_name=mapping_data['style_name'],
                requires_page_break=mapping_data.get('requires_page_break', False),
                custom_formatting=mapping_data.get('custom_formatting')
            )
            mappings.append(mapping)
        
        template = TemplateConfig(
            name=data['name'],
            description=data['description'],
            version=data.get('version', '1.0'),
            author=data.get('author', ''),
            page_margins_cm=data.get('page_margins_cm', {"top": 2.5, "bottom": 2.5, "left": 2.5, "right": 2.5}),
            page_orientation=data.get('page_orientation', 'portrait'),
            page_size=data.get('page_size', 'A4'),
            styles=styles,
            component_mappings=mappings,
            header_config=data.get('header_config'),
            footer_config=data.get('footer_config'),
            custom_rules=data.get('custom_rules')
        )
        
        return template
    
    def create_custom_template(self, name: str, description: str, base_template: str = "default") -> TemplateConfig:
        """基于现有模板创建自定义模板"""
        base = self.get_template(base_template)
        
        # 深拷贝基础模板
        import copy
        custom_template = copy.deepcopy(base)
        custom_template.name = name
        custom_template.description = description
        custom_template.version = "1.0"
        custom_template.author = "Custom"
        
        return custom_template


# 全局模板管理器实例
template_manager = EnhancedTemplateManager()


def get_template(name: str) -> TemplateConfig:
    """获取模板配置的便捷函数"""
    return template_manager.get_template(name)


def list_templates() -> Dict[str, str]:
    """列出所有模板的便捷函数"""
    return template_manager.list_templates()


def save_template(template: TemplateConfig, format_type: str = "json"):
    """保存模板的便捷函数"""
    return template_manager.save_template(template, format_type)


def create_custom_template(name: str, description: str, base_template: str = "default") -> TemplateConfig:
    """创建自定义模板的便捷函数"""
    return template_manager.create_custom_template(name, description, base_template)