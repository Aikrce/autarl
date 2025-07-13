#!/usr/bin/env python3
"""
Humanities-Centered Template Merger
以人文社科模板为主体的智能模板融合工具
"""

import os
import json
import logging
from typing import Dict, List, Optional, Any
from pathlib import Path
import shutil
from datetime import datetime

from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.style import WD_STYLE_TYPE

from word_template_analyzer import TemplateLibrary, analyze_word_template

logger = logging.getLogger(__name__)


class HumanitiesCenteredMerger:
    """以人文社科为主体的模板融合器"""
    
    def __init__(self, template_library_path: str = "template_library"):
        self.template_library = TemplateLibrary(template_library_path)
        self.library_path = template_library_path
        
    def create_optimized_humanities_template(self,
                                           humanities_template_id: str,
                                           output_name: str = "人文社科优化版") -> str:
        """
        创建以人文社科为主体的优化模板
        
        Args:
            humanities_template_id: 人文社科模板ID
            output_name: 输出模板名称
            
        Returns:
            str: 新模板ID
        """
        try:
            # 获取人文社科模板信息
            templates = self.template_library.list_templates()
            if humanities_template_id not in templates:
                raise ValueError(f"人文社科模板不存在: {humanities_template_id}")
            
            humanities_template_data = templates[humanities_template_id]
            
            # 加载人文社科模板文档作为主体
            humanities_doc_path = humanities_template_data.get('word_file')
            if not humanities_doc_path or not os.path.exists(humanities_doc_path):
                raise ValueError(f"模板文件不存在: {humanities_doc_path}")
                
            base_doc = Document(humanities_doc_path)
            
            # 创建人文社科为主体的优化配置
            optimization_config = self._create_humanities_optimization_config(humanities_template_data)
            
            # 应用人文社科特色优化
            self._apply_humanities_optimizations(base_doc, optimization_config)
            
            # 增强学术论文特性
            self._enhance_academic_features(base_doc, optimization_config)
            
            # 生成新的模板ID
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            new_template_id = f"humanities_optimized_{timestamp}"
            
            # 创建新模板目录
            new_template_dir = os.path.join(self.library_path, new_template_id)
            os.makedirs(new_template_dir, exist_ok=True)
            
            # 保存优化模板
            optimized_template_path = os.path.join(new_template_dir, f"{new_template_id}.docx")
            base_doc.save(optimized_template_path)
            
            # 分析新模板
            optimized_template_info = analyze_word_template(optimized_template_path)
            
            # 保存模板配置
            config_path = os.path.join(new_template_dir, "template_config.json")
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(optimized_template_info.to_dict(), f, ensure_ascii=False, indent=2)
            
            # 创建内容结构配置
            structure_config = self._create_humanities_structure_config(optimization_config)
            structure_path = os.path.join(new_template_dir, "content_structure.json")
            with open(structure_path, 'w', encoding='utf-8') as f:
                json.dump(structure_config, f, ensure_ascii=False, indent=2)
            
            # 更新模板库索引
            self._update_template_index(new_template_id, output_name, optimized_template_path,
                                      config_path, structure_path, optimized_template_info)
            
            logger.info(f"人文社科优化模板创建成功: {new_template_id}")
            return new_template_id
            
        except Exception as e:
            logger.error(f"创建人文社科优化模板失败: {e}")
            raise
    
    def _create_humanities_optimization_config(self, humanities_template_data: Dict[str, Any]) -> Dict[str, Any]:
        """创建人文社科优化配置"""
        return {
            "name": "人文社科优化版",
            "description": "以人文社科为主体，融合东北师范大学格式规范的优化模板",
            "base_template": humanities_template_data.get('name', '人文社科'),
            "optimization_focus": "humanities",
            "features": {
                "humanities_citation_style": True,
                "enhanced_chinese_formatting": True,
                "academic_structure_support": True,
                "table_and_figure_optimization": True,
                "reference_management": True,
                "multilevel_headings": True,
                "footnote_support": True
            },
            "humanities_specific": {
                # 人文社科特色设置
                "citation_format": "chinese_academic",
                "bibliography_style": "humanities",
                "note_system": "footnotes_preferred",
                "paragraph_style": "chinese_academic",
                "heading_numbering": "chinese_traditional"
            },
            "style_optimizations": {
                "normal_paragraph": {
                    "first_line_indent": 0.75,  # 人文社科标准首行缩进
                    "line_spacing": 1.5,
                    "space_after": 0,
                    "font_family": "Times New Roman",
                    "font_size": 12,
                    "justification": True
                },
                "heading_1": {
                    "font_size": 16,
                    "alignment": "center",
                    "space_before": 24,
                    "space_after": 18,
                    "bold": True,
                    "numbering_style": "chinese"
                },
                "heading_2": {
                    "font_size": 14,
                    "alignment": "left", 
                    "space_before": 12,
                    "space_after": 6,
                    "bold": True,
                    "first_line_indent": 0.75
                },
                "heading_3": {
                    "font_size": 12,
                    "alignment": "left",
                    "space_before": 6,
                    "space_after": 3,
                    "bold": True,
                    "first_line_indent": 0.75
                },
                "footnote": {
                    "font_size": 9,
                    "line_spacing": 1.0,
                    "hanging_indent": 0.35
                },
                "bibliography": {
                    "font_size": 10.5,
                    "line_spacing": 1.5,
                    "hanging_indent": 0.75,
                    "space_after": 6
                },
                "table_caption": {
                    "font_size": 10.5,
                    "alignment": "center",
                    "space_before": 6,
                    "space_after": 6,
                    "numbering": "sequential"
                },
                "figure_caption": {
                    "font_size": 10.5,
                    "alignment": "center", 
                    "space_before": 6,
                    "space_after": 12,
                    "numbering": "sequential"
                }
            },
            "page_setup": {
                "margins": {
                    "top": 2.5,
                    "bottom": 2.0,
                    "left": 2.5,
                    "right": 2.0
                },
                "header_footer": {
                    "header_margin": 1.5,
                    "footer_margin": 1.25
                }
            }
        }
    
    def _apply_humanities_optimizations(self, doc: Document, config: Dict[str, Any]):
        """应用人文社科特色优化"""
        try:
            style_opts = config.get("style_optimizations", {})
            
            # 优化正文段落样式（人文社科核心）
            self._optimize_normal_paragraph_for_humanities(doc, style_opts.get("normal_paragraph", {}))
            
            # 优化标题样式系统
            self._optimize_heading_system_for_humanities(doc, style_opts)
            
            # 优化脚注样式
            self._optimize_footnote_style(doc, style_opts.get("footnote", {}))
            
            # 优化参考文献样式
            self._optimize_bibliography_style(doc, style_opts.get("bibliography", {}))
            
            # 应用人文社科页面设置
            self._apply_humanities_page_setup(doc, config.get("page_setup", {}))
            
            logger.info("人文社科特色优化应用完成")
            
        except Exception as e:
            logger.warning(f"人文社科优化应用失败: {e}")
    
    def _optimize_normal_paragraph_for_humanities(self, doc: Document, normal_config: Dict[str, Any]):
        """优化正文段落样式 - 人文社科特色"""
        try:
            if 'Normal' in doc.styles:
                normal_style = doc.styles['Normal']
                
                # 人文社科标准首行缩进
                if 'first_line_indent' in normal_config:
                    normal_style.paragraph_format.first_line_indent = Cm(normal_config['first_line_indent'])
                
                # 行间距
                if 'line_spacing' in normal_config:
                    normal_style.paragraph_format.line_spacing = normal_config['line_spacing']
                
                # 段后间距
                if 'space_after' in normal_config:
                    normal_style.paragraph_format.space_after = Pt(normal_config['space_after'])
                
                # 字体设置
                if 'font_family' in normal_config:
                    normal_style.font.name = normal_config['font_family']
                
                if 'font_size' in normal_config:
                    normal_style.font.size = Pt(normal_config['font_size'])
                
                # 两端对齐（人文社科标准）
                if normal_config.get('justification'):
                    normal_style.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
                    
        except Exception as e:
            logger.warning(f"优化正文段落样式失败: {e}")
    
    def _optimize_heading_system_for_humanities(self, doc: Document, style_opts: Dict[str, Any]):
        """优化标题系统 - 人文社科层级结构"""
        heading_configs = {
            'Heading 1': style_opts.get('heading_1', {}),
            'Heading 2': style_opts.get('heading_2', {}),
            'Heading 3': style_opts.get('heading_3', {})
        }
        
        for style_name, config in heading_configs.items():
            try:
                if style_name in doc.styles:
                    style = doc.styles[style_name]
                    
                    # 字体大小
                    if 'font_size' in config:
                        style.font.size = Pt(config['font_size'])
                    
                    # 对齐方式
                    if 'alignment' in config:
                        alignment_map = {
                            'center': WD_PARAGRAPH_ALIGNMENT.CENTER,
                            'left': WD_PARAGRAPH_ALIGNMENT.LEFT,
                            'justify': WD_PARAGRAPH_ALIGNMENT.JUSTIFY
                        }
                        style.paragraph_format.alignment = alignment_map.get(config['alignment'])
                    
                    # 段前段后间距
                    if 'space_before' in config:
                        style.paragraph_format.space_before = Pt(config['space_before'])
                    if 'space_after' in config:
                        style.paragraph_format.space_after = Pt(config['space_after'])
                    
                    # 加粗
                    if config.get('bold'):
                        style.font.bold = True
                    
                    # 首行缩进（二三级标题）
                    if 'first_line_indent' in config:
                        style.paragraph_format.first_line_indent = Cm(config['first_line_indent'])
                        
            except Exception as e:
                logger.warning(f"优化标题样式 {style_name} 失败: {e}")
    
    def _optimize_footnote_style(self, doc: Document, footnote_config: Dict[str, Any]):
        """优化脚注样式 - 人文社科重要特性"""
        try:
            footnote_styles = ['footnote text', '脚注文本']
            
            for style_name in footnote_styles:
                if style_name in doc.styles:
                    style = doc.styles[style_name]
                    
                    if 'font_size' in footnote_config:
                        style.font.size = Pt(footnote_config['font_size'])
                    
                    if 'line_spacing' in footnote_config:
                        style.paragraph_format.line_spacing = footnote_config['line_spacing']
                    
                    if 'hanging_indent' in footnote_config:
                        style.paragraph_format.hanging_indent = Cm(footnote_config['hanging_indent'])
                    
                    break
                    
        except Exception as e:
            logger.warning(f"优化脚注样式失败: {e}")
    
    def _optimize_bibliography_style(self, doc: Document, bib_config: Dict[str, Any]):
        """优化参考文献样式"""
        try:
            bib_styles = ['Bibliography', '参考文献', 'Reference']
            
            for style_name in bib_styles:
                if style_name in doc.styles:
                    style = doc.styles[style_name]
                    
                    if 'font_size' in bib_config:
                        style.font.size = Pt(bib_config['font_size'])
                    
                    if 'line_spacing' in bib_config:
                        style.paragraph_format.line_spacing = bib_config['line_spacing']
                    
                    if 'hanging_indent' in bib_config:
                        style.paragraph_format.hanging_indent = Cm(bib_config['hanging_indent'])
                    
                    if 'space_after' in bib_config:
                        style.paragraph_format.space_after = Pt(bib_config['space_after'])
                    
                    break
                    
        except Exception as e:
            logger.warning(f"优化参考文献样式失败: {e}")
    
    def _apply_humanities_page_setup(self, doc: Document, page_config: Dict[str, Any]):
        """应用人文社科页面设置"""
        try:
            margins = page_config.get('margins', {})
            
            for section in doc.sections:
                if 'top' in margins:
                    section.top_margin = Cm(margins['top'])
                if 'bottom' in margins:
                    section.bottom_margin = Cm(margins['bottom'])
                if 'left' in margins:
                    section.left_margin = Cm(margins['left'])
                if 'right' in margins:
                    section.right_margin = Cm(margins['right'])
                    
        except Exception as e:
            logger.warning(f"应用人文社科页面设置失败: {e}")
    
    def _enhance_academic_features(self, doc: Document, config: Dict[str, Any]):
        """增强学术论文特性"""
        try:
            # 优化表格和图片标题
            self._enhance_caption_styles(doc, config.get("style_optimizations", {}))
            
            # 确保学术编号系统
            self._ensure_academic_numbering(doc)
            
            logger.info("学术论文特性增强完成")
            
        except Exception as e:
            logger.warning(f"学术特性增强失败: {e}")
    
    def _enhance_caption_styles(self, doc: Document, style_opts: Dict[str, Any]):
        """增强标题样式"""
        caption_configs = {
            'Caption': style_opts.get('table_caption', {}),
            '表题': style_opts.get('table_caption', {}),
            '图题': style_opts.get('figure_caption', {})
        }
        
        for style_name, config in caption_configs.items():
            try:
                if style_name in doc.styles:
                    style = doc.styles[style_name]
                    
                    if 'font_size' in config:
                        style.font.size = Pt(config['font_size'])
                    
                    if 'alignment' in config:
                        alignment_map = {
                            'center': WD_PARAGRAPH_ALIGNMENT.CENTER,
                            'left': WD_PARAGRAPH_ALIGNMENT.LEFT
                        }
                        style.paragraph_format.alignment = alignment_map.get(config['alignment'])
                    
                    if 'space_before' in config:
                        style.paragraph_format.space_before = Pt(config['space_before'])
                    if 'space_after' in config:
                        style.paragraph_format.space_after = Pt(config['space_after'])
                        
            except Exception as e:
                logger.warning(f"增强标题样式 {style_name} 失败: {e}")
    
    def _ensure_academic_numbering(self, doc: Document):
        """确保学术编号系统"""
        # 这里可以添加编号格式的优化逻辑
        # 例如确保章节编号、图表编号等符合人文社科规范
        pass
    
    def _create_humanities_structure_config(self, optimization_config: Dict[str, Any]) -> Dict[str, Any]:
        """创建人文社科内容结构配置"""
        return {
            "document_structure": {
                "title_page": True,
                "abstract_cn": True,
                "abstract_en": True,
                "keywords": True,
                "table_of_contents": True,
                "main_content": True,
                "conclusion": True,
                "references": True,
                "appendix": True,
                "acknowledgments": True
            },
            "humanities_features": {
                "footnote_system": True,
                "bibliography_management": True,
                "citation_tracking": True,
                "multilevel_headings": True,
                "academic_formatting": True
            },
            "formatting_rules": {
                "heading_numbering": "chinese_humanities",
                "table_numbering": "sequential",
                "figure_numbering": "sequential", 
                "reference_style": "humanities_academic",
                "footnote_style": "chicago_humanities"
            },
            "optimization_info": {
                "created_at": datetime.now().isoformat(),
                "base_template": optimization_config.get("base_template"),
                "optimization_focus": optimization_config.get("optimization_focus"),
                "features": optimization_config.get("features", {}),
                "humanities_specific": optimization_config.get("humanities_specific", {})
            }
        }
    
    def _update_template_index(self, template_id: str, name: str, word_file: str,
                             config_file: str, structure_file: str, template_info) -> None:
        """更新模板库索引"""
        try:
            index_path = os.path.join(self.library_path, "template_index.json")
            
            # 读取现有索引
            if os.path.exists(index_path):
                with open(index_path, 'r', encoding='utf-8') as f:
                    index_data = json.load(f)
            else:
                index_data = {"templates": {}, "version": "1.0"}
            
            # 添加新的优化模板
            index_data["templates"][template_id] = {
                "name": name,
                "description": "以人文社科为主体的优化学术论文模板，融合东北师范大学格式规范",
                "tags": ["academic", "humanities", "thesis", "nenu", "optimized", "primary"],
                "created_at": datetime.now().isoformat(),
                "word_file": word_file,
                "config_file": config_file,
                "structure_file": structure_file,
                "styles_count": len(template_info.styles),
                "page_setup": {
                    "width": template_info.page_width,
                    "height": template_info.page_height,
                    "orientation": "portrait"
                },
                "is_optimized": True,
                "optimization_focus": "humanities",
                "template_priority": "primary"
            }
            
            # 保存索引
            with open(index_path, 'w', encoding='utf-8') as f:
                json.dump(index_data, f, ensure_ascii=False, indent=2)
                
        except Exception as e:
            logger.error(f"更新模板索引失败: {e}")
            raise


def create_humanities_optimized_template(template_library_path: str = "template_library") -> str:
    """创建人文社科优化模板的便捷函数"""
    merger = HumanitiesCenteredMerger(template_library_path)
    return merger.create_optimized_humanities_template("f93fa41a5664")


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    
    try:
        new_template_id = create_humanities_optimized_template()
        print(f"✅ 人文社科优化模板创建成功！模板ID: {new_template_id}")
    except Exception as e:
        print(f"❌ 创建人文社科优化模板失败: {e}")