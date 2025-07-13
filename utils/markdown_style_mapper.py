#!/usr/bin/env python3
"""
Markdown to Word Style Mapping System
Markdown到Word智能样式映射系统 - 基于分析的Word模板实现精确样式映射
"""

import re
import logging
from typing import Dict, List, Optional, Any, Tuple, Union
from dataclasses import dataclass, field
from enum import Enum
import json
from pathlib import Path

from word_template_analyzer import WordDocumentInfo, WordStyleInfo, TemplateLibrary
from enhanced_document_analyzer import ContentSection, SectionType, DocumentType

logger = logging.getLogger(__name__)


class MarkdownElementType(Enum):
    """Markdown元素类型"""
    HEADING_1 = "heading_1"
    HEADING_2 = "heading_2"
    HEADING_3 = "heading_3"
    HEADING_4 = "heading_4"
    HEADING_5 = "heading_5"
    HEADING_6 = "heading_6"
    PARAGRAPH = "paragraph"
    BOLD = "bold"
    ITALIC = "italic"
    CODE_INLINE = "code_inline"
    CODE_BLOCK = "code_block"
    QUOTE = "quote"
    LIST_UNORDERED = "list_unordered"
    LIST_ORDERED = "list_ordered"
    LIST_ITEM = "list_item"
    TABLE = "table"
    TABLE_HEADER = "table_header"
    TABLE_CELL = "table_cell"
    LINK = "link"
    IMAGE = "image"
    HORIZONTAL_RULE = "horizontal_rule"
    
    # 学术文档特殊元素
    ABSTRACT_TITLE = "abstract_title"
    ABSTRACT_CONTENT = "abstract_content"
    KEYWORDS = "keywords"
    CHAPTER_TITLE = "chapter_title"
    SECTION_TITLE = "section_title"
    REFERENCE_TITLE = "reference_title"
    REFERENCE_ITEM = "reference_item"
    TOC_TITLE = "toc_title"
    TOC_ITEM = "toc_item"


@dataclass
class StyleMapping:
    """样式映射规则"""
    markdown_element: MarkdownElementType
    word_style_id: str
    word_style_name: str
    priority: int = 1  # 优先级，数字越大优先级越高
    conditions: List[str] = field(default_factory=list)  # 应用条件
    fallback_style: Optional[str] = None  # 备用样式
    
    def to_dict(self) -> Dict[str, Any]:
        return {
            'markdown_element': self.markdown_element.value,
            'word_style_id': self.word_style_id,
            'word_style_name': self.word_style_name,
            'priority': self.priority,
            'conditions': self.conditions,
            'fallback_style': self.fallback_style
        }


@dataclass
class ContextualRule:
    """上下文规则"""
    rule_name: str
    condition: str  # 条件表达式
    markdown_pattern: str  # 匹配的Markdown模式
    target_style: str  # 目标样式
    context_requirements: List[str] = field(default_factory=list)  # 上下文要求
    
    def to_dict(self) -> Dict[str, Any]:
        return {
            'rule_name': self.rule_name,
            'condition': self.condition,
            'markdown_pattern': self.markdown_pattern,
            'target_style': self.target_style,
            'context_requirements': self.context_requirements
        }


class MarkdownWordStyleMapper:
    """Markdown到Word样式映射器"""
    
    def __init__(self, word_template_info: WordDocumentInfo):
        self.template_info = word_template_info
        self.style_mappings: List[StyleMapping] = []
        self.contextual_rules: List[ContextualRule] = []
        
        # 构建样式索引
        self.word_styles_by_id = {style.style_id: style for style in word_template_info.styles}
        self.word_styles_by_name = {style.name: style for style in word_template_info.styles}
        
        # 初始化智能映射
        self._initialize_smart_mappings()
        
        logger.info(f"样式映射器初始化完成，模板: {word_template_info.filename}")
    
    def _initialize_smart_mappings(self):
        """初始化智能映射规则"""
        # 基础标题映射
        self._map_headings()
        
        # 段落和文本格式映射
        self._map_text_formatting()
        
        # 列表映射
        self._map_lists()
        
        # 代码和引用映射
        self._map_code_and_quotes()
        
        # 学术文档特殊映射
        self._map_academic_elements()
        
        # 表格映射
        self._map_tables()
        
        # 设置上下文规则
        self._setup_contextual_rules()
    
    def _map_headings(self):
        """映射标题样式"""
        # 查找标题样式
        heading_styles = self._find_heading_styles()
        
        # 映射各级标题
        for level in range(1, 7):
            markdown_element = getattr(MarkdownElementType, f"HEADING_{level}")
            
            # 查找对应的Word样式
            word_style = self._find_best_heading_style(level, heading_styles)
            
            if word_style:
                mapping = StyleMapping(
                    markdown_element=markdown_element,
                    word_style_id=word_style.style_id,
                    word_style_name=word_style.name,
                    priority=10 - level  # 高级标题优先级更高
                )
                self.style_mappings.append(mapping)
                logger.debug(f"映射 H{level} -> {word_style.name}")
    
    def _find_heading_styles(self) -> List[WordStyleInfo]:
        """查找所有标题样式"""
        heading_styles = []
        
        for style in self.template_info.styles:
            style_name_lower = style.name.lower()
            
            # 常见的标题样式名称模式
            heading_patterns = [
                r'heading\s*\d+',
                r'标题\s*\d*',
                r'title',
                r'章.*标题',
                r'节.*标题',
                r'h\d+',
                r'subtitle'
            ]
            
            for pattern in heading_patterns:
                if re.search(pattern, style_name_lower):
                    heading_styles.append(style)
                    break
        
        return heading_styles
    
    def _find_best_heading_style(self, level: int, heading_styles: List[WordStyleInfo]) -> Optional[WordStyleInfo]:
        """为指定级别查找最佳标题样式"""
        # 首先查找明确包含级别数字的样式
        for style in heading_styles:
            if f"heading {level}" in style.name.lower() or f"heading{level}" in style.name.lower():
                return style
            if f"标题 {level}" in style.name or f"标题{level}" in style.name:
                return style
        
        # 查找可能的匹配
        level_keywords = {
            1: ['title', '标题', '章', 'chapter'],
            2: ['subtitle', '副标题', '节', 'section'],
            3: ['subheading', '小标题', '条', 'subsection'],
            4: ['minor', '次标题'],
            5: ['small', '小'],
            6: ['smallest', '最小']
        }
        
        keywords = level_keywords.get(level, [])
        for keyword in keywords:
            for style in heading_styles:
                if keyword in style.name.lower():
                    return style
        
        # 如果没有找到，返回第一个标题样式
        return heading_styles[0] if heading_styles else None
    
    def _map_text_formatting(self):
        """映射文本格式"""
        # 正文段落
        normal_style = self._find_style_by_patterns(['normal', '正文', 'body', 'paragraph'])
        if normal_style:
            mapping = StyleMapping(
                markdown_element=MarkdownElementType.PARAGRAPH,
                word_style_id=normal_style.style_id,
                word_style_name=normal_style.name,
                priority=1
            )
            self.style_mappings.append(mapping)
        
        # 粗体（通常不需要特殊样式，使用字符格式）
        # 斜体（同样使用字符格式）
        # 这些会在转换时直接应用字符格式
    
    def _map_lists(self):
        """映射列表样式"""
        # 查找列表样式
        list_styles = self._find_list_styles()
        
        if list_styles.get('unordered'):
            mapping = StyleMapping(
                markdown_element=MarkdownElementType.LIST_UNORDERED,
                word_style_id=list_styles['unordered'].style_id,
                word_style_name=list_styles['unordered'].name,
                priority=3
            )
            self.style_mappings.append(mapping)
        
        if list_styles.get('ordered'):
            mapping = StyleMapping(
                markdown_element=MarkdownElementType.LIST_ORDERED,
                word_style_id=list_styles['ordered'].style_id,
                word_style_name=list_styles['ordered'].name,
                priority=3
            )
            self.style_mappings.append(mapping)
    
    def _find_list_styles(self) -> Dict[str, WordStyleInfo]:
        """查找列表样式"""
        list_styles = {}
        
        for style in self.template_info.styles:
            style_name_lower = style.name.lower()
            
            # 无序列表
            if any(pattern in style_name_lower for pattern in ['list bullet', '项目符号', 'bullet']):
                list_styles['unordered'] = style
            
            # 有序列表
            elif any(pattern in style_name_lower for pattern in ['list number', '编号', 'numbered']):
                list_styles['ordered'] = style
        
        return list_styles
    
    def _map_code_and_quotes(self):
        """映射代码和引用样式"""
        # 代码块
        code_style = self._find_style_by_patterns(['code', '代码', 'monospace', 'pre'])
        if code_style:
            mapping = StyleMapping(
                markdown_element=MarkdownElementType.CODE_BLOCK,
                word_style_id=code_style.style_id,
                word_style_name=code_style.name,
                priority=5
            )
            self.style_mappings.append(mapping)
        
        # 引用
        quote_style = self._find_style_by_patterns(['quote', '引用', 'blockquote', 'citation'])
        if quote_style:
            mapping = StyleMapping(
                markdown_element=MarkdownElementType.QUOTE,
                word_style_id=quote_style.style_id,
                word_style_name=quote_style.name,
                priority=4
            )
            self.style_mappings.append(mapping)
    
    def _map_academic_elements(self):
        """映射学术文档特殊元素"""
        # 摘要标题
        abstract_title_style = self._find_style_by_patterns([
            'abstract title', '摘要标题', 'abstract heading'
        ])
        if abstract_title_style:
            mapping = StyleMapping(
                markdown_element=MarkdownElementType.ABSTRACT_TITLE,
                word_style_id=abstract_title_style.style_id,
                word_style_name=abstract_title_style.name,
                priority=8
            )
            self.style_mappings.append(mapping)
        
        # 摘要内容
        abstract_content_style = self._find_style_by_patterns([
            'abstract body', '摘要正文', 'abstract content', 'abstract'
        ])
        if abstract_content_style:
            mapping = StyleMapping(
                markdown_element=MarkdownElementType.ABSTRACT_CONTENT,
                word_style_id=abstract_content_style.style_id,
                word_style_name=abstract_content_style.name,
                priority=7
            )
            self.style_mappings.append(mapping)
        
        # 关键词
        keywords_style = self._find_style_by_patterns(['keywords', '关键词', 'key words'])
        if keywords_style:
            mapping = StyleMapping(
                markdown_element=MarkdownElementType.KEYWORDS,
                word_style_id=keywords_style.style_id,
                word_style_name=keywords_style.name,
                priority=6
            )
            self.style_mappings.append(mapping)
        
        # 章标题
        chapter_style = self._find_style_by_patterns([
            'chapter title', '章标题', 'chapter heading', 'chapter'
        ])
        if chapter_style:
            mapping = StyleMapping(
                markdown_element=MarkdownElementType.CHAPTER_TITLE,
                word_style_id=chapter_style.style_id,
                word_style_name=chapter_style.name,
                priority=9
            )
            self.style_mappings.append(mapping)
        
        # 参考文献标题
        ref_title_style = self._find_style_by_patterns([
            'reference title', '参考文献标题', 'bibliography title'
        ])
        if ref_title_style:
            mapping = StyleMapping(
                markdown_element=MarkdownElementType.REFERENCE_TITLE,
                word_style_id=ref_title_style.style_id,
                word_style_name=ref_title_style.name,
                priority=8
            )
            self.style_mappings.append(mapping)
        
        # 参考文献条目
        ref_item_style = self._find_style_by_patterns([
            'reference content', '参考文献', 'bibliography', 'reference'
        ])
        if ref_item_style:
            mapping = StyleMapping(
                markdown_element=MarkdownElementType.REFERENCE_ITEM,
                word_style_id=ref_item_style.style_id,
                word_style_name=ref_item_style.name,
                priority=6
            )
            self.style_mappings.append(mapping)
        
        # 目录标题
        toc_title_style = self._find_style_by_patterns(['toc title', '目录标题', 'contents title'])
        if toc_title_style:
            mapping = StyleMapping(
                markdown_element=MarkdownElementType.TOC_TITLE,
                word_style_id=toc_title_style.style_id,
                word_style_name=toc_title_style.name,
                priority=8
            )
            self.style_mappings.append(mapping)
    
    def _map_tables(self):
        """映射表格样式"""
        # 查找表格样式
        table_styles = self._find_table_styles()
        
        if table_styles.get('normal'):
            mapping = StyleMapping(
                markdown_element=MarkdownElementType.TABLE,
                word_style_id=table_styles['normal'].style_id,
                word_style_name=table_styles['normal'].name,
                priority=4
            )
            self.style_mappings.append(mapping)
    
    def _find_table_styles(self) -> Dict[str, WordStyleInfo]:
        """查找表格样式"""
        table_styles = {}
        
        for style in self.template_info.styles:
            if style.style_type == 'table':
                style_name_lower = style.name.lower()
                
                # 普通表格
                if any(pattern in style_name_lower for pattern in ['table grid', 'normal table', '普通表格']):
                    table_styles['normal'] = style
                # 表格标题
                elif any(pattern in style_name_lower for pattern in ['table heading', '表格标题']):
                    table_styles['header'] = style
        
        return table_styles
    
    def _find_style_by_patterns(self, patterns: List[str]) -> Optional[WordStyleInfo]:
        """根据模式查找样式"""
        for pattern in patterns:
            pattern_lower = pattern.lower()
            
            # 精确匹配
            for style in self.template_info.styles:
                if pattern_lower == style.name.lower():
                    return style
            
            # 包含匹配
            for style in self.template_info.styles:
                if pattern_lower in style.name.lower():
                    return style
        
        return None
    
    def _setup_contextual_rules(self):
        """设置上下文规则"""
        # 中文摘要识别规则
        cn_abstract_rule = ContextualRule(
            rule_name="Chinese Abstract Detection",
            condition="contains_chinese_abstract",
            markdown_pattern=r"#+\s*摘\s*要|#+\s*摘　　要",
            target_style="Abstract Title CN",
            context_requirements=["has_chinese_keywords"]
        )
        self.contextual_rules.append(cn_abstract_rule)
        
        # 英文摘要识别规则
        en_abstract_rule = ContextualRule(
            rule_name="English Abstract Detection",
            condition="contains_english_abstract",
            markdown_pattern=r"#+\s*Abstract",
            target_style="Abstract Title EN",
            context_requirements=["has_english_keywords"]
        )
        self.contextual_rules.append(en_abstract_rule)
        
        # 章节标题识别规则
        chapter_rule = ContextualRule(
            rule_name="Chapter Title Detection",
            condition="is_chapter_title",
            markdown_pattern=r"#+\s*第[一二三四五六七八九十\d]+章",
            target_style="Chapter Title",
            context_requirements=[]
        )
        self.contextual_rules.append(chapter_rule)
    
    def get_style_mapping(self, markdown_element: MarkdownElementType, 
                         context: Optional[Dict[str, Any]] = None) -> Optional[StyleMapping]:
        """
        获取Markdown元素的样式映射
        
        Args:
            markdown_element: Markdown元素类型
            context: 上下文信息
            
        Returns:
            StyleMapping: 样式映射或None
        """
        # 查找直接映射
        direct_mappings = [m for m in self.style_mappings if m.markdown_element == markdown_element]
        
        if not direct_mappings:
            return None
        
        # 如果有多个映射，选择优先级最高的
        best_mapping = max(direct_mappings, key=lambda m: m.priority)
        
        # 检查条件
        if best_mapping.conditions and context:
            if not self._check_conditions(best_mapping.conditions, context):
                # 尝试备用样式
                if best_mapping.fallback_style:
                    fallback_style = self.word_styles_by_name.get(best_mapping.fallback_style)
                    if fallback_style:
                        return StyleMapping(
                            markdown_element=markdown_element,
                            word_style_id=fallback_style.style_id,
                            word_style_name=fallback_style.name,
                            priority=0
                        )
                return None
        
        return best_mapping
    
    def get_contextual_style(self, text: str, markdown_element: MarkdownElementType,
                           document_context: Dict[str, Any]) -> Optional[StyleMapping]:
        """
        根据上下文获取样式映射
        
        Args:
            text: 文本内容
            markdown_element: Markdown元素类型
            document_context: 文档上下文
            
        Returns:
            StyleMapping: 样式映射或None
        """
        for rule in self.contextual_rules:
            if re.search(rule.markdown_pattern, text, re.IGNORECASE):
                # 检查上下文要求
                if self._check_context_requirements(rule.context_requirements, document_context):
                    # 查找目标样式
                    target_style = self.word_styles_by_name.get(rule.target_style)
                    if target_style:
                        return StyleMapping(
                            markdown_element=markdown_element,
                            word_style_id=target_style.style_id,
                            word_style_name=target_style.name,
                            priority=10  # 上下文规则优先级高
                        )
        
        # 没有匹配的上下文规则，使用普通映射
        return self.get_style_mapping(markdown_element)
    
    def _check_conditions(self, conditions: List[str], context: Dict[str, Any]) -> bool:
        """检查条件是否满足"""
        for condition in conditions:
            if condition not in context or not context[condition]:
                return False
        return True
    
    def _check_context_requirements(self, requirements: List[str], 
                                  document_context: Dict[str, Any]) -> bool:
        """检查上下文要求"""
        for requirement in requirements:
            if requirement == "has_chinese_keywords":
                if not document_context.get('has_chinese_keywords', False):
                    return False
            elif requirement == "has_english_keywords":
                if not document_context.get('has_english_keywords', False):
                    return False
            # 可以添加更多上下文要求
        
        return True
    
    def add_custom_mapping(self, markdown_element: MarkdownElementType,
                          word_style_name: str, priority: int = 5,
                          conditions: List[str] = None) -> bool:
        """
        添加自定义映射
        
        Args:
            markdown_element: Markdown元素类型
            word_style_name: Word样式名称
            priority: 优先级
            conditions: 条件列表
            
        Returns:
            bool: 是否添加成功
        """
        word_style = self.word_styles_by_name.get(word_style_name)
        if not word_style:
            logger.warning(f"Word样式不存在: {word_style_name}")
            return False
        
        mapping = StyleMapping(
            markdown_element=markdown_element,
            word_style_id=word_style.style_id,
            word_style_name=word_style.name,
            priority=priority,
            conditions=conditions or []
        )
        
        self.style_mappings.append(mapping)
        logger.info(f"已添加自定义映射: {markdown_element.value} -> {word_style_name}")
        return True
    
    def remove_mapping(self, markdown_element: MarkdownElementType, 
                      word_style_name: str) -> bool:
        """删除映射"""
        original_count = len(self.style_mappings)
        
        self.style_mappings = [
            m for m in self.style_mappings 
            if not (m.markdown_element == markdown_element and m.word_style_name == word_style_name)
        ]
        
        removed_count = original_count - len(self.style_mappings)
        if removed_count > 0:
            logger.info(f"已删除 {removed_count} 个映射")
            return True
        
        return False
    
    def export_mappings(self, output_path: str):
        """导出映射配置"""
        try:
            export_data = {
                'template_info': {
                    'filename': self.template_info.filename,
                    'styles_count': len(self.template_info.styles)
                },
                'style_mappings': [mapping.to_dict() for mapping in self.style_mappings],
                'contextual_rules': [rule.to_dict() for rule in self.contextual_rules]
            }
            
            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(export_data, f, ensure_ascii=False, indent=2)
            
            logger.info(f"映射配置已导出到: {output_path}")
            
        except Exception as e:
            logger.error(f"导出映射配置失败: {e}")
    
    def import_mappings(self, config_path: str):
        """导入映射配置"""
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                config_data = json.load(f)
            
            # 清空现有映射
            self.style_mappings.clear()
            self.contextual_rules.clear()
            
            # 导入样式映射
            for mapping_data in config_data.get('style_mappings', []):
                markdown_element = MarkdownElementType(mapping_data['markdown_element'])
                mapping = StyleMapping(
                    markdown_element=markdown_element,
                    word_style_id=mapping_data['word_style_id'],
                    word_style_name=mapping_data['word_style_name'],
                    priority=mapping_data['priority'],
                    conditions=mapping_data.get('conditions', []),
                    fallback_style=mapping_data.get('fallback_style')
                )
                self.style_mappings.append(mapping)
            
            # 导入上下文规则
            for rule_data in config_data.get('contextual_rules', []):
                rule = ContextualRule(
                    rule_name=rule_data['rule_name'],
                    condition=rule_data['condition'],
                    markdown_pattern=rule_data['markdown_pattern'],
                    target_style=rule_data['target_style'],
                    context_requirements=rule_data.get('context_requirements', [])
                )
                self.contextual_rules.append(rule)
            
            logger.info(f"映射配置已导入: {len(self.style_mappings)} 个映射, {len(self.contextual_rules)} 个规则")
            
        except Exception as e:
            logger.error(f"导入映射配置失败: {e}")
    
    def get_mapping_statistics(self) -> Dict[str, Any]:
        """获取映射统计信息"""
        element_counts = {}
        for mapping in self.style_mappings:
            element_type = mapping.markdown_element.value
            if element_type not in element_counts:
                element_counts[element_type] = 0
            element_counts[element_type] += 1
        
        return {
            'total_mappings': len(self.style_mappings),
            'total_rules': len(self.contextual_rules),
            'element_distribution': element_counts,
            'template_styles_count': len(self.template_info.styles),
            'mapped_styles_count': len(set(m.word_style_name for m in self.style_mappings)),
            'unmapped_styles': [
                style.name for style in self.template_info.styles
                if style.name not in [m.word_style_name for m in self.style_mappings]
            ]
        }
    
    def suggest_improvements(self) -> List[str]:
        """建议映射改进"""
        suggestions = []
        stats = self.get_mapping_statistics()
        
        # 检查未映射的样式
        if stats['unmapped_styles']:
            suggestions.append(f"发现 {len(stats['unmapped_styles'])} 个未映射的Word样式，可能需要添加映射")
        
        # 检查缺失的基本元素映射
        basic_elements = [
            MarkdownElementType.HEADING_1,
            MarkdownElementType.HEADING_2,
            MarkdownElementType.PARAGRAPH,
            MarkdownElementType.LIST_UNORDERED,
            MarkdownElementType.LIST_ORDERED
        ]
        
        mapped_elements = set(m.markdown_element for m in self.style_mappings)
        missing_elements = [elem for elem in basic_elements if elem not in mapped_elements]
        
        if missing_elements:
            missing_names = [elem.value for elem in missing_elements]
            suggestions.append(f"缺少基本元素映射: {', '.join(missing_names)}")
        
        # 检查优先级冲突
        priority_conflicts = self._find_priority_conflicts()
        if priority_conflicts:
            suggestions.append(f"发现 {len(priority_conflicts)} 个优先级冲突")
        
        return suggestions
    
    def _find_priority_conflicts(self) -> List[str]:
        """查找优先级冲突"""
        conflicts = []
        element_priorities = {}
        
        for mapping in self.style_mappings:
            element = mapping.markdown_element
            if element not in element_priorities:
                element_priorities[element] = []
            element_priorities[element].append((mapping.word_style_name, mapping.priority))
        
        for element, priorities in element_priorities.items():
            if len(priorities) > 1:
                # 检查是否有相同优先级
                priority_values = [p[1] for p in priorities]
                if len(set(priority_values)) < len(priority_values):
                    conflicts.append(f"{element.value}: 多个样式具有相同优先级")
        
        return conflicts


class MarkdownAnalyzer:
    """Markdown内容分析器，用于辅助样式映射"""
    
    def __init__(self):
        self.patterns = {
            'chinese_abstract': r'#+\s*摘\s*要|#+\s*摘　　要',
            'english_abstract': r'#+\s*Abstract',
            'chinese_keywords': r'关键词[:：]\s*',
            'english_keywords': r'Key\s*words?[:：]\s*',
            'chapter_title': r'#+\s*第[一二三四五六七八九十\d]+章',
            'section_title': r'#+\s*\d+\.\d+',
            'references': r'#+\s*(参考文献|References|Bibliography)',
            'toc': r'#+\s*(目录|目　　录|Table\s*of\s*Contents)'
        }
    
    def analyze_document_context(self, content: str) -> Dict[str, Any]:
        """分析文档上下文"""
        context = {
            'has_chinese_abstract': bool(re.search(self.patterns['chinese_abstract'], content, re.IGNORECASE)),
            'has_english_abstract': bool(re.search(self.patterns['english_abstract'], content, re.IGNORECASE)),
            'has_chinese_keywords': bool(re.search(self.patterns['chinese_keywords'], content, re.IGNORECASE)),
            'has_english_keywords': bool(re.search(self.patterns['english_keywords'], content, re.IGNORECASE)),
            'has_chapters': bool(re.search(self.patterns['chapter_title'], content, re.IGNORECASE)),
            'has_sections': bool(re.search(self.patterns['section_title'], content, re.IGNORECASE)),
            'has_references': bool(re.search(self.patterns['references'], content, re.IGNORECASE)),
            'has_toc': bool(re.search(self.patterns['toc'], content, re.IGNORECASE)),
            'document_language': self._detect_language(content),
            'estimated_type': self._estimate_document_type(content)
        }
        
        return context
    
    def _detect_language(self, content: str) -> str:
        """检测文档主要语言"""
        # 简单的语言检测
        chinese_chars = len(re.findall(r'[\u4e00-\u9fff]', content))
        english_chars = len(re.findall(r'[a-zA-Z]', content))
        
        if chinese_chars > english_chars:
            return 'chinese'
        elif english_chars > chinese_chars:
            return 'english'
        else:
            return 'mixed'
    
    def _estimate_document_type(self, content: str) -> str:
        """估计文档类型"""
        # 检查学术文档特征
        academic_indicators = [
            self.patterns['chinese_abstract'],
            self.patterns['english_abstract'],
            self.patterns['references'],
            r'doi[:：]\s*\d+',
            r'\[[^\]]*\d{4}[^\]]*\]'  # 年份引用
        ]
        
        academic_score = sum(1 for pattern in academic_indicators if re.search(pattern, content, re.IGNORECASE))
        
        if academic_score >= 2:
            return 'academic'
        elif re.search(r'API|SDK|函数|方法|class|function', content, re.IGNORECASE):
            return 'technical'
        elif re.search(r'报告|分析|总结|商业|市场', content, re.IGNORECASE):
            return 'business'
        else:
            return 'general'


# 便捷函数
def create_style_mapper(word_template_path: str) -> MarkdownWordStyleMapper:
    """创建样式映射器的便捷函数"""
    from word_template_analyzer import analyze_word_template
    
    template_info = analyze_word_template(word_template_path)
    return MarkdownWordStyleMapper(template_info)


def analyze_markdown_for_mapping(content: str) -> Dict[str, Any]:
    """分析Markdown内容用于样式映射的便捷函数"""
    analyzer = MarkdownAnalyzer()
    return analyzer.analyze_document_context(content)