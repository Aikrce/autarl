#!/usr/bin/env python3
"""
Markdown文档结构分析器
用于识别文档中的各种学术论文组件
"""

import re
import logging
from typing import Dict, List, Set, Optional
from dataclasses import dataclass

logger = logging.getLogger(__name__)

@dataclass
class DocumentSection:
    """文档章节信息"""
    name: str
    level: int
    start_line: int
    end_line: int
    content: str
    section_type: str  # 章节类型：abstract, toc, chapter, references, appendix等

class MarkdownDocumentAnalyzer:
    """Markdown文档结构分析器"""
    
    def __init__(self):
        # 定义各种学术论文组件的识别模式
        self.section_patterns = {
            'abstract_cn': [r'^#\s*摘\s*要', r'^#\s*中文摘要', r'^摘\s*要'],
            'abstract_en': [r'^#\s*Abstract', r'^#\s*ABSTRACT', r'^Abstract'],
            'keywords_cn': [r'关键词[：:]', r'关键字[：:]'],
            'keywords_en': [r'Key\s*words?[：:]', r'Keywords?[：:]'],
            'toc': [r'^#\s*目\s*录', r'^#\s*Table\s*of\s*Contents?', r'^目\s*录'],
            'introduction': [r'^#\s*引\s*言', r'^#\s*前\s*言', r'^#\s*绪\s*论', r'^#\s*Introduction'],
            'literature_review': [r'^#.*文献综述', r'^#.*Literature\s*Review', r'^#.*相关工作'],
            'methodology': [r'^#.*研究方法', r'^#.*方法', r'^#.*Methodology', r'^#.*Method'],
            'results': [r'^#.*结果', r'^#.*Results?', r'^#.*实验结果'],
            'discussion': [r'^#.*讨论', r'^#.*Discussion'],
            'conclusion': [r'^#\s*结\s*论', r'^#\s*总\s*结', r'^#\s*Conclusion'],
            'references': [r'^#\s*参考文献', r'^#\s*References?', r'^#\s*Bibliography'],
            'appendix': [r'^#\s*附\s*录', r'^#\s*Appendix', r'^附\s*录'],
            'acknowledgments': [r'^#\s*致\s*谢', r'^#\s*Acknowledgments?', r'^致\s*谢'],
            'symbols': [r'^#.*符号.*说明', r'^#.*缩略语.*说明', r'^#.*Symbols?'],
            'figures_list': [r'^#.*插图目录', r'^#.*图.*目录', r'^#.*List\s*of\s*Figures?'],
            'tables_list': [r'^#.*附表目录', r'^#.*表.*目录', r'^#.*List\s*of\s*Tables?']
        }
        
        # 章节编号模式
        self.chapter_patterns = [
            r'^#\s*第[一二三四五六七八九十\d]+章',  # 第一章
            r'^#\s*Chapter\s*\d+',  # Chapter 1
            r'^#\s*\d+[\.\s]',  # 1. 或 1 
        ]
    
    def analyze_document(self, content: str) -> Dict:
        """分析文档结构并返回分析结果"""
        lines = content.split('\n')
        
        analysis_result = {
            'sections': [],
            'detected_components': set(),
            'document_type': 'unknown',
            'has_academic_structure': False,
            'content_mapping': {}
        }
        
        # 分析每一行
        current_section = None
        sections = []
        
        for i, line in enumerate(lines):
            line_stripped = line.strip()
            
            # 检查是否是标题行
            if line_stripped.startswith('#'):
                # 如果有当前章节，先保存
                if current_section:
                    current_section.end_line = i - 1
                    current_section.content = '\n'.join(lines[current_section.start_line:i])
                    sections.append(current_section)
                
                # 创建新章节
                level = len(line_stripped) - len(line_stripped.lstrip('#'))
                section_name = line_stripped.lstrip('#').strip()
                section_type = self._identify_section_type(section_name)
                
                current_section = DocumentSection(
                    name=section_name,
                    level=level,
                    start_line=i,
                    end_line=i,
                    content=line,
                    section_type=section_type
                )
                
                # 记录检测到的组件
                if section_type != 'unknown':
                    analysis_result['detected_components'].add(section_type)
        
        # 保存最后一个章节
        if current_section:
            current_section.end_line = len(lines) - 1
            current_section.content = '\n'.join(lines[current_section.start_line:])
            sections.append(current_section)
        
        analysis_result['sections'] = sections
        
        # 判断文档类型
        analysis_result['document_type'] = self._determine_document_type(analysis_result['detected_components'])
        analysis_result['has_academic_structure'] = self._has_academic_structure(analysis_result['detected_components'])
        
        # 创建内容映射
        analysis_result['content_mapping'] = self._create_content_mapping(sections)
        
        logger.info(f"文档分析完成: 检测到 {len(analysis_result['detected_components'])} 个学术组件")
        logger.info(f"文档类型: {analysis_result['document_type']}")
        
        return analysis_result
    
    def _identify_section_type(self, section_name: str) -> str:
        """识别章节类型"""
        section_name_lower = section_name.lower()
        
        for section_type, patterns in self.section_patterns.items():
            for pattern in patterns:
                if re.search(pattern, section_name, re.IGNORECASE):
                    return section_type
        
        # 检查是否是章节
        for pattern in self.chapter_patterns:
            if re.search(pattern, f"# {section_name}", re.IGNORECASE):
                return 'chapter'
        
        return 'unknown'
    
    def _determine_document_type(self, components: Set[str]) -> str:
        """根据检测到的组件判断文档类型"""
        academic_components = {'abstract_cn', 'abstract_en', 'references', 'introduction', 'conclusion'}
        business_components = {'executive_summary', 'recommendations', 'analysis'}
        
        if len(components.intersection(academic_components)) >= 2:
            return 'academic_thesis'
        elif 'abstract_cn' in components or 'abstract_en' in components:
            return 'academic_paper'
        elif len(components.intersection(business_components)) >= 1:
            return 'business_report'
        elif 'chapter' in [comp.split('_')[0] for comp in components]:
            return 'structured_document'
        else:
            return 'general_document'
    
    def _has_academic_structure(self, components: Set[str]) -> bool:
        """判断是否具有学术论文结构"""
        required_components = {'references'}  # 至少需要参考文献
        optional_components = {'abstract_cn', 'abstract_en', 'introduction', 'conclusion'}
        
        has_required = len(components.intersection(required_components)) > 0
        has_optional = len(components.intersection(optional_components)) >= 1
        
        return has_required or has_optional
    
    def _create_content_mapping(self, sections: List[DocumentSection]) -> Dict:
        """创建内容映射，将检测到的内容与模板组件对应"""
        mapping = {}
        
        for section in sections:
            if section.section_type != 'unknown':
                mapping[section.section_type] = {
                    'name': section.name,
                    'content': section.content,
                    'level': section.level,
                    'start_line': section.start_line,
                    'end_line': section.end_line
                }
        
        return mapping
    
    def get_missing_components(self, template_name: str, detected_components: Set[str]) -> Set[str]:
        """获取模板中缺失的组件"""
        # 定义各模板的标准组件
        template_components = {
            'nenu_thesis': {
                'abstract_cn', 'abstract_en', 'keywords_cn', 'keywords_en',
                'toc', 'symbols', 'figures_list', 'tables_list',
                'introduction', 'literature_review', 'methodology', 
                'results', 'discussion', 'conclusion', 
                'references', 'appendix', 'acknowledgments'
            },
            'academic_paper': {
                'abstract_cn', 'abstract_en', 'keywords_cn', 'keywords_en',
                'introduction', 'methodology', 'results', 'conclusion', 'references'
            },
            'business_report': {
                'executive_summary', 'introduction', 'analysis', 
                'recommendations', 'conclusion', 'references'
            },
            'default': set()
        }
        
        expected_components = template_components.get(template_name, set())
        missing_components = expected_components - detected_components
        
        return missing_components

def analyze_markdown_document(content: str) -> Dict:
    """便捷函数：分析Markdown文档"""
    analyzer = MarkdownDocumentAnalyzer()
    return analyzer.analyze_document(content)