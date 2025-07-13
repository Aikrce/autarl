#!/usr/bin/env python3
"""
Enhanced Document Analyzer
增强的文档分析器，提供更智能的内容识别和分类能力
"""

import re
import logging
from typing import Dict, List, Set, Optional, Tuple, Any
from dataclasses import dataclass, field
from enum import Enum
import hashlib


logger = logging.getLogger(__name__)


class DocumentType(Enum):
    """文档类型枚举"""
    ACADEMIC_THESIS = "academic_thesis"
    RESEARCH_PAPER = "research_paper"
    TECHNICAL_REPORT = "technical_report"
    BUSINESS_REPORT = "business_report"
    MANUAL = "manual"
    GENERAL_DOCUMENT = "general_document"


class SectionType(Enum):
    """章节类型枚举"""
    COVER_PAGE = "cover_page"
    ENGLISH_COVER = "english_cover"
    DECLARATION = "declaration"
    AUTHORIZATION = "authorization"
    ABSTRACT_CN = "abstract_cn"
    ABSTRACT_EN = "abstract_en"
    KEYWORDS_CN = "keywords_cn"
    KEYWORDS_EN = "keywords_en"
    TOC = "toc"
    SYMBOLS = "symbols"
    FIGURES_LIST = "figures_list"
    TABLES_LIST = "tables_list"
    INTRODUCTION = "introduction"
    LITERATURE_REVIEW = "literature_review"
    METHODOLOGY = "methodology"
    RESULTS = "results"
    DISCUSSION = "discussion"
    CONCLUSION = "conclusion"
    REFERENCES = "references"
    APPENDIX = "appendix"
    ACKNOWLEDGMENTS = "acknowledgments"
    CHAPTER = "chapter"
    SECTION = "section"
    SUBSECTION = "subsection"
    UNKNOWN = "unknown"


@dataclass
class ContentSection:
    """内容章节类"""
    name: str
    content: str
    section_type: SectionType
    level: int = 1
    start_line: int = 0
    end_line: int = 0
    confidence: float = 1.0
    metadata: Dict[str, Any] = field(default_factory=dict)
    
    def __post_init__(self):
        self.content_hash = hashlib.md5(self.content.encode()).hexdigest()


@dataclass
class DocumentStructure:
    """文档结构分析结果"""
    document_type: DocumentType
    sections: List[ContentSection]
    detected_components: Set[str]
    content_mapping: Dict[str, str]
    statistics: Dict[str, Any]
    confidence_score: float = 0.0
    
    def __post_init__(self):
        self.confidence_score = self._calculate_confidence()
    
    def _calculate_confidence(self) -> float:
        """计算整体置信度"""
        if not self.sections:
            return 0.0
        
        total_confidence = sum(section.confidence for section in self.sections)
        return min(total_confidence / len(self.sections), 1.0)


class EnhancedDocumentAnalyzer:
    """增强的文档分析器"""
    
    def __init__(self):
        self.academic_keywords = {
            'cn': [
                '摘要', '关键词', '引言', '绪论', '文献综述', '研究方法', '实验方法',
                '研究结果', '结果分析', '讨论', '结论', '参考文献', '致谢', '附录',
                '学位论文', '硕士', '博士', '学校代码', '研究生学号', '指导教师',
                '学科专业', '研究方向', '东北师范大学', '独创性声明', '使用授权书'
            ],
            'en': [
                'abstract', 'keywords', 'introduction', 'literature review',
                'methodology', 'methods', 'results', 'discussion', 'conclusion',
                'references', 'acknowledgments', 'appendix', 'thesis', 'dissertation',
                'master', 'doctor', 'phd', 'university', 'supervisor', 'advisor'
            ]
        }
        
        self.technical_keywords = [
            'api', 'sdk', 'framework', 'algorithm', 'implementation', 'code',
            'function', 'class', 'method', 'interface', 'system', 'architecture',
            'design pattern', 'database', 'server', 'client', 'protocol'
        ]
        
        self.business_keywords = [
            '市场分析', '商业模式', '财务报告', '项目管理', '战略规划', '风险评估',
            'market analysis', 'business model', 'financial report', 'project management',
            'strategic planning', 'risk assessment', 'roi', 'kpi', 'revenue'
        ]
        
        # 编译正则表达式以提高性能
        self._compile_patterns()
    
    def _compile_patterns(self):
        """编译常用正则表达式模式"""
        self.patterns = {
            'chapter_cn': re.compile(r'^第[一二三四五六七八九十\d]+章\s+(.+)', re.MULTILINE),
            'chapter_num': re.compile(r'^第?\s*(\d+)\s*章\s+(.+)', re.MULTILINE),
            'section_num': re.compile(r'^(\d+\.)+\s*(.+)', re.MULTILINE),
            'heading': re.compile(r'^#{1,6}\s+(.+)', re.MULTILINE),
            'abstract_cn': re.compile(r'摘\s*要|摘　　要', re.IGNORECASE),
            'abstract_en': re.compile(r'\babstract\b', re.IGNORECASE),
            'keywords_cn': re.compile(r'关键词[:：]', re.IGNORECASE),
            'keywords_en': re.compile(r'key\s*words?[:：]', re.IGNORECASE),
            'toc': re.compile(r'目\s*录|目　　录|table\s+of\s+contents', re.IGNORECASE),
            'references': re.compile(r'参考文献|references|bibliography', re.IGNORECASE),
            'appendix': re.compile(r'附录|appendix', re.IGNORECASE),
            'acknowledgments': re.compile(r'致谢|acknowledgments?', re.IGNORECASE),
            'email': re.compile(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'),
            'url': re.compile(r'https?://[^\s<>"{}|\\^`\[\]]+'),
            'code_block': re.compile(r'```[\s\S]*?```|`[^`]+`'),
            'citation': re.compile(r'\[[^\]]*\d+[^\]]*\]|\([^)]*\d{4}[^)]*\)'),
            'figure_ref': re.compile(r'图\s*\d+|figure\s*\d+', re.IGNORECASE),
            'table_ref': re.compile(r'表\s*\d+|table\s*\d+', re.IGNORECASE)
        }
    
    def analyze_document(self, content: str) -> DocumentStructure:
        """分析文档结构和内容"""
        logger.info("开始文档结构分析")
        
        # 预处理内容
        normalized_content = self._normalize_content(content)
        
        # 检测文档类型
        doc_type = self._detect_document_type(normalized_content)
        logger.info(f"检测到文档类型: {doc_type.value}")
        
        # 分割章节
        sections = self._extract_sections(normalized_content)
        logger.info(f"提取到 {len(sections)} 个章节")
        
        # 分类章节
        classified_sections = self._classify_sections(sections, doc_type)
        
        # 检测组件
        detected_components = self._detect_components(classified_sections)
        
        # 创建内容映射
        content_mapping = self._create_content_mapping(classified_sections)
        
        # 生成统计信息
        statistics = self._generate_statistics(normalized_content, classified_sections)
        
        return DocumentStructure(
            document_type=doc_type,
            sections=classified_sections,
            detected_components=detected_components,
            content_mapping=content_mapping,
            statistics=statistics
        )
    
    def _normalize_content(self, content: str) -> str:
        """标准化内容格式"""
        # 统一换行符
        content = content.replace('\r\n', '\n').replace('\r', '\n')
        
        # 移除多余的空行，保留段落结构
        content = re.sub(r'\n{3,}', '\n\n', content)
        
        # 统一空格
        content = re.sub(r'[ \t]+', ' ', content)
        
        return content.strip()
    
    def _detect_document_type(self, content: str) -> DocumentType:
        """检测文档类型"""
        content_lower = content.lower()
        
        # 学术论文特征检测
        academic_score = 0
        for keyword in self.academic_keywords['cn']:
            if keyword in content:
                academic_score += 2
        for keyword in self.academic_keywords['en']:
            if keyword in content_lower:
                academic_score += 1
        
        # 检测特定的学术论文标志
        if any(pattern in content for pattern in ['学位论文', '硕士', '博士', '研究生学号']):
            academic_score += 10
        
        if self.patterns['abstract_cn'].search(content) and self.patterns['abstract_en'].search(content):
            academic_score += 5
        
        # 技术文档特征检测
        technical_score = 0
        for keyword in self.technical_keywords:
            if keyword in content_lower:
                technical_score += 1
        
        if self.patterns['code_block'].search(content):
            technical_score += 3
        
        # 商务报告特征检测
        business_score = 0
        for keyword in self.business_keywords:
            if keyword in content_lower:
                business_score += 1
        
        # 根据得分判断文档类型
        if academic_score >= 10:
            return DocumentType.ACADEMIC_THESIS
        elif academic_score >= 5:
            return DocumentType.RESEARCH_PAPER
        elif technical_score >= 5:
            return DocumentType.TECHNICAL_REPORT
        elif business_score >= 3:
            return DocumentType.BUSINESS_REPORT
        else:
            return DocumentType.GENERAL_DOCUMENT
    
    def _extract_sections(self, content: str) -> List[ContentSection]:
        """提取文档章节"""
        sections = []
        lines = content.split('\n')
        current_section = None
        current_content = []
        
        for i, line in enumerate(lines):
            line = line.strip()
            
            # 检测标题
            heading_match = self._is_heading(line)
            if heading_match:
                # 保存前一个章节
                if current_section:
                    current_section.content = '\n'.join(current_content).strip()
                    current_section.end_line = i - 1
                    if current_section.content:  # 只保存非空章节
                        sections.append(current_section)
                
                # 创建新章节
                title, level = heading_match
                current_section = ContentSection(
                    name=title,
                    content='',
                    section_type=SectionType.UNKNOWN,
                    level=level,
                    start_line=i
                )
                current_content = [line]  # 包含标题行
            else:
                if current_section:
                    current_content.append(line)
                else:
                    # 文档开头的内容，创建一个前言章节
                    if line and not current_section:
                        current_section = ContentSection(
                            name="前言",
                            content='',
                            section_type=SectionType.UNKNOWN,
                            level=0,
                            start_line=0
                        )
                        current_content = [line]
                    elif current_section:
                        current_content.append(line)
        
        # 保存最后一个章节
        if current_section:
            current_section.content = '\n'.join(current_content).strip()
            current_section.end_line = len(lines) - 1
            if current_section.content:
                sections.append(current_section)
        
        return sections
    
    def _is_heading(self, line: str) -> Optional[Tuple[str, int]]:
        """判断是否为标题行，返回标题文本和级别"""
        if not line.strip():
            return None
        
        # Markdown标题
        if line.startswith('#'):
            level = len(line) - len(line.lstrip('#'))
            if level <= 6:
                title = line.lstrip('#').strip()
                return (title, level)
        
        # 中文章节标题
        chapter_match = self.patterns['chapter_cn'].match(line)
        if chapter_match:
            return (line, 1)
        
        # 数字章节标题
        chapter_num_match = self.patterns['chapter_num'].match(line)
        if chapter_num_match:
            return (line, 1)
        
        # 编号标题
        section_match = self.patterns['section_num'].match(line)
        if section_match:
            level = line.count('.') + 1
            return (line, min(level, 6))
        
        return None
    
    def _classify_sections(self, sections: List[ContentSection], doc_type: DocumentType) -> List[ContentSection]:
        """对章节进行分类"""
        classified = []
        
        for section in sections:
            section_type, confidence = self._classify_single_section(section, doc_type)
            section.section_type = section_type
            section.confidence = confidence
            classified.append(section)
        
        return classified
    
    def _classify_single_section(self, section: ContentSection, doc_type: DocumentType) -> Tuple[SectionType, float]:
        """分类单个章节"""
        title_lower = section.name.lower()
        content_lower = section.content.lower()
        
        # 摘要检测
        if self.patterns['abstract_cn'].search(section.name):
            return (SectionType.ABSTRACT_CN, 0.95)
        if self.patterns['abstract_en'].search(section.name):
            return (SectionType.ABSTRACT_EN, 0.95)
        
        # 关键词检测
        if self.patterns['keywords_cn'].search(section.content[:100]):
            return (SectionType.KEYWORDS_CN, 0.9)
        if self.patterns['keywords_en'].search(section.content[:100]):
            return (SectionType.KEYWORDS_EN, 0.9)
        
        # 目录检测
        if self.patterns['toc'].search(section.name):
            return (SectionType.TOC, 0.95)
        
        # 参考文献检测
        if self.patterns['references'].search(section.name):
            return (SectionType.REFERENCES, 0.95)
        
        # 附录检测
        if self.patterns['appendix'].search(section.name):
            return (SectionType.APPENDIX, 0.9)
        
        # 致谢检测
        if self.patterns['acknowledgments'].search(section.name):
            return (SectionType.ACKNOWLEDGMENTS, 0.9)
        
        # 学术论文特定章节
        if doc_type == DocumentType.ACADEMIC_THESIS:
            return self._classify_academic_section(section)
        
        # 通用章节分类
        return self._classify_general_section(section)
    
    def _classify_academic_section(self, section: ContentSection) -> Tuple[SectionType, float]:
        """分类学术论文章节"""
        title_lower = section.name.lower()
        
        # 引言/绪论
        if any(keyword in title_lower for keyword in ['引言', '绪论', 'introduction']):
            return (SectionType.INTRODUCTION, 0.9)
        
        # 文献综述
        if any(keyword in title_lower for keyword in ['文献综述', 'literature review', '相关工作']):
            return (SectionType.LITERATURE_REVIEW, 0.9)
        
        # 研究方法
        if any(keyword in title_lower for keyword in ['研究方法', 'methodology', '方法', '实验方法']):
            return (SectionType.METHODOLOGY, 0.85)
        
        # 结果
        if any(keyword in title_lower for keyword in ['结果', 'results', '实验结果']):
            return (SectionType.RESULTS, 0.85)
        
        # 讨论
        if any(keyword in title_lower for keyword in ['讨论', 'discussion', '分析']):
            return (SectionType.DISCUSSION, 0.8)
        
        # 结论
        if any(keyword in title_lower for keyword in ['结论', 'conclusion']):
            return (SectionType.CONCLUSION, 0.9)
        
        # 声明相关
        if '声明' in title_lower or 'declaration' in title_lower:
            return (SectionType.DECLARATION, 0.95)
        
        if '授权' in title_lower or 'authorization' in title_lower:
            return (SectionType.AUTHORIZATION, 0.95)
        
        # 封面相关
        if any(keyword in title_lower for keyword in ['封面', 'cover', '学位论文']):
            if 'english' in title_lower or '英文' in title_lower:
                return (SectionType.ENGLISH_COVER, 0.9)
            return (SectionType.COVER_PAGE, 0.9)
        
        # 默认按层级分类
        if section.level == 1:
            return (SectionType.CHAPTER, 0.6)
        elif section.level == 2:
            return (SectionType.SECTION, 0.6)
        else:
            return (SectionType.SUBSECTION, 0.6)
    
    def _classify_general_section(self, section: ContentSection) -> Tuple[SectionType, float]:
        """分类通用章节"""
        if section.level == 1:
            return (SectionType.CHAPTER, 0.7)
        elif section.level == 2:
            return (SectionType.SECTION, 0.7)
        else:
            return (SectionType.SUBSECTION, 0.7)
    
    def _detect_components(self, sections: List[ContentSection]) -> Set[str]:
        """检测文档组件"""
        components = set()
        
        for section in sections:
            if section.section_type != SectionType.UNKNOWN:
                components.add(section.section_type.value)
        
        return components
    
    def _create_content_mapping(self, sections: List[ContentSection]) -> Dict[str, str]:
        """创建内容映射"""
        mapping = {}
        
        for section in sections:
            if section.section_type != SectionType.UNKNOWN:
                mapping[section.section_type.value] = section.content
        
        return mapping
    
    def _generate_statistics(self, content: str, sections: List[ContentSection]) -> Dict[str, Any]:
        """生成文档统计信息"""
        stats = {
            'total_characters': len(content),
            'total_words': len(content.split()),
            'total_lines': len(content.split('\n')),
            'total_sections': len(sections),
            'sections_by_type': {},
            'sections_by_level': {},
            'has_citations': bool(self.patterns['citation'].search(content)),
            'has_figures': bool(self.patterns['figure_ref'].search(content)),
            'has_tables': bool(self.patterns['table_ref'].search(content)),
            'has_code': bool(self.patterns['code_block'].search(content)),
            'has_urls': bool(self.patterns['url'].search(content)),
            'has_emails': bool(self.patterns['email'].search(content))
        }
        
        # 按类型统计章节
        for section in sections:
            section_type = section.section_type.value
            if section_type not in stats['sections_by_type']:
                stats['sections_by_type'][section_type] = 0
            stats['sections_by_type'][section_type] += 1
        
        # 按级别统计章节
        for section in sections:
            level = section.level
            if level not in stats['sections_by_level']:
                stats['sections_by_level'][level] = 0
            stats['sections_by_level'][level] += 1
        
        return stats
    
    def get_missing_components(self, template_name: str, detected_components: Set[str]) -> Set[str]:
        """获取模板中缺失的组件"""
        # 定义各种模板的标准组件
        template_components = {
            'nenu_thesis': {
                'cover_page', 'english_cover', 'declaration', 'authorization',
                'abstract_cn', 'abstract_en', 'toc', 'introduction',
                'literature_review', 'methodology', 'results', 'discussion',
                'conclusion', 'references', 'acknowledgments'
            },
            'research_paper': {
                'abstract_en', 'introduction', 'literature_review',
                'methodology', 'results', 'discussion', 'conclusion', 'references'
            },
            'technical_report': {
                'introduction', 'methodology', 'results', 'conclusion', 'references'
            },
            'business_report': {
                'introduction', 'methodology', 'results', 'conclusion'
            },
            'default': {
                'introduction', 'conclusion'
            }
        }
        
        expected_components = template_components.get(template_name, set())
        return expected_components - detected_components
    
    def analyze_content_quality(self, content: str) -> Dict[str, Any]:
        """分析内容质量"""
        quality_metrics = {
            'readability_score': self._calculate_readability(content),
            'structure_score': self._calculate_structure_score(content),
            'completeness_score': self._calculate_completeness_score(content),
            'academic_indicators': self._analyze_academic_indicators(content),
            'formatting_issues': self._detect_formatting_issues(content)
        }
        
        # 综合质量评分
        scores = [
            quality_metrics['readability_score'],
            quality_metrics['structure_score'],
            quality_metrics['completeness_score']
        ]
        quality_metrics['overall_score'] = sum(scores) / len(scores)
        
        return quality_metrics
    
    def _calculate_readability(self, content: str) -> float:
        """计算可读性评分（简化版）"""
        words = content.split()
        sentences = re.split(r'[.!?。！？]', content)
        
        if not words or not sentences:
            return 0.0
        
        avg_word_length = sum(len(word) for word in words) / len(words)
        avg_sentence_length = len(words) / len([s for s in sentences if s.strip()])
        
        # 简化的可读性评分（0-1之间）
        readability = max(0, min(1, 1 - (avg_word_length - 4) * 0.1 - (avg_sentence_length - 15) * 0.01))
        return readability
    
    def _calculate_structure_score(self, content: str) -> float:
        """计算结构评分"""
        structure_score = 0.0
        
        # 检查是否有标题结构
        if self.patterns['heading'].search(content):
            structure_score += 0.3
        
        # 检查是否有引用
        if self.patterns['citation'].search(content):
            structure_score += 0.2
        
        # 检查段落结构
        paragraphs = [p.strip() for p in content.split('\n\n') if p.strip()]
        if len(paragraphs) >= 3:
            structure_score += 0.3
        
        # 检查列表结构
        if re.search(r'^\s*[-*+]\s+', content, re.MULTILINE):
            structure_score += 0.2
        
        return min(structure_score, 1.0)
    
    def _calculate_completeness_score(self, content: str) -> float:
        """计算完整性评分"""
        completeness = 0.0
        
        # 检查基本组件
        if self.patterns['abstract_cn'].search(content) or self.patterns['abstract_en'].search(content):
            completeness += 0.2
        
        if self.patterns['references'].search(content):
            completeness += 0.3
        
        # 检查内容长度
        word_count = len(content.split())
        if word_count >= 1000:
            completeness += 0.3
        elif word_count >= 500:
            completeness += 0.2
        
        # 检查图表引用
        if self.patterns['figure_ref'].search(content) or self.patterns['table_ref'].search(content):
            completeness += 0.2
        
        return min(completeness, 1.0)
    
    def _analyze_academic_indicators(self, content: str) -> Dict[str, Any]:
        """分析学术指标"""
        indicators = {
            'citation_count': len(self.patterns['citation'].findall(content)),
            'figure_references': len(self.patterns['figure_ref'].findall(content)),
            'table_references': len(self.patterns['table_ref'].findall(content)),
            'academic_keywords_count': 0,
            'technical_terms_count': 0
        }
        
        content_lower = content.lower()
        
        # 统计学术关键词
        for keyword in self.academic_keywords['cn'] + self.academic_keywords['en']:
            indicators['academic_keywords_count'] += content_lower.count(keyword.lower())
        
        # 统计技术术语
        for keyword in self.technical_keywords:
            indicators['technical_terms_count'] += content_lower.count(keyword.lower())
        
        return indicators
    
    def _detect_formatting_issues(self, content: str) -> List[str]:
        """检测格式问题"""
        issues = []
        
        # 检查连续空行
        if re.search(r'\n{4,}', content):
            issues.append("存在过多连续空行")
        
        # 检查标点符号问题
        if re.search(r'[，。！？；：][a-zA-Z]', content):
            issues.append("中文标点后直接跟英文字符")
        
        # 检查括号匹配
        open_parens = content.count('(')
        close_parens = content.count(')')
        if open_parens != close_parens:
            issues.append("括号不匹配")
        
        # 检查引号匹配
        quotes = content.count('"')
        if quotes % 2 != 0:
            issues.append("引号不匹配")
        
        return issues


# 便捷函数
def analyze_markdown_document(content: str) -> DocumentStructure:
    """分析Markdown文档的便捷函数"""
    analyzer = EnhancedDocumentAnalyzer()
    return analyzer.analyze_document(content)


def analyze_content_quality(content: str) -> Dict[str, Any]:
    """分析内容质量的便捷函数"""
    analyzer = EnhancedDocumentAnalyzer()
    return analyzer.analyze_content_quality(content)