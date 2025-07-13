#!/usr/bin/env python3
"""
Enhanced Markdown to Word Converter
增强的定制化Markdown转Word转换器 - 主入口模块
"""

import os
import sys
import argparse
import logging
from pathlib import Path
from typing import Dict, List, Optional, Any
import json
import time

# 导入增强模块
try:
    from enhanced_templates_config import (
        get_template, list_templates, save_template, 
        create_custom_template, TemplateConfig
    )
    from enhanced_document_analyzer import (
        analyze_markdown_document, analyze_content_quality,
        DocumentType, SectionType
    )
    from enhanced_style_engine import (
        EnhancedStyleEngine, StyleEngineFactory
    )
    from enhanced_output_formats import (
        OutputManager, OutputFormat, OutputConfig,
        convert_markdown_to_format
    )
except ImportError as e:
    print(f"导入增强模块失败: {e}")
    print("请确保所有模块文件都在正确位置")
    sys.exit(1)

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler('converter.log', encoding='utf-8')
    ]
)
logger = logging.getLogger(__name__)


class EnhancedMarkdownConverter:
    """增强的Markdown转换器"""
    
    def __init__(self, template_name: str = 'default', enable_analysis: bool = True):
        """
        初始化转换器
        
        Args:
            template_name: 模板名称
            enable_analysis: 是否启用智能文档分析
        """
        self.template_name = template_name
        self.enable_analysis = enable_analysis
        
        # 加载模板配置
        try:
            self.template_config = get_template(template_name)
            logger.info(f"已加载模板: {self.template_config.name} - {self.template_config.description}")
        except Exception as e:
            logger.error(f"加载模板失败: {e}")
            raise
        
        # 初始化组件
        self.output_manager = OutputManager(self.template_config)
        self.style_engine = StyleEngineFactory.create_optimized_engine(self.template_config)
        
        # 转换统计
        self.conversion_stats = {
            'total_conversions': 0,
            'successful_conversions': 0,
            'failed_conversions': 0,
            'total_processing_time': 0.0
        }
        
        logger.info("增强Markdown转换器初始化完成")
    
    def convert_file(self, input_file: str, output_file: str, 
                    output_format: str = 'docx', 
                    output_config: Optional[Dict[str, Any]] = None) -> bool:
        """
        转换单个文件
        
        Args:
            input_file: 输入Markdown文件路径
            output_file: 输出文件路径
            output_format: 输出格式 (docx, html, pdf, latex)
            output_config: 输出配置参数
            
        Returns:
            bool: 转换是否成功
        """
        start_time = time.time()
        self.conversion_stats['total_conversions'] += 1
        
        try:
            logger.info(f"开始转换: {input_file} -> {output_file}")
            
            # 验证输入文件
            if not os.path.exists(input_file):
                raise FileNotFoundError(f"输入文件不存在: {input_file}")
            
            # 读取文件内容
            with open(input_file, 'r', encoding='utf-8') as f:
                content = f.read()
            
            # 智能文档分析
            document_structure = None
            if self.enable_analysis:
                logger.info("开始智能文档分析...")
                document_structure = analyze_markdown_document(content)
                
                # 记录分析结果
                logger.info(f"文档类型: {document_structure.document_type.value}")
                logger.info(f"检测到组件: {', '.join(document_structure.detected_components)}")
                logger.info(f"章节数量: {len(document_structure.sections)}")
                
                # 内容质量分析
                quality = analyze_content_quality(content)
                logger.info(f"内容质量评分: {quality['overall_score']:.2f}")
            
            # 准备输出配置
            format_enum = self._parse_output_format(output_format)
            config = self._prepare_output_config(format_enum, output_config)
            
            # 执行转换
            if output_format.lower() == 'docx':
                success = self._convert_to_docx(content, document_structure, output_file, config)
            else:
                success = self.output_manager.convert_document(
                    input_file, format_enum, output_file, config
                )
            
            # 更新统计
            processing_time = time.time() - start_time
            self.conversion_stats['total_processing_time'] += processing_time
            
            if success:
                self.conversion_stats['successful_conversions'] += 1
                logger.info(f"转换成功完成，耗时: {processing_time:.2f}秒")
                
                # 生成转换报告
                if self.enable_analysis and document_structure:
                    self._generate_conversion_report(
                        input_file, output_file, document_structure, 
                        quality, processing_time
                    )
            else:
                self.conversion_stats['failed_conversions'] += 1
                logger.error("转换失败")
            
            return success
            
        except Exception as e:
            self.conversion_stats['failed_conversions'] += 1
            logger.error(f"转换过程中发生错误: {e}")
            return False
    
    def _convert_to_docx(self, content: str, document_structure: Any, 
                        output_file: str, config: OutputConfig) -> bool:
        """
        转换为Word文档（使用原有逻辑）
        """
        try:
            # 这里可以集成原有的python-docx转换逻辑
            # 或者使用pandoc进行转换
            from markdown_to_word import MarkdownToWordConverter
            
            converter = MarkdownToWordConverter(self.template_name)
            
            # 创建临时文件
            import tempfile
            with tempfile.NamedTemporaryFile(mode='w', suffix='.md', delete=False, encoding='utf-8') as temp_file:
                temp_file.write(content)
                temp_path = temp_file.name
            
            try:
                success = converter.convert_with_python_docx(temp_path, output_file)
                return success
            finally:
                os.unlink(temp_path)
                
        except ImportError:
            logger.warning("原有转换器不可用，使用输出管理器")
            # 回退到输出管理器
            return self.output_manager.convert_document(
                input_file="", output_format=OutputFormat.DOCX, 
                output_file=output_file, config=config
            )
        except Exception as e:
            logger.error(f"Word转换失败: {e}")
            return False
    
    def _parse_output_format(self, format_str: str) -> OutputFormat:
        """解析输出格式"""
        format_map = {
            'docx': OutputFormat.DOCX,
            'html': OutputFormat.HTML,
            'pdf': OutputFormat.PDF,
            'latex': OutputFormat.LATEX,
            'tex': OutputFormat.LATEX,
            'epub': OutputFormat.EPUB
        }
        
        format_lower = format_str.lower()
        if format_lower not in format_map:
            raise ValueError(f"不支持的输出格式: {format_str}")
        
        return format_map[format_lower]
    
    def _prepare_output_config(self, output_format: OutputFormat, 
                             config_dict: Optional[Dict[str, Any]]) -> OutputConfig:
        """准备输出配置"""
        config = OutputConfig(format_type=output_format)
        
        if config_dict:
            for key, value in config_dict.items():
                if hasattr(config, key):
                    setattr(config, key, value)
        
        return config
    
    def _generate_conversion_report(self, input_file: str, output_file: str,
                                  document_structure: Any, quality: Dict[str, Any],
                                  processing_time: float):
        """生成转换报告"""
        report = {
            'conversion_info': {
                'input_file': input_file,
                'output_file': output_file,
                'template': self.template_name,
                'processing_time': processing_time,
                'timestamp': time.strftime('%Y-%m-%d %H:%M:%S')
            },
            'document_analysis': {
                'document_type': document_structure.document_type.value,
                'total_sections': len(document_structure.sections),
                'detected_components': list(document_structure.detected_components),
                'confidence_score': document_structure.confidence_score,
                'statistics': document_structure.statistics
            },
            'quality_metrics': quality
        }
        
        # 保存报告
        report_file = output_file.replace(Path(output_file).suffix, '_report.json')
        with open(report_file, 'w', encoding='utf-8') as f:
            json.dump(report, f, ensure_ascii=False, indent=2)
        
        logger.info(f"转换报告已生成: {report_file}")
    
    def batch_convert(self, input_dir: str, output_dir: str, 
                     output_format: str = 'docx',
                     output_config: Optional[Dict[str, Any]] = None) -> Dict[str, bool]:
        """
        批量转换目录中的文件
        
        Args:
            input_dir: 输入目录
            output_dir: 输出目录
            output_format: 输出格式
            output_config: 输出配置
            
        Returns:
            Dict[str, bool]: 每个文件的转换结果
        """
        input_path = Path(input_dir)
        output_path = Path(output_dir)
        
        # 创建输出目录
        output_path.mkdir(parents=True, exist_ok=True)
        
        # 查找所有Markdown文件
        md_files = list(input_path.glob('**/*.md'))
        
        if not md_files:
            logger.warning(f"在目录 {input_dir} 中未找到Markdown文件")
            return {}
        
        logger.info(f"找到 {len(md_files)} 个Markdown文件")
        
        results = {}
        format_enum = self._parse_output_format(output_format)
        
        for md_file in md_files:
            try:
                # 计算相对路径
                relative_path = md_file.relative_to(input_path)
                output_file = output_path / relative_path.with_suffix(f'.{output_format}')
                
                # 创建输出文件的父目录
                output_file.parent.mkdir(parents=True, exist_ok=True)
                
                # 转换文件
                success = self.convert_file(
                    str(md_file), str(output_file), 
                    output_format, output_config
                )
                
                results[str(md_file)] = success
                
            except Exception as e:
                logger.error(f"处理文件 {md_file} 时发生错误: {e}")
                results[str(md_file)] = False
        
        # 输出批量转换统计
        success_count = sum(results.values())
        logger.info(f"批量转换完成: {success_count}/{len(md_files)} 个文件成功")
        
        return results
    
    def get_conversion_statistics(self) -> Dict[str, Any]:
        """获取转换统计信息"""
        stats = self.conversion_stats.copy()
        
        if stats['total_conversions'] > 0:
            stats['success_rate'] = stats['successful_conversions'] / stats['total_conversions'] * 100
            stats['average_processing_time'] = stats['total_processing_time'] / stats['total_conversions']
        else:
            stats['success_rate'] = 0.0
            stats['average_processing_time'] = 0.0
        
        return stats
    
    def optimize_performance(self):
        """优化性能设置"""
        self.style_engine.optimize_performance()
        logger.info("性能优化完成")
    
    def export_template(self, template_name: str, output_file: str, 
                       format_type: str = 'json'):
        """导出模板配置"""
        try:
            template = get_template(template_name)
            save_template(template, format_type)
            logger.info(f"模板 '{template_name}' 已导出")
        except Exception as e:
            logger.error(f"导出模板失败: {e}")
    
    def create_custom_template_from_base(self, name: str, description: str,
                                       base_template: str = 'default',
                                       modifications: Optional[Dict[str, Any]] = None) -> bool:
        """基于现有模板创建自定义模板"""
        try:
            template = create_custom_template(name, description, base_template)
            
            # 应用修改
            if modifications:
                # 这里可以添加模板修改逻辑
                pass
            
            save_template(template)
            logger.info(f"自定义模板 '{name}' 创建成功")
            return True
            
        except Exception as e:
            logger.error(f"创建自定义模板失败: {e}")
            return False


def main():
    """主函数"""
    parser = argparse.ArgumentParser(
        description='增强的Markdown转Word转换器',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例用法:
  %(prog)s input.md -o output.docx                    # 基本转换
  %(prog)s input.md -o output.html -t nenu_thesis    # 使用特定模板转换为HTML
  %(prog)s input_dir --batch -o output_dir -f pdf    # 批量转换为PDF
  %(prog)s --list-templates                           # 列出所有可用模板
  %(prog)s --analyze input.md                         # 仅分析文档结构
        """
    )
    
    # 基本参数
    parser.add_argument('input', nargs='?', help='输入的Markdown文件或目录')
    parser.add_argument('-o', '--output', help='输出文件或目录')
    parser.add_argument('-f', '--format', choices=['docx', 'html', 'pdf', 'latex', 'epub'], 
                       default='docx', help='输出格式 (默认: docx)')
    parser.add_argument('-t', '--template', default='default', 
                       help='使用的模板 (默认: default)')
    
    # 模式参数
    parser.add_argument('--batch', action='store_true', help='批量转换模式')
    parser.add_argument('--analyze', action='store_true', help='仅分析文档结构，不转换')
    
    # 模板管理
    parser.add_argument('--list-templates', action='store_true', help='列出所有可用模板')
    parser.add_argument('--export-template', help='导出模板配置')
    parser.add_argument('--create-template', nargs=2, metavar=('NAME', 'DESCRIPTION'),
                       help='创建自定义模板')
    
    # 输出选项
    parser.add_argument('--quality', choices=['low', 'medium', 'high'], 
                       default='high', help='输出质量 (默认: high)')
    parser.add_argument('--no-toc', action='store_true', help='不包含目录')
    parser.add_argument('--no-analysis', action='store_true', help='禁用智能文档分析')
    
    # 其他选项
    parser.add_argument('--verbose', '-v', action='store_true', help='详细输出')
    parser.add_argument('--config', help='配置文件路径')
    
    args = parser.parse_args()
    
    # 设置日志级别
    if args.verbose:
        logging.getLogger().setLevel(logging.DEBUG)
    
    # 处理模板管理命令
    if args.list_templates:
        templates = list_templates()
        print("可用模板:")
        for name, description in templates.items():
            print(f"  {name:20} - {description}")
        return
    
    if args.export_template:
        converter = EnhancedMarkdownConverter()
        converter.export_template(args.export_template, f"{args.export_template}.json")
        return
    
    if args.create_template:
        name, description = args.create_template
        converter = EnhancedMarkdownConverter()
        success = converter.create_custom_template_from_base(name, description)
        if success:
            print(f"模板 '{name}' 创建成功")
        else:
            print(f"模板 '{name}' 创建失败")
        return
    
    # 验证输入参数
    if not args.input:
        parser.error("需要指定输入文件或目录")
    
    # 创建转换器
    try:
        converter = EnhancedMarkdownConverter(
            template_name=args.template,
            enable_analysis=not args.no_analysis
        )
    except Exception as e:
        print(f"创建转换器失败: {e}")
        sys.exit(1)
    
    # 仅分析模式
    if args.analyze:
        try:
            with open(args.input, 'r', encoding='utf-8') as f:
                content = f.read()
            
            print("正在分析文档结构...")
            document_structure = analyze_markdown_document(content)
            quality = analyze_content_quality(content)
            
            print(f"\n文档分析结果:")
            print(f"  文档类型: {document_structure.document_type.value}")
            print(f"  章节数量: {len(document_structure.sections)}")
            print(f"  检测到的组件: {', '.join(document_structure.detected_components)}")
            print(f"  置信度: {document_structure.confidence_score:.2f}")
            print(f"  内容质量评分: {quality['overall_score']:.2f}")
            print(f"  统计信息:")
            print(f"    总字数: {document_structure.statistics['total_words']}")
            print(f"    总行数: {document_structure.statistics['total_lines']}")
            print(f"    包含引用: {'是' if document_structure.statistics['has_citations'] else '否'}")
            print(f"    包含图表: {'是' if document_structure.statistics['has_figures'] else '否'}")
            
        except Exception as e:
            print(f"分析失败: {e}")
            sys.exit(1)
        return
    
    # 准备输出配置
    output_config = {
        'quality': args.quality,
        'include_toc': not args.no_toc,
        'include_page_numbers': True,
        'include_headers': True,
        'include_footers': True
    }
    
    # 加载额外配置
    if args.config and os.path.exists(args.config):
        try:
            with open(args.config, 'r', encoding='utf-8') as f:
                extra_config = json.load(f)
                output_config.update(extra_config)
        except Exception as e:
            logger.warning(f"加载配置文件失败: {e}")
    
    # 执行转换
    try:
        if args.batch or os.path.isdir(args.input):
            # 批量转换
            output_dir = args.output or f"output_{args.format}"
            results = converter.batch_convert(
                args.input, output_dir, args.format, output_config
            )
            
            # 显示结果
            success_count = sum(results.values())
            total_count = len(results)
            print(f"\n批量转换完成: {success_count}/{total_count} 个文件成功")
            
            if args.verbose:
                for file, success in results.items():
                    status = "✅" if success else "❌"
                    print(f"  {status} {file}")
        
        else:
            # 单文件转换
            output_file = args.output or args.input.replace('.md', f'.{args.format}')
            
            success = converter.convert_file(
                args.input, output_file, args.format, output_config
            )
            
            if success:
                print(f"转换成功: {args.input} -> {output_file}")
            else:
                print(f"转换失败: {args.input}")
                sys.exit(1)
        
        # 显示统计信息
        if args.verbose:
            stats = converter.get_conversion_statistics()
            print(f"\n转换统计:")
            print(f"  总转换数: {stats['total_conversions']}")
            print(f"  成功率: {stats['success_rate']:.1f}%")
            print(f"  平均处理时间: {stats['average_processing_time']:.2f}秒")
    
    except KeyboardInterrupt:
        print("\n转换被用户中断")
        sys.exit(1)
    except Exception as e:
        print(f"转换过程中发生错误: {e}")
        if args.verbose:
            import traceback
            traceback.print_exc()
        sys.exit(1)


if __name__ == '__main__':
    main()