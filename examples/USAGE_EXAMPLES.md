# 定制化Markdown转换器 - 使用示例

## 快速体验

```bash
# 基本使用（需要安装依赖）
python3 enhanced_markdown_converter.py --list-templates

# 分析文档结构
python3 enhanced_markdown_converter.py sample.md --analyze
```

## 核心功能展示

### 1. 智能模板系统

```python
from enhanced_templates_config import get_template, list_templates

# 查看所有可用模板
templates = list_templates()
print("可用模板:")
for name, desc in templates.items():
    print(f"  {name}: {desc}")

# 加载特定模板
nenu_template = get_template('nenu_thesis')
print(f"模板: {nenu_template.name}")
print(f"样式数量: {len(nenu_template.styles)}")
```

### 2. 智能文档分析

```python
from enhanced_document_analyzer import analyze_markdown_document

content = """
# 基于AI的图像识别研究

## 摘要
本研究提出了一种新的图像识别方法...

关键词：人工智能；图像识别；深度学习

## Abstract
This research proposes a novel image recognition method...

Key words: AI; image recognition; deep learning

## 第一章 引言
### 1.1 研究背景
随着人工智能技术的发展...

## 第二章 相关工作
### 2.1 图像识别技术
现有的图像识别技术主要包括...

## 第三章 方法
### 3.1 数据预处理
数据预处理包括以下步骤...

## 结论
本研究的主要贡献包括...

## 参考文献
[1] LeCun, Y. Deep Learning. Nature, 2015.
[2] 张三. 图像识别技术综述[J]. 计算机学报, 2020.
"""

# 分析文档结构
result = analyze_markdown_document(content)

print(f"文档类型: {result.document_type.value}")
print(f"检测到的学术组件: {', '.join(result.detected_components)}")
print(f"章节总数: {len(result.sections)}")
print(f"置信度: {result.confidence_score:.2f}")

# 详细章节信息
print("\n章节结构:")
for section in result.sections:
    print(f"  {section.section_type.value}: {section.name}")
```

### 3. 多格式输出支持

```python
from enhanced_output_formats import OutputManager, OutputFormat, OutputConfig
from enhanced_templates_config import get_template

# 创建输出管理器
template = get_template('nenu_thesis')
manager = OutputManager(template)

# 配置输出参数
html_config = OutputConfig(
    format_type=OutputFormat.HTML,
    quality='high',
    include_toc=True,
    include_page_numbers=True
)

# 转换为HTML（示例）
# success = manager.convert_document('input.md', OutputFormat.HTML, 'output.html', html_config)
```

### 4. 样式引擎使用

```python
from enhanced_style_engine import EnhancedStyleEngine
from enhanced_templates_config import get_template

# 创建样式引擎
template = get_template('nenu_thesis')
style_engine = EnhancedStyleEngine(template)

# 获取样式统计
stats = style_engine.get_style_statistics()
print(f"总样式数: {stats['total_styles']}")
print(f"样式类型分布: {stats['style_types']}")
```

## 实际使用场景

### 场景1：学位论文转换

```bash
# 分析论文结构
python3 enhanced_markdown_converter.py thesis.md --analyze

# 输出示例：
# 文档类型: academic_thesis
# 检测到的组件: abstract_cn, abstract_en, introduction, methodology, results, conclusion, references
# 章节数量: 15
# 置信度: 0.95
# 内容质量评分: 0.88

# 转换为符合格式的Word文档
python3 enhanced_markdown_converter.py thesis.md -o thesis.docx -t nenu_thesis
```

### 场景2：技术文档转换

```python
# 使用编程接口
from enhanced_markdown_converter import EnhancedMarkdownConverter

converter = EnhancedMarkdownConverter(template_name='technical_doc')

# 批量转换API文档
results = converter.batch_convert(
    input_dir='api_docs/', 
    output_dir='html_docs/', 
    output_format='html'
)

print(f"转换完成: {sum(results.values())}/{len(results)} 个文件成功")
```

### 场景3：自定义模板创建

```python
from enhanced_templates_config import (
    TemplateConfig, StyleConfig, FontConfig, ParagraphConfig,
    template_manager
)

# 创建自定义模板
custom_template = TemplateConfig(
    name="my_template",
    description="我的自定义模板",
    author="用户"
)

# 定义自定义样式
title_style = StyleConfig(
    name="Custom Title",
    font=FontConfig(name="微软雅黑", size_pt=18, bold=True),
    paragraph=ParagraphConfig(
        alignment="center",
        space_before_pt=24,
        space_after_pt=18
    )
)

custom_template.styles.append(title_style)

# 保存模板
template_manager.save_template(custom_template, 'json')
```

## 系统架构说明

### 模块结构

```
enhanced_markdown_converter.py     # 主程序入口
├── enhanced_templates_config.py   # 模板配置系统
├── enhanced_document_analyzer.py  # 智能分析器
├── enhanced_style_engine.py       # 样式引擎
└── enhanced_output_formats.py     # 输出格式支持
```

### 处理流程

1. **文档分析阶段**
   - 解析Markdown结构
   - 识别文档类型和组件
   - 评估内容质量

2. **模板匹配阶段**
   - 加载指定模板
   - 智能组件匹配
   - 样式规则应用

3. **格式转换阶段**
   - 根据目标格式处理
   - 应用模板样式
   - 生成最终文档

4. **质量保证阶段**
   - 格式验证
   - 错误检查
   - 性能优化

## 扩展开发

### 添加新的输出格式

```python
# 在enhanced_output_formats.py中扩展
class CustomOutputGenerator(EnhancedOutputGenerator):
    def generate_custom_format(self, content, structure, config, output_path):
        # 实现自定义格式转换
        pass
```

### 创建新的文档分析器

```python
# 扩展文档分析功能
class CustomDocumentAnalyzer(EnhancedDocumentAnalyzer):
    def _detect_custom_components(self, content):
        # 实现自定义组件检测
        pass
```

### 自定义样式引擎

```python
# 扩展样式处理能力
class CustomStyleEngine(EnhancedStyleEngine):
    def apply_custom_formatting(self, element, custom_rules):
        # 实现自定义格式化规则
        pass
```

## 性能优化建议

1. **批量处理**: 使用批量转换功能处理大量文件
2. **缓存优化**: 启用样式缓存提升重复转换性能
3. **内存管理**: 处理大型文档时监控内存使用
4. **并行处理**: 可以扩展支持多进程并行转换

## 故障排除

### 常见问题

1. **依赖安装问题**
   ```bash
   # 安装完整依赖
   pip install -r requirements_enhanced.txt
   ```

2. **字体显示问题**
   - 确保系统安装了所需中文字体
   - 检查字体配置文件

3. **PDF生成失败**
   - 安装WeasyPrint系统依赖
   - 验证HTML中间格式

### 调试技巧

```python
# 启用详细日志
import logging
logging.basicConfig(level=logging.DEBUG)

# 获取转换统计
converter = EnhancedMarkdownConverter()
stats = converter.get_conversion_statistics()
print(f"处理统计: {stats}")
```

---

这个增强的Markdown转换器系统为您提供了强大而灵活的文档转换能力，特别优化了学术论文的格式要求。通过模块化设计，您可以轻松扩展功能以满足特定需求。