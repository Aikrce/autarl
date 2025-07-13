# Markdown to Word Converter

一个功能强大的Markdown到Word文档转换器，支持表格、Mermaid图表、中文格式等特性。

## 功能特性

- ✅ **Markdown语法支持**：标题、列表、代码块、引用等
- ✅ **表格转换**：支持复杂表格，保留格式
- ✅ **Mermaid图表**：自动转换为图片嵌入
- ✅ **智能清理**：自动处理重复标题和格式问题
- ✅ **中文优化**：完美支持中文内容和编号
- ✅ **Web界面**：提供友好的Web操作界面

## 快速开始

### 命令行使用

```bash
# 基本转换
python markdown_converter.py input.md

# 指定输出文件
python markdown_converter.py input.md -o output.docx

# 禁用Mermaid支持
python markdown_converter.py input.md --no-mermaid

# 禁用自动清理
python markdown_converter.py input.md --no-clean
```

### Web界面使用

```bash
# 启动Web服务
python web_app.py

# 访问 http://localhost:8080
```

## 项目结构

```
markdown-to-word-converter/
├── core/                      # 核心转换模块
│   ├── markdown_to_word.py    # 主转换器
│   ├── enhanced_table_converter.py  # 表格转换
│   └── mermaid_converter.py   # Mermaid图表转换
├── utils/                     # 工具函数
│   ├── document_analyzer.py   # 文档分析
│   └── style_engine.py        # 样式引擎
├── templates/                 # HTML模板
├── examples/                  # 示例文件
├── web_app.py                # Web应用
├── markdown_converter.py      # 命令行工具
├── requirements.txt          # 依赖包
└── README.md                # 说明文档
```

## 安装依赖

```bash
pip install -r requirements.txt
```

## 高级功能

### Mermaid图表支持

支持所有主要的Mermaid图表类型：

- 流程图 (flowchart)
- 序列图 (sequence diagram)
- 类图 (class diagram)
- 状态图 (state diagram)
- 甘特图 (gantt)
- 饼图 (pie chart)
- 等等...

### 智能清理功能

自动处理以下问题：

- 重复的标题标记
- 双重编号系统
- 多余的空行
- 格式不一致

## 许可证

MIT License

## 贡献

欢迎提交Issue和Pull Request！
