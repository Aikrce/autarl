# README

## 使用说明

这个Python脚本可以将Markdown文件转换为Word文档。

### 功能特点

- 支持单文件转换
- 支持批量转换
- 保留格式（标题、列表、代码块等）
- 支持两种转换方法：pandoc和python-docx

### 安装依赖

```bash
pip install -r requirements.txt
```

### 使用方法

#### 单文件转换

```bash
python markdown_to_word.py input.md -o output.docx
```

#### 批量转换

```bash
python markdown_to_word.py input_directory --batch -o output_directory
```

### 注意事项

- 推荐安装pandoc以获得更好的转换效果
- python-docx方法对复杂格式的支持有限