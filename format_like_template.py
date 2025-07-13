#!/usr/bin/env python3
"""
按照complete_thesis_test.md的格式来重新整理论文_简单合并.md
"""

import re

def format_thesis_like_template(input_file, output_file):
    """
    按照模板格式重新整理论文
    """
    print(f"正在按模板格式整理: {input_file}")
    
    try:
        with open(input_file, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # 1. 提取主标题
        title_match = re.search(r'^# (.+)$', content, re.MULTILINE)
        main_title = title_match.group(1) if title_match else "论文标题"
        
        # 2. 提取摘要内容
        abstract_match = re.search(r'## 摘要\n\n(.*?)\n\n\*\*关键词', content, re.DOTALL)
        abstract_content = abstract_match.group(1) if abstract_match else ""
        
        # 3. 提取关键词
        keywords_match = re.search(r'\*\*关键词\*\*：(.+)', content)
        keywords = keywords_match.group(1) if keywords_match else ""
        
        # 4. 重新构建格式规范的论文
        formatted_content = f"""# 摘要

{abstract_content}

**关键词：** {keywords}

# Abstract

(English abstract to be added)

**Key words:** (English keywords to be added)

# 第一章 引言

## 1.1 研究背景

(研究背景内容)

## 1.2 研究现状

(研究现状内容)

## 1.3 研究目的与意义

### 1.3.1 研究目的

(研究目的内容)

### 1.3.2 研究意义

(研究意义内容)

## 1.4 论文结构

(论文结构内容)

# 第二章 相关工作

## 2.1 理论基础

(理论基础内容)

## 2.2 国内外研究现状

(研究现状内容)

# 第三章 研究方法

## 3.1 研究设计

(研究设计内容)

## 3.2 数据收集

(数据收集内容)

## 3.3 数据分析

(数据分析内容)

# 第四章 结果与分析

## 4.1 研究发现

(研究发现内容)

## 4.2 结果分析

(结果分析内容)

# 第五章 结论

## 5.1 主要结论

(主要结论内容)

## 5.2 研究贡献

(研究贡献内容)

## 5.3 局限性与展望

(局限性与展望内容)

# 参考文献

(参考文献列表)
"""
        
        # 写入格式化后的文件
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(formatted_content)
        
        print(f"✓ 格式化完成: {output_file}")
        print("注意：这是按照模板格式生成的框架，您需要将原始内容填入相应章节")
        return True
        
    except Exception as e:
        print(f"✗ 格式化失败: {str(e)}")
        return False

def main():
    input_file = '/Users/niqian/Documents/GitHub/autarl/论文_简单合并.md'
    output_file = '/Users/niqian/Documents/GitHub/autarl/论文_格式化模板.md'
    
    print("=== 按照模板格式整理论文 ===")
    
    if format_thesis_like_template(input_file, output_file):
        print(f"\n✓ 已生成格式化模板: {output_file}")
        print("接下来您可以：")
        print("1. 将原始内容按章节填入模板")
        print("2. 添加英文摘要")
        print("3. 调整章节结构")
        print("4. 转换为Word格式")

if __name__ == "__main__":
    main()