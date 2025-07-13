#!/usr/bin/env python3
"""
直接处理论文文件
"""

import os
import sys
sys.path.append('/Users/niqian/Documents/GitHub/autarl')

from thesis_format_converter import ThesisFormatConverter

def main():
    # 设置路径
    input_dir = '/Users/niqian/02.知识库源/01.教学储备/06.毕业论文/04-最终稿/改进版'
    output_file = '/Users/niqian/Documents/GitHub/autarl/完整论文.md'
    
    print("=== 开始处理论文 ===")
    print(f"输入目录: {input_dir}")
    print(f"输出文件: {output_file}")
    
    # 创建转换器
    converter = ThesisFormatConverter()
    
    # 合并论文
    success = converter.merge_files(input_dir, output_file)
    
    if success:
        print("\n✓ 论文合并完成！")
        print(f"输出文件: {output_file}")
        
        # 转换为Word
        word_file = output_file.replace('.md', '.docx')
        print(f"\n开始转换为Word: {word_file}")
        
        try:
            from markdown_to_word import MarkdownToWordConverter
            word_converter = MarkdownToWordConverter(template_name='graduation')
            word_success = word_converter.convert_with_python_docx(output_file, word_file)
            
            if word_success:
                print(f"✓ Word转换完成: {word_file}")
            else:
                print("✗ Word转换失败")
        except Exception as e:
            print(f"✗ Word转换出错: {str(e)}")
    else:
        print("✗ 论文合并失败")

if __name__ == "__main__":
    main()