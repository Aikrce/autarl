#!/usr/bin/env python3
"""
论文格式转换器
将多个论文md文件转换为符合东北师范大学格式要求的统一格式
"""

import os
import re
import glob
from pathlib import Path

class ThesisFormatConverter:
    def __init__(self):
        # 根据东北师范大学格式要求定义标题格式
        self.title_formats = {
            'chapter': '# 第{}章 {}',      # 章标题
            'section': '## {}',            # 二级标题
            'subsection': '### {}',        # 三级标题
            'subsubsection': '#### {}'     # 四级标题
        }
        
        self.chapter_mapping = {
            '论文整合版-第一章': '第一章 引言',
            '论文整合版-第二章': '第二章 文献综述',
            '论文整合版-第三章-第一部分': '第三章 研究方法',
            '论文整合版-第三章-第二部分': '第三章 研究设计',
            '论文整合版-第三章-第三部分': '第三章 数据收集',
            '论文整合版-第三章-第四部分': '第三章 数据分析',
            '论文整合版-结论和参考文献': '第四章 结论',
            '论文整合版-结论部分': '第四章 结论',
            '论文整合版-摘要和目录': '摘要'
        }
    
    def standardize_title_format(self, content, chapter_name):
        """
        标准化标题格式
        """
        lines = content.split('\n')
        formatted_lines = []
        
        for line in lines:
            line = line.strip()
            if not line:
                formatted_lines.append('')
                continue
                
            # 处理章标题
            if chapter_name in self.chapter_mapping:
                if line.startswith('#') and not line.startswith('##'):
                    # 替换为标准章标题格式
                    title_text = re.sub(r'^#+\s*', '', line)
                    formatted_lines.append(f"# {self.chapter_mapping[chapter_name]}")
                    continue
            
            # 处理其他标题级别
            if line.startswith('####'):
                title_text = re.sub(r'^####\s*', '', line)
                formatted_lines.append(f"#### {title_text}")
            elif line.startswith('###'):
                title_text = re.sub(r'^###\s*', '', line)
                formatted_lines.append(f"### {title_text}")
            elif line.startswith('##'):
                title_text = re.sub(r'^##\s*', '', line)
                formatted_lines.append(f"## {title_text}")
            elif line.startswith('#'):
                title_text = re.sub(r'^#\s*', '', line)
                formatted_lines.append(f"# {title_text}")
            else:
                formatted_lines.append(line)
        
        return '\n'.join(formatted_lines)
    
    def add_thesis_structure(self, content):
        """
        添加论文结构元素
        """
        # 添加摘要格式
        if '摘要' in content:
            content = re.sub(r'摘要', '# 摘要', content)
        
        # 添加关键词格式
        if '关键词' in content:
            content = re.sub(r'关键词[：:]', '**关键词：**', content)
        
        # 添加图表格式
        content = re.sub(r'图(\d+)', r'图 \1', content)
        content = re.sub(r'表(\d+)', r'表 \1', content)
        
        return content
    
    def convert_single_file(self, input_path, output_path=None):
        """
        转换单个文件
        """
        try:
            with open(input_path, 'r', encoding='utf-8') as f:
                content = f.read()
            
            # 获取文件名作为章节名
            file_name = Path(input_path).stem
            
            # 标准化格式
            formatted_content = self.standardize_title_format(content, file_name)
            formatted_content = self.add_thesis_structure(formatted_content)
            
            # 输出文件
            if output_path is None:
                output_path = input_path.replace('.md', '_formatted.md')
            
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(formatted_content)
            
            print(f"✓ 转换完成: {input_path} -> {output_path}")
            return True
            
        except Exception as e:
            print(f"✗ 转换失败: {input_path} - {str(e)}")
            return False
    
    def merge_files(self, input_dir, output_path):
        """
        合并多个文件为完整论文
        """
        # 定义文件顺序
        file_order = [
            '论文整合版-摘要和目录.md',
            '论文整合版-第一章.md',
            '论文整合版-第二章.md',
            '论文整合版-第三章-第一部分.md',
            '论文整合版-第三章-第二部分.md',
            '论文整合版-第三章-第三部分.md',
            '论文整合版-第三章-第四部分.md',
            '论文整合版-结论和参考文献.md',
            '论文整合版-结论部分.md'
        ]
        
        merged_content = []
        
        for file_name in file_order:
            file_path = os.path.join(input_dir, file_name)
            if os.path.exists(file_path):
                try:
                    with open(file_path, 'r', encoding='utf-8') as f:
                        content = f.read()
                    
                    # 格式化内容
                    formatted_content = self.standardize_title_format(content, file_name.replace('.md', ''))
                    formatted_content = self.add_thesis_structure(formatted_content)
                    
                    merged_content.append(formatted_content)
                    merged_content.append('\n\n---\n\n')  # 添加分隔符
                    
                    print(f"✓ 已处理: {file_name}")
                    
                except Exception as e:
                    print(f"✗ 处理失败: {file_name} - {str(e)}")
        
        # 写入合并后的文件
        if merged_content:
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write('\n'.join(merged_content))
            
            print(f"✓ 合并完成: {output_path}")
            return True
        else:
            print("✗ 没有找到可合并的文件")
            return False
    
    def batch_convert(self, input_dir, output_dir=None):
        """
        批量转换目录中的所有md文件
        """
        if output_dir is None:
            output_dir = input_dir + '_formatted'
        
        os.makedirs(output_dir, exist_ok=True)
        
        md_files = glob.glob(os.path.join(input_dir, '*.md'))
        success_count = 0
        
        for md_file in md_files:
            file_name = os.path.basename(md_file)
            output_path = os.path.join(output_dir, file_name.replace('.md', '_formatted.md'))
            
            if self.convert_single_file(md_file, output_path):
                success_count += 1
        
        print(f"批量转换完成: {success_count}/{len(md_files)} 个文件成功")
        return success_count > 0

def main():
    converter = ThesisFormatConverter()
    
    print("=== 论文格式转换器 ===")
    print("1. 单文件转换")
    print("2. 批量转换")
    print("3. 合并论文")
    
    choice = input("请选择操作 (1-3): ")
    
    if choice == '1':
        input_file = input("请输入md文件路径: ")
        if os.path.exists(input_file):
            converter.convert_single_file(input_file)
        else:
            print("文件不存在")
    
    elif choice == '2':
        input_dir = input("请输入目录路径: ")
        if os.path.exists(input_dir):
            converter.batch_convert(input_dir)
        else:
            print("目录不存在")
    
    elif choice == '3':
        input_dir = input("请输入论文文件所在目录: ")
        output_file = input("请输入输出文件路径: ")
        if os.path.exists(input_dir):
            converter.merge_files(input_dir, output_file)
        else:
            print("目录不存在")

if __name__ == "__main__":
    main()