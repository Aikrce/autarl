#!/usr/bin/env python3
"""
论文处理完整流程脚本
自动处理论文格式转换和Word生成
"""

import os
import sys
import subprocess
from pathlib import Path
from thesis_format_converter import ThesisFormatConverter

class ThesisProcessor:
    def __init__(self, input_dir, output_dir="./output"):
        self.input_dir = Path(input_dir)
        self.output_dir = Path(output_dir)
        self.converter = ThesisFormatConverter()
        
        # 确保输出目录存在
        self.output_dir.mkdir(exist_ok=True)
    
    def process_all(self):
        """
        完整处理流程
        """
        print("=== 论文处理完整流程 ===")
        
        # 步骤1: 格式转换
        print("\n1. 开始格式转换...")
        formatted_dir = self.output_dir / "formatted_md"
        formatted_dir.mkdir(exist_ok=True)
        
        success = self.converter.batch_convert(
            str(self.input_dir), 
            str(formatted_dir)
        )
        
        if not success:
            print("格式转换失败，终止处理")
            return False
        
        # 步骤2: 合并论文
        print("\n2. 开始合并论文...")
        merged_file = self.output_dir / "完整论文.md"
        
        success = self.converter.merge_files(
            str(formatted_dir),
            str(merged_file)
        )
        
        if not success:
            print("论文合并失败，终止处理")
            return False
        
        # 步骤3: 转换为Word
        print("\n3. 开始转换为Word...")
        word_file = self.output_dir / "完整论文.docx"
        
        success = self.convert_to_word(merged_file, word_file)
        
        if success:
            print(f"\n✓ 处理完成！输出文件: {word_file}")
            return True
        else:
            print("\nWord转换失败")
            return False
    
    def convert_to_word(self, md_file, word_file):
        """
        使用autarl工具转换为Word
        """
        try:
            # 导入本地的转换器
            from markdown_to_word import MarkdownToWordConverter
            
            # 创建转换器实例，使用论文模板
            converter = MarkdownToWordConverter(template_name='graduation')
            
            # 转换文件
            success = converter.convert_with_python_docx(
                str(md_file), 
                str(word_file)
            )
            
            if success:
                print(f"✓ Word转换成功: {word_file}")
                return True
            else:
                print("✗ Word转换失败")
                return False
                
        except Exception as e:
            print(f"✗ Word转换出错: {str(e)}")
            return False
    
    def start_web_service(self):
        """
        启动Web服务进行在线转换
        """
        print("\n=== 启动Web转换服务 ===")
        print("访问 http://localhost:5002 进行在线转换")
        print("按 Ctrl+C 停止服务")
        
        try:
            subprocess.run([sys.executable, "web_app.py"], cwd=str(Path(__file__).parent))
        except KeyboardInterrupt:
            print("\n服务已停止")

def main():
    print("=== 论文处理系统 ===")
    print("1. 完整流程处理")
    print("2. 启动Web服务")
    
    choice = input("请选择操作 (1-2): ")
    
    if choice == '1':
        input_dir = input("请输入论文文件目录路径: ")
        if not os.path.exists(input_dir):
            print("目录不存在")
            return
        
        output_dir = input("请输入输出目录路径 (默认: ./output): ") or "./output"
        
        processor = ThesisProcessor(input_dir, output_dir)
        processor.process_all()
    
    elif choice == '2':
        processor = ThesisProcessor(".")
        processor.start_web_service()

if __name__ == "__main__":
    main()