#!/usr/bin/env python3
"""
直接将论文_简单合并.md转换为人文社科格式，保留全部内容
"""

import re
import os

def convert_to_humanistic_format(input_file, output_file):
    """
    将原始文件直接转换为人文社科格式，保留所有内容
    """
    print(f"正在转换为人文社科格式: {input_file}")
    
    try:
        with open(input_file, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # 添加英文摘要
        english_abstract = """Cultural identity is an important foundation for national cohesion and national identity, and has fundamental significance in the cultural construction and social development of multi-ethnic countries. General Secretary Xi Jinping pointed out that "cultural identity is the deepest level of identity, the root of national unity and the soul of national harmony." As a multi-ethnic region in China, Yunnan has typicality and representativeness in cultural diversity and identity education. This study takes ideological and political courses in high schools in ethnic minority areas of Yunnan as the research object, deeply explores how to carry out cultural identity education in ideological and political course teaching in ethnic minority areas of high schools, focuses on Xuanwei area as an in-depth case analysis, and aims to build a cultural identity education strategy system suitable for ideological and political courses in ethnic minority areas of high schools.

Based on Habermas's communication theory, Erikson's identity theory, and Fei Xiaotong's pluralistic unity theory, this research innovatively constructs a theoretical analysis framework of "communication-identity-pluralistic unity" for cultural identity. This framework achieves three theoretical innovations: methodological innovation of cross-cultural theoretical integration, localized practice of Chinese-Western theoretical dialogue, and theoretical breakthrough in cultural identity research in multi-ethnic countries. Using mixed research methods such as literature research, questionnaire surveys, interviews, and case analysis, a comprehensive investigation and analysis was conducted on the current status and influencing factors of cultural identity among high school students in Yunnan, especially in Xuanwei area. Taking 4 main high schools in Xuanwei City as survey objects, through distributing 1,200 questionnaires (1,132 valid responses, recovery rate 94.3%), interviewing 20 teachers and 40 students, observing and recording 40 ideological and political course teaching processes, a large amount of first-hand data was collected.

The research findings show that the overall level of cultural identity among high school students in Xuanwei area of Yunnan is 3.78 points (above average), presenting the characteristics of "strong cognition, relatively strong emotion, and weak behavior", with obvious "knowledge-action separation" phenomenon. Ethnic minority students present "dual cultural identity" characteristics, identifying with both their own ethnic culture and Chinese culture, reflecting the vivid practice of "pluralistic unity" pattern. The key factors affecting cultural identity include ideological and political course teachers (r=0.47), family cultural environment (r=0.42), peer groups (r=0.39), etc. As the main channel of cultural identity education, ideological and political courses still have deficiencies in teaching content, teaching methods, and teaching resource integration, especially the attention to ethnic minority cultural characteristics and the utilization of local cultural resources need to be strengthened.

In response to the problems found in the survey, the research constructs a systematic cultural identity education strategy system: first, excavating school regional traditional culture to enhance identity with national traditional culture; second, using ideological and political course textbook resources to enhance identity with revolutionary culture and socialist advanced culture; third, exploring current affairs resources to promote the realization of cultural identity education goals; fourth, exploring innovative paths empowered by digital technology; fifth, summarizing the characteristic teaching practices in Xuanwei area and their promotion value. Based on the "communication-identity-pluralistic unity" theoretical framework, the basic concepts and principles of cultural identity education in ideological and political courses are proposed, emphasizing the promotion value of characteristic teaching practices such as "Yi-Han bilingual ideological and political courses" in Xuanwei area.

The research shows that cultural identity education in ideological and political courses in ethnic minority areas should adhere to the educational concept of "pluralistic unity", follow the educational laws of "knowledge, emotion, will, and action", pay attention to student subject participation, and achieve localization of teaching content and diversification of teaching methods. By paying attention to ethnic cultural characteristics, promoting cultural exchange and mutual learning, and using modern information technology, the cultural confidence and national identity of ethnic minority students can be effectively enhanced, laying a solid emotional and cultural foundation for forging a strong sense of Chinese national community.

This study deepens the understanding of cultural identity education in ethnic minority areas from both theoretical and practical levels, innovatively constructs a theoretical framework of cultural identity education with Chinese characteristics, provides operational strategic references for ideological and political course teaching in ethnic areas, and provides empirical support for relevant policy formulation and curriculum reform, which is of great significance for promoting the construction of a culturally strong country and an educationally strong country."""

        english_keywords = "Ideological and Political Course in High School; Cultural Identity; Teaching Strategy; Xuanwei, Yunnan; Ethnic Minority Education; Consciousness of Chinese National Community"
        
        # 提取摘要和关键词
        abstract_match = re.search(r'## 摘要\n\n(.*?)\n\n\*\*关键词', content, re.DOTALL)
        abstract_content = abstract_match.group(1).strip() if abstract_match else ""
        
        keywords_match = re.search(r'\*\*关键词\*\*：(.+)', content)
        keywords_content = keywords_match.group(1).strip() if keywords_match else ""
        
        # 找到摘要结束后的位置，保留所有后续内容
        abstract_end = content.find('**关键词**')
        if abstract_end != -1:
            keywords_end = content.find('\n', content.find('**关键词**') + len('**关键词**'))
            if keywords_end != -1:
                remaining_content = content[keywords_end:].strip()
            else:
                remaining_content = ""
        else:
            remaining_content = ""
        
        # 格式转换：将目录部分转换为人文社科格式
        remaining_content = re.sub(r'## 目录', '# 目录', remaining_content)
        
        # 转换章节标题格式
        remaining_content = re.sub(r'## 第一章 文化认同的内涵及其理论基础', '# 第一章 文化认同的内涵及其理论基础', remaining_content)
        remaining_content = re.sub(r'## 第二章', '# 第二章', remaining_content)
        remaining_content = re.sub(r'## 第三章', '# 第三章', remaining_content)
        remaining_content = re.sub(r'## 第四章', '# 第四章', remaining_content)
        remaining_content = re.sub(r'## 第五章', '# 第五章', remaining_content)
        
        # 转换节级标题（一、二、三）
        remaining_content = re.sub(r'### （([一二三四五六七八九十])）([^#\n]*)', r'## \1、\2', remaining_content)
        
        # 转换目级标题（（一）、（二）、（三））
        remaining_content = re.sub(r'#### (\d+)\. ([^#\n]*)', r'### （\1）\2', remaining_content)
        
        # 构建完整的人文社科格式论文
        formatted_content = f"""# 摘要

{abstract_content}

**关键词：** {keywords_content}

# Abstract

{english_abstract}

**Key words:** {english_keywords}

{remaining_content}
"""
        
        # 写入文件
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(formatted_content)
        
        print(f"✓ 人文社科格式转换完成: {output_file}")
        return True
        
    except Exception as e:
        print(f"✗ 格式转换失败: {str(e)}")
        return False

def main():
    input_file = '/Users/niqian/Documents/GitHub/autarl/论文_简单合并.md'
    output_file = '/Users/niqian/Documents/GitHub/autarl/论文_完整人文社科格式.md'
    
    print("=== 完整转换为人文社科格式 ===")
    print(f"输入文件: {input_file}")
    print(f"输出文件: {output_file}")
    
    if convert_to_humanistic_format(input_file, output_file):
        print(f"\n✓ 转换成功，保留了所有原始内容")
        
        # 转换为Word
        print("\n开始转换为Word...")
        try:
            import sys
            sys.path.append('/Users/niqian/Documents/GitHub/autarl')
            from markdown_to_word import MarkdownToWordConverter
            
            word_file = output_file.replace('.md', '.docx')
            converter = MarkdownToWordConverter(template_name='graduation')
            success = converter.convert_with_python_docx(output_file, word_file)
            
            if success:
                print(f"✓ Word转换完成: {word_file}")
            else:
                print("✗ Word转换失败")
                
        except Exception as e:
            print(f"✗ Word转换出错: {str(e)}")
    else:
        print("转换失败")

if __name__ == "__main__":
    main()