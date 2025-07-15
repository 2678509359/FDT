# FDT
文本去重工具介绍与使用指南

中文介绍

文本去重工具 - 多格式文件智能处理

工具简介  
文本去重工具是一款功能强大的图形界面应用程序，专门用于处理各种格式文件中的重复内容。它支持多种常见文件格式，包括TXT文本文件、Word文档（DOC/DOCX）和Excel表格（XLS/XLSX），无需安装Microsoft Office即可使用。

核心功能  
• 多格式支持：智能识别并处理多种文件格式

• 智能去重：保留原始顺序，精确识别重复内容

• 预览功能：处理前预览结果，避免误操作

• 详细统计：提供处理前后的行数对比和重复内容统计

• 现代化界面：彩色主题设计，操作直观简单

适用场景  
• 数据清洗与整理

• 日志文件分析

• 文本内容优化

• 数据库导出处理

• 日常办公文档整理

使用方式

1. 安装准备
   • 确保已安装Python 3.7+

   • 打开命令提示符/终端，运行：

     pip install pandas openpyxl xlrd
     

2. 启动工具
   • 将脚本保存为deduplicate_gui.py

   • 运行命令：

     python deduplicate_gui.py
     

3. 操作步骤
   1. 选择输入文件：点击"浏览"按钮选择要处理的文件
   2. 设置输出文件：指定处理后的文件保存位置
   3. 选择处理范围（仅Word文档）：可选全部内容/仅段落/仅表格
   4. 预览结果：点击"预览结果"查看处理效果
   5. 执行去重：确认无误后点击"执行去重"
   6. 查看结果：界面显示处理统计信息，文件保存到指定位置

4. 注意事项
   • 首次运行可能自动更新依赖库

   • DOC格式文件处理能力有限，建议转换为DOCX

   • 大型文件处理可能需要较长时间

   • 覆盖原文件前会二次确认

English Introduction

Text Deduplication Tool - Intelligent Multi-Format File Processing

Tool Overview  
The Text Deduplication Tool is a powerful GUI application designed to remove duplicate content from various file formats. It supports multiple common file types including TXT text files, Word documents (DOC/DOCX), and Excel spreadsheets (XLS/XLSX), all without requiring Microsoft Office installation.

Key Features  
• Multi-format support: Intelligently recognizes and processes various file formats

• Smart deduplication: Preserves original order while accurately identifying duplicates

• Preview function: Preview results before processing to avoid mistakes

• Detailed statistics: Provides before/after line count comparison and duplicate statistics

• Modern interface: Colorful theme design with intuitive operation

Use Cases  
• Data cleaning and organization

• Log file analysis

• Text content optimization

• Database export processing

• Daily office document management

Usage Instructions

1. Installation Preparation
   • Ensure Python 3.7+ is installed

   • Open Command Prompt/Terminal and run:

     pip install pandas openpyxl xlrd
     

2. Launching the Tool
   • Save the script as deduplicate_gui.py

   • Run the command:

     python deduplicate_gui.py
     

3. Operation Steps
   1. Select Input File: Click "Browse" to choose the file to process
   2. Set Output File: Specify where to save the processed file
   3. Select Processing Scope (Word only): Choose all content/paragraphs only/tables only
   4. Preview Results: Click "Preview Results" to see processing effect
   5. Execute Deduplication: Click "Execute Deduplication" after confirmation
   6. View Results: Interface displays processing statistics, file saved to specified location

4. Important Notes
   • First run may automatically update dependencies

   • DOC format has limited processing capability - conversion to DOCX recommended

   • Large files may require longer processing time

   • Overwriting original files requires secondary confirmation

中英对照功能说明 (Chinese-English Feature Comparison)

功能 中文说明 English Description

文件格式支持 支持TXT, DOC, DOCX, XLS, XLSX格式 Supports TXT, DOC, DOCX, XLS, XLSX formats

去重算法 保留原始顺序，精确识别重复内容 Preserves original order while accurately identifying duplicates

预览功能 处理前可预览结果，避免误操作 Preview results before processing to avoid mistakes

统计信息 提供处理前后的行数对比和重复内容统计 Provides before/after line count comparison and duplicate statistics

用户界面 彩色主题设计，操作直观简单 Colorful theme design with intuitive operation

安装要求 无需安装Microsoft Office No Microsoft Office installation required

处理速度 支持大文件处理，速度优化 Optimized for large file processing speed

安全机制 覆盖原文件前二次确认 Secondary confirmation before overwriting original files

错误处理 自动检测并修复依赖问题 Automatically detects and fixes dependency issues

工具优势 (Tool Advantages)

1. 跨平台支持  
   • 支持Windows, macOS, Linux系统  

   • Supports Windows, macOS, Linux systems

2. 免Office依赖  
   • 无需安装Microsoft Office即可处理文档  

   • Processes documents without Microsoft Office installation

3. 智能格式处理  
   • 自动识别文件格式并采用最佳处理方式  

   • Automatically recognizes file formats and applies optimal processing

4. 用户友好设计  
   • 彩色界面直观展示处理进度和结果  

   • Colorful interface visually displays processing progress and results

5. 完整处理流程  
   • 从文件选择到结果展示一站式解决  

   • One-stop solution from file selection to result display

此工具通过现代化的界面设计和强大的处理能力，为用户提供了高效、便捷的文件去重解决方案，特别适合需要处理多种格式文件的办公人员和数据分析师使用。
