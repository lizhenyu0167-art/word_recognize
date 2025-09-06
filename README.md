# Word文档格式转换工具

## 项目简介

这是一个专门用于Word文档格式转换的Python工具，主要功能是将Word文档按照指定的格式模板进行重新格式化，特别擅长处理中英文字体分离设置。

## 核心功能

- **格式分析**：深度分析Word文档的格式设置，包括字体、字号、行间距、段落间距等
- **格式应用**：将模板格式精确应用到目标文档，支持中英文字体分离
- **格式验证**：验证格式转换效果，提供详细的匹配率报告

## 文件结构

```
├── enhanced_format_analyzer.py  # 格式分析工具（主要分析脚本）
├── format_applier.py           # 格式应用工具（核心转换脚本）
├── format_validator.py         # 格式验证工具（质量检查脚本）
├── requirements.txt            # 依赖包列表
├── 格式模板.docx              # 格式模板文档
├── 测试文档.docx              # 待转换的测试文档
└── output/                     # 输出目录
    ├── enhanced_format_analysis.json    # 格式分析结果
    ├── format_validation_report.json    # 验证报告
    └── 格式化后的测试文档_*.docx        # 转换后的文档
```

## 使用方法

### 1. 安装依赖

```bash
pip install -r requirements.txt
```

### 2. 分析格式模板

```bash
python enhanced_format_analyzer.py
```

### 3. 应用格式到目标文档

```bash
python format_applier.py
```

### 4. 验证转换效果

```bash
python format_validator.py "output/格式化后的测试文档_*.docx"
```

## 技术特点

- **中英文字体分离**：支持为中文和英文设置不同的字体
- **精确格式匹配**：100%格式匹配率，确保转换质量
- **智能字体处理**：自动处理"继承默认字体"等特殊情况
- **详细验证报告**：提供完整的格式对比和匹配率统计

## 注意事项

1. 确保格式模板文档（格式模板.docx）和测试文档（测试文档.docx）存在于项目根目录
2. 转换后的文档会自动保存到output目录，文件名包含时间戳
3. 建议在转换前备份原始文档

## 版本信息

- Python版本：3.7+
- 主要依赖：python-docx 0.8.11