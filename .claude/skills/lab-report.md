---
name: lab-report
description: 电子专业实验报告生成工具，支持从讲义提取内容生成Word报告或Markdown大纲
---

# 实验报告生成 Skill

## 触发条件
当用户说以下任意指令时触发：
- "开始写实验报告"
- "生成实验报告"
- "写实验报告"
- "lab report"

## 执行步骤

### 1. 检查实验信息来源

首先检查 `05-实验讲义/` 文件夹是否有文件：
- 如果有 PDF 文件：使用 pdfplumber 或 PyPDF2 提取文本
- 如果有 Word 文件：使用 python-docx 读取内容
- 如果没有讲义：读取 `实验信息.txt`

同时检查 `06-实验材料/` 文件夹：
- `代码工程/`：读取源代码文件（.c, .h, .py, .v, .vhd 等）
- `实验数据/`：读取数据文件（.csv, .txt, .xlsx 等）
- 这些内容可辅助生成更准确的报告

### 2. 解析实验信息

从提取的内容中识别以下部分：
- 实验名称
- 实验目的
- 实验原理
- 硬件电路设计（如有）
- 软件设计与代码（如有）
- 实验结果
- 实验总结

### 3. 询问用户选择方案

使用 AskUserQuestion 工具询问：

```
请选择生成方式：
1. 方案一：生成 Word 报告（自动填充模板）
2. 方案二：生成 Markdown 大纲（手动复制到模板）
```

### 4. 执行生成

#### 方案一：生成 Word 报告

1. 读取 `01-模板/` 中的 .docx 模板文件
2. 使用 python-docx 打开模板
3. 查找标记段落：【实验预习】、【实验过程】、【结论与分析】
4. 在对应位置插入内容：
   - 使用 `add_paragraph_after()` 插入段落
   - 使用 `add_table_after()` 插入真实表格（禁止伪表格）
   - 代码块使用 Courier New 等宽字体（9号字）
5. 图片位置标注占位符：【此处插入xxx图】
6. 保存到 `02-生成的报告/`

#### 方案二：生成 Markdown 大纲

1. 生成完整的 Markdown 文档
2. 包含所有章节框架
3. 数据表格留空，方便填写
4. 保存到 `04-大纲草稿/`

### 5. 输出结果

告诉用户：
- 报告已保存的位置
- 需要手动插入的图片位置
- 后续编辑建议

## 技术要点

### Word 表格生成
```python
# 必须使用真实表格
table = doc.add_table(rows=len(rows), cols=len(headers))
table.alignment = WD_TABLE_ALIGNMENT.CENTER

# 设置表头
for i, header in enumerate(headers):
    cell = table.rows[0].cells[i]
    cell.text = header
    for p in cell.paragraphs:
        for run in p.runs:
            run.bold = True
```

### 段落插入
```python
# 在指定位置插入段落
new_para = doc.add_paragraph()
new_para.text = "内容"
ref_paragraph._element.addnext(new_para._element)
```

### 代码块格式
```python
# 等宽字体
for run in code_para.runs:
    run.font.name = 'Courier New'
    run.font.size = Pt(9)
```

## 文件夹结构

```
01-模板/          学校模板（不要修改）
02-生成的报告/    Word报告输出
03-实验素材/      图片素材
04-大纲草稿/      Markdown大纲输出
05-实验讲义/      实验讲义输入
06-实验材料/      实验代码和数据
  ├── 代码工程/   源代码、工程文件
  └── 实验数据/   测试数据、测量结果
```

## 注意事项

1. 表格必须使用真实 Word 表格，禁止伪表格
2. 每个大节的编号重新从"一"开始
3. 代码块使用等宽字体
4. 图片位置必须标注占位符
5. 保留原模板的格式和样式
