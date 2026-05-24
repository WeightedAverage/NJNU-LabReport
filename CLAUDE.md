# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

---

# 电子专业实验报告生成模板

本文件夹是通用实验报告生成模板，适用于电子专业所有课程（数字电子技术、模拟电子技术、单片机原理、嵌入式系统等）。

## 项目背景

使用 python-docx 读取学校 Word 模板，在模板的标记位置（如【实验预习】【实验过程】【结论与分析】）自动插入实验报告内容，生成符合格式要求的 Word 文档。

## 使用方法

### 方式一：使用 Claude Code Skill（推荐）

1. 复制整个文件夹到实验目录
2. 把实验讲义放到 "05-实验讲义/" 或编辑 "实验信息.txt"
3. 在终端运行 `claude` 进入交互模式
4. 输入：`开始写实验报告`
5. Claude 会自动读取讲义并生成报告

### 方式二：使用 Python 脚本

1. 复制整个文件夹到实验目录
2. 把实验讲义放到 "05-实验讲义/" 或编辑 "实验信息.txt"
3. 在终端运行：`python generate_report.py`
4. 选择生成方式：
   - 方案一：Word报告（自动填充模板）
   - 方案二：Markdown大纲（手动复制到模板）

## 文件夹结构

```
实验报告模板/
├── 01-模板/              学校实验报告模板（不要修改）
├── 02-生成的报告/        AI生成的Word报告
├── 03-实验素材/          放实验相关的图片
│   ├── 电路图/           电路原理图、接线图
│   ├── 实验截图/         软件界面、运行结果
│   ├── 数据波形/         示波器波形、时序图
│   └── 其他/             其他素材
├── 04-大纲草稿/          生成的Markdown/TXT大纲
├── 05-实验讲义/          放实验讲义文件
├── 06-实验材料/          实验代码和数据
│   ├── 代码工程/         实验代码、工程文件
│   └── 实验数据/         测试数据、测量结果
├── 07-使用说明/          使用说明文档
├── generate_report.py    主生成脚本
├── word_helper.py        Word辅助工具
└── 实验信息.txt          备选，没有讲义时用
```

## 两种生成方案

### 方案一：生成Word报告
- 读取 "01-模板/" 中的学校模板
- 保留原模板格式，填充完整内容
- 表格使用真实Word表格（doc.add_table()）
- 图片位置标注【此处插入xxx图】
- 保存到 "02-生成的报告/"

### 方案二：生成Markdown大纲
- 生成详细大纲，包含所有内容框架
- 数据表格留空，方便填写
- 保存到 "04-大纲草稿/"
- 用户可直接复制内容到模板，或手写

---

## 核心踩坑与解决方案（必读）

### 问题1：段落只能追加到末尾，无法在模板中间插入

**现象**：`doc.add_paragraph()` 只能在文档末尾追加段落，无法插入到模板标记位置之后。

**原因**：python-docx 的 `add_paragraph()` 设计上就是追加操作，没有提供"在指定段落后插入"的 API。

**解决**：使用 XML 层面的 `addnext()` 方法：

```python
new_para = doc.add_paragraph()          # 先创建段落（会追加到末尾）
new_para.text = '内容'
ref_paragraph._element.addnext(new_para._element)  # 用XML方法移动到目标位置
```

**踩坑点**：`addnext()` 是 lxml 的方法，操作的是底层 XML 元素。必须先 `add_paragraph()` 创建段落对象，再用 `addnext()` 移动它。不能跳过创建步骤。

---

### 问题2：所有段落样式都是"Normal"，无法批量修改

**现象**：生成的报告在 Word 中打开，选中任何段落，样式栏都显示"正文"（Normal）。虽然字号、加粗等视觉效果正确，但无法批量选择同类段落统一修改格式。

**原因**：只在 `run` 级别设置了内联格式（`run.bold = True`、`run.font.size = Pt(12)`），没有给段落分配 Word 样式。

**解决**：创建自定义段落样式并通过 `para.style` 分配：

```python
# 创建样式
h1 = doc.styles.add_style('报告标题1', 1)  # 1 = PARAGRAPH 类型
h1.font.name = '宋体'
h1.font.size = Pt(12)
h1.font.bold = True
h1.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')  # 必须设置中文字体

# 分配样式
para = doc.add_paragraph()
para.text = '一、实验目的'
para.style = doc.styles['报告标题1']  # 关键！不是在 run 上设置格式
```

**踩坑点**：
- `w:eastAsia` 属性必须单独设置，否则中文字符可能使用默认字体而不是指定的宋体。
- 样式必须在插入段落之前创建好，否则 `doc.styles['样式名']` 会报 KeyError。
- 如果模板中已有同名样式，`add_style()` 会报错，需要先检查。

---

### 问题3：模板头部图片尺寸被重置

**现象**：用 python-docx 打开模板再保存后，校徽和横幅图片的大小变了。

**原因**：python-docx 处理文档时，图片的 `wp:extent`（宽高 EMU 值）可能被重置为图片的实际像素尺寸，而不是模板中设定的显示尺寸。

**解决**：保存前手动恢复模板图片的原始 EMU 尺寸：

```python
# 模板中两张图片的原始尺寸
template_sizes = [
    ('391160', '391160'),   # 校徽: 约0.4cm正方形
    ('1700530', '318135'),  # 横幅: 约1.9cm x 0.3cm
]

for r in doc.paragraphs[0].runs:
    for idx, d in enumerate(r._element.findall('.//' + qn('w:drawing'))):
        for ext in d.findall('.//' + qn('wp:extent')):
            ext.set('cx', template_sizes[idx][0])
            ext.set('cy', template_sizes[idx][1])
```

**踩坑点**：
- 这个问题非常隐蔽，不仔细对比原模板和生成文档根本发现不了。
- EMU（English Metric Units）是 Office 的内部单位，1 cm = 360000 EMU。
- 必须在 `doc.save()` 之前执行恢复操作。

---

### 问题4：三大节编号必须独立

**现象**：编号从实验预习到结论与分析连续递增（一、二、三、四、五、六），但要求每节从"一"重新开始。

**原因**：三个 section（实验预习、实验过程、结论与分析）各自是独立的编号域，每节内部从"一"开始。

**正确编号**：
- 【实验预习】：一、实验目的 → 二、算法原理
- 【实验过程】：一、程序流程 → 二、主要函数代码及注释
- 【结论与分析】：一、运行结果与分析 → 二、实验总结

**解决**：在代码中每处理一个新 section 时，重置编号计数器。

---

### 问题5：对比图片必须左右并排

**现象**：两张对比图片（如"最邻近插值 vs 双线性插值"）上下排列，占了太多空间，不符合要求。

**解决**：使用无边框表格实现左右并排：

```python
table = doc.add_table(rows=1, cols=2)
table.alignment = WD_TABLE_ALIGNMENT.CENTER

# 移除边框
for row in table.rows:
    for cell in row.cells:
        for border_name in ['top', 'bottom', 'left', 'right']:
            # 通过XML设置无边框
            ...

# 左单元格放图片1，右单元格放图片2
cell_left = table.rows[0].cells[0]
p = cell_left.paragraphs[0]
run = p.add_run()
run.add_picture('img1.jpg', width=Cm(7))

cell_right = table.rows[0].cells[1]
p = cell_right.paragraphs[0]
run = p.add_run()
run.add_picture('img2.jpg', width=Cm(7))
```

**踩坑点**：
- 图片宽度：单张 14cm，两张并排各 7cm，三张并排各 5cm。
- 必须移除表格边框，否则会看到表格线。
- 图片下方需要中文图注，不能用英文文件名。

---

### 问题6：.doc 文件无法读取

**现象**：实验讲义是 `.doc` 格式，python-docx 打开报错。

**原因**：python-docx 只支持 `.docx`（OOXML 格式），不支持旧版 `.doc`（二进制格式）。

**解决**：需要先将 `.doc` 转为 `.docx`（可用 LibreOffice、WPS 或在线工具），或使用其他方式提取文本。

---

### 问题7：Windows 终端编码乱码

**现象**：`print()` 输出中文时出现乱码或报错。

**解决**：在脚本开头添加：

```python
import sys
sys.stdout.reconfigure(encoding='utf-8')
```

---

### 问题8：保存时文件被占用

**现象**：如果生成的报告文件在 Word 中打开，再次运行脚本保存时会报 `PermissionError`。

**解决**：使用临时文件 + 重命名策略：

```python
import random
temp_path = output_dir / f"temp_{random.randint(1000,9999)}.docx"
doc.save(str(temp_path))

if final_path.exists():
    try:
        final_path.unlink()
    except PermissionError:
        final_path = output_dir / f"{name}_v2.docx"

temp_path.rename(final_path)
```

---

## 模板结构分析

学校模板的段落结构：

| 段落索引 | 内容 | 字体 | 说明 |
|---------|------|------|------|
| 0 | 校徽+横幅图片 | - | 不可修改尺寸 |
| 1 | 标题 | 黑体 22pt 加粗 | 实验报告标题 |
| 3-5 | 信息栏 | 楷体 12pt | 姓名、学号、实验名称 |
| 7 | 【实验预习】 | 黑体 14pt 加粗 | **第一个插入点** |
| 12 | 【实验过程】 | 黑体 14pt 加粗 | **第二个插入点** |
| 17 | 【结论与分析】 | 宋体 | **第三个插入点** |

---

## 字体规范速查

| 元素 | 字体 | 字号 | 加粗 |
|------|------|------|------|
| 一级标题（一、二、三） | 黑体 | 14pt | 是 |
| 二级标题（1. 2. 3.） | 宋体 | 12pt | 是 |
| 正文 | 宋体 | 12pt | 否 |
| 代码块 | Courier New | 9pt | 否 |
| 图注 | 宋体 | 10pt | 否（斜体） |

---

## 内容结构映射

```
模板
├── 【实验预习】section
│   ├── 一、实验目的          ← 从讲义提取
│   └── 二、算法原理          ← 从讲义/代码提取
│
├── 【实验过程】section
│   ├── 一、程序流程          ← 从代码分析
│   └── 二、主要函数代码及注释  ← 从代码提取+添加注释
│
└── 【结论与分析】section
    ├── 一、运行结果与分析     ← 结果图片 + 分析文字
    └── 二、实验总结          ← 总结心得
```

---

## 可复用的工具函数清单

| 函数 | 作用 | 关键实现 |
|------|------|----------|
| `ensure_styles(doc)` | 创建4个自定义Word样式 | 检查是否存在再创建 |
| `add_paragraph_after()` | 在指定段落后插入 | `addnext()` XML方法 |
| `add_section_title_after()` | 插入一级标题 | 黑体 14pt 加粗 |
| `add_body_text_after()` | 插入正文段落 | 宋体 12pt |
| `add_code_block_after()` | 插入代码块 | Courier New 9pt，按行拆分 |
| `add_image()` | 插入图片+图注 | `add_picture()` + 居中 |
| `find_markers()` | 查找模板标记段落 | 遍历 paragraphs |
| `restore_template_images()` | 恢复模板图片尺寸 | 设置 `wp:extent` EMU值 |

---

## 总结：三个最核心的坑

1. **位置问题**：`add_paragraph()` 只能追加到末尾，必须用 `addnext()` 在指定位置插入
2. **样式问题**：不能只用 run 级别内联格式，必须创建并分配段落样式
3. **图片尺寸问题**：python-docx 会重置模板图片尺寸，必须保存前手动恢复

---

## 讲义读取
- PDF: 使用PyPDF2或pdfplumber提取文本
- Word: 使用python-docx读取
- 图片: 使用OCR或让用户手动描述内容

## 适用课程

本模板适用于电子专业所有实验课程：
- 数字电子技术
- 模拟电子技术
- 单片机原理与应用
- 嵌入式系统开发
- 通信原理
- 信号与系统
- 电工电子技术
- 其他相关课程
