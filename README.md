# NJNU-e-LabKit

内江师范学院电子专业实验报告自动生成工具

## 简介

NJNU-e-LabKit 是一个面向电子专业学生的实验报告生成模板。通过 AI（Claude Code）自动读取实验讲义，生成符合学校格式要求的实验报告。

支持课程：数字电子技术、模拟电子技术、单片机原理、嵌入式系统、通信原理等电子专业所有实验课程。

## 功能特性

- 自动读取 PDF/Word 实验讲义
- 生成 Word 格式报告（自动填充学校模板）
- 生成 Markdown 大纲（方便手写或手动调整）
- 支持真实 Word 表格生成
- 代码块自动格式化（等宽字体）
- 图片位置自动标注

## 使用步骤

```
1. 克隆项目 → 2. 放讲义 → 3. 运行脚本 → 4. 得到报告
```

| 步骤 | 操作 | 说明 |
|------|------|------|
| 1 | 复制文件夹 | 复制到实验目录 |
| 2 | 放讲义 | 放入 `05-实验讲义/` |
| 3 | 运行脚本 | `python generate_report.py` |
| 4 | 选择方案 | Word报告 或 Markdown大纲 |
| 5 | 插入图片 | 从 `03-实验素材/` 选取 |

### 获取项目

```bash
# 克隆项目
git clone https://github.com/WeightedAverage/NJNU-LabReport.git

# 进入项目目录
cd NJNU-LabReport

# 复制到你的实验目录
xcopy /E /I . "D:\LabReports\DigitalElectronics\Lab3"
```

## 快速开始

### 1. 安装依赖

```bash
pip install python-docx PyPDF2
```

### 2. 复制模板

将整个文件夹复制到你的实验报告目录：

```
D:\LabReports\DigitalElectronics\Lab3
```

### 3. 准备实验信息

**方式一：放入实验讲义**

将 PDF 或 Word 格式的讲义放入 `05-实验讲义/` 文件夹。

**方式二：编辑实验信息**

编辑 `实验信息.txt`，按格式填写实验内容。

### 4. 生成报告

在终端中运行：

```bash
python generate_report.py
```

选择生成方式：
- **方案一**：生成 Word 报告（自动填充模板）
- **方案二**：生成 Markdown 大纲（手动复制到模板）

## 文件夹结构

```
NJNU-LabReport/
├── 01-模板/              学校实验报告模板（不要修改）
├── 02-生成的报告/        AI 生成的 Word 报告
├── 03-实验素材/          实验相关的图片素材
│   ├── 电路图/           电路原理图、接线图
│   ├── 实验截图/         软件界面、运行结果
│   ├── 数据波形/         示波器波形、时序图
│   └── 其他/             其他素材
├── 04-大纲草稿/          生成的 Markdown 大纲
├── 05-实验讲义/          放置实验讲义文件
├── 06-实验材料/          实验代码和数据
│   ├── 代码工程/         源代码、工程文件
│   └── 实验数据/         测试数据、测量结果
├── 07-使用说明/          使用说明文档
├── generate_report.py    主生成脚本
├── word_helper.py        Word 格式辅助工具
└── 实验信息.txt          备选实验信息输入
```

## 使用 Claude Code（推荐）

本项目专为 Claude Code CLI 设计，支持两种使用方式：

### 方式一：使用 Skill（推荐）

1. 安装 [Claude Code](https://claude.ai/code)
2. 在实验文件夹中启动终端
3. 运行 `claude` 进入交互模式
4. 输入：`开始写实验报告`

Claude 会自动读取讲义内容并生成报告。

### 方式二：使用 Python 脚本

```bash
python generate_report.py
```

选择生成方式：
- **方案一**：生成 Word 报告（自动填充模板）
- **方案二**：生成 Markdown 大纲（手动复制到模板）

## 实验信息格式

`实验信息.txt` 格式示例：

```
## 实验三：LED流水灯实验

### 实验目的
1. 学习 STM32 GPIO 输出配置
2. 掌握 HAL 库编程方法

### 实验原理
STM32F103 的 GPIO 可配置为多种模式...

### 硬件电路设计
本实验使用 LED 灯...

### 软件设计
```c
void led_init(void) {
    // 初始化代码
}
```

### 实验结果
LED 流水灯效果正常...

### 实验总结
通过本次实验，我掌握了...
```

## 生成效果

### Word 报告
- 自动填充学校模板格式
- 真实 Word 表格（非伪表格）
- 代码块使用等宽字体
- 图片位置标注占位符

### Markdown 大纲
- 完整的内容框架
- 数据表格留空
- 方便手动填写或手写

## 常见问题

**Q: 编码显示乱码？**
A: 脚本已添加 `sys.stdout.reconfigure(encoding='utf-8')`，如仍有问题请检查终端编码设置。

**Q: 文件被占用无法保存？**
A: 关闭 Word 中打开的报告文件，或使用新文件名。

**Q: 模板格式不对？**
A: 使用方案二生成 Markdown 大纲，手动调整格式。

## 依赖

- Python 3.7+
- python-docx
- PyPDF2（可选，用于读取 PDF 讲义）

## 许可证

MIT License

## 贡献

欢迎提交 Issue 和 Pull Request。

## 作者

![头像](https://raw.githubusercontent.com/WeightedAverage/NJNU-LabReport/main/logo.jpg)

**加权平均数** · [findme@xiaoding.club](mailto:findme@xiaoding.club)
