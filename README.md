<div align="center">

# 🔗 Supplier Merger

**供应商表格整合工具 · 按供应商编码智能合并 Excel 数据**

<br>

![Python](https://img.shields.io/badge/Python-3776AB?style=for-the-badge&logo=python&logoColor=white)
![Pandas](https://img.shields.io/badge/Pandas-150458?style=for-the-badge&logo=pandas&logoColor=white)

<br>

根据供应商编码合并两张 Excel 表，智能拼接金额与原因字段

[功能特性](#功能特性) · [快速启动](#快速启动) · [输入格式](#输入格式)

</div>

---

## 功能特性

- 🔗 按供应商编码自动匹配两张表
- 📊 单子金额、上货金额、原因字段用 `+` 拼接（不求和）
- 🔢 按供应商编码升序排列
- 📐 保留源表格式（列宽、字体、对齐）
- 📋 所有单元格带边框

---

## 快速启动

### 安装依赖

```bash
pip install pandas openpyxl
```

### 运行

```bash
python main.py
```

### 打包 exe

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name "供应商表格整合工具" main.py
```

---

## 输入格式

两张 Excel 表需包含以下列：

| 列名 | 说明 |
|------|------|
| 供应商编码 | 主键，用于匹配 |
| 供应商名称 | 供应商名称 |
| 单子金额 | 用 `+` 拼接 |
| 上货金额 | 用 `+` 拼接 |
| 原因 | 用 `+` 拼接 |
| 开户名称 | 保留表 1 数据 |
| 帐号 | 保留表 1 数据 |
| 开户行信息 | 保留表 1 数据 |

### 示例

**表 1：**
| 供应商编码 | 单子金额 |
|-----------|---------|
| 101001 | 100 |

**表 2：**
| 供应商编码 | 单子金额 |
|-----------|---------|
| 101001 | 200 |

**合并结果：**
| 供应商编码 | 单子金额 |
|-----------|---------|
| 101001 | 100+200 |

---

## 开源许可

MIT © [H1nk5](https://github.com/H1nk5)
