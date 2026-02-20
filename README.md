# ExpenseCombination

使用python实现合并微信个人对账单和支付宝个人交易流水证明

## 使用前的准备

1. 下载微信对账单和支付宝个人交易流水证明
2. 确保已安装 Python 3.x 和必要的依赖库

## 安装依赖

```shell
pip install pandas openpyxl keyboard
```

## 使用方法

### 方法一：使用打包好的可执行文件

下载 Release 中打包好的 exe 文件，直接双击执行。

### 方法二：使用 Python 脚本

#### 基本用法

```shell
python src/main.py
```

默认行为：
- 导出为 XLSX 格式
- 单文件模式（包含"明细"和"统计"两个工作表）
- 输出文件保存在 `./output/` 目录，文件名带时间戳

#### 命令行参数

**收支分离模式** (`-s` 或 `--separate`)
```shell
python src/main.py -s
```
- 分别导出收入和支出两张表
- 默认 XLSX 格式，输出为 `支出_{时间戳}.xlsx` 和 `收入_{时间戳}.xlsx`

**指定输出格式** (`-f` 或 `--format`)
```shell
python src/main.py -f csv
```
- 支持格式：`xlsx`（默认）或 `csv`
- CSV 格式下，明细和统计分别输出为两个文件

**组合使用**
```shell
# 收支分离 + CSV 格式
python src/main.py -s -f csv

# 查看帮助信息
python src/main.py -h
```

## 输出说明

- **默认模式**（单文件）：
  - XLSX 格式：`result_{时间戳}.xlsx`（包含"明细"和"统计"两个工作表）
  - CSV 格式：`result_{时间戳}_明细.csv` 和 `result_{时间戳}_统计.csv`

- **收支分离模式**：
  - XLSX 格式：`支出_{时间戳}.xlsx` 和 `收入_{时间戳}.xlsx`
  - CSV 格式：`支出_{时间戳}.csv` 和 `收入_{时间戳}.csv`

所有输出文件均保存在 `./output/` 目录中，文件名包含时间戳（格式：`YYYYMMDD_HHMMSS`），避免文件覆盖。

## 功能特性

- ✅ 支持微信账单（CSV/XLSX 格式）
- ✅ 支持支付宝账单（CSV 格式）
- ✅ 支持批量导入多个文件
- ✅ 自动合并微信和支付宝账单
- ✅ 自动生成月度收支统计
- ✅ 支持收支分离导出
- ✅ 支持 XLSX 和 CSV 两种输出格式
- ✅ 输出文件自动添加时间戳
