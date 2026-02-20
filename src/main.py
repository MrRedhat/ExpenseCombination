"""
Author: MrRedhat
Date: 2024-01-04
Description: 这个项目用于合并微信和支付宝的个人对账单
Version: 1.1.0
"""

import argparse
import os
from datetime import datetime
from tkinter import filedialog

import keyboard as keyboard
import pandas as pd


def get_output_time_suffix() -> str:
    """返回用于输出文件名的时间后缀，格式：YYYYMMDD_HHMMSS"""
    return datetime.now().strftime('%Y%m%d_%H%M%S')

GREEN = '\033[92m'
YELLOW = '\033[93m'
BLUE = '\033[94m'
END = '\033[0m'
WECHAT = GREEN + '微信' + END
ALIPAY = BLUE + '支付宝' + END


def get_colored_str(color: str, target: str) -> str:
    return color + target + END


# 微信账单格式：交易时间 交易类型 交易对方 商品 收/支 金额(元) 支付方式 当前状态 交易单号 商户单号 备注 共11列
def get_wechat_bill_data_from_csv(path: str) -> pd.DataFrame:
    if path == '':
        return pd.DataFrame()

    data = pd.read_csv(path, header=16, skipfooter=0, encoding='utf-8')  # 微信的对账单格式从第17行开始
    return wechat_bill_data_format(data)

def get_wechat_bill_data_from_xlsx(path: str) -> pd.DataFrame:
    data = pd.read_excel(path, header=16)  # 微信的对账单格式从第17行开始
    return wechat_bill_data_format(data)


def wechat_bill_data_format(data: pd.DataFrame) -> pd.DataFrame:
    data.rename(columns={'当前状态': '交易状态', '交易类型': '类型', '金额(元)': '金额'}, inplace=True)  # 修改列名称
    data['交易时间'] = data['交易时间'].astype('datetime64[ns]')  # 修改交易时间格式
    data['金额'] = data['金额'].map(lambda x: x.strip().strip('¥') if isinstance(x, str) else x)  # 去除金额符号
    data['金额'] = data['金额'].astype('float64')  # 修改金额格式
    data = data.drop(data[data['收/支'] == '/'].index)  # 删除不计收支项
    data['备注'].replace('/', '', inplace=True)  # 微信的空备注是'/'
    data.insert(10, '来源', "微信", allow_duplicates=True)  # tag
    print("成功修改" + get_colored_str(GREEN, str(len(data))) + "条" + WECHAT + "账单数据")
    return data

# 支付宝账单格式：交易时间 交易分类 交易对方 对方账号 商品说明 收/支 金额 收/付款方式 交易状态 交易订单号 商家订单号 备注
def get_alipay_bill_data(path: str) -> pd.DataFrame:
    if path == '':
        return pd.DataFrame()

    data = pd.read_csv(path, header=22, skipfooter=0, encoding='gbk')  # 支付宝的对账单格式从第25行开始
    data.drop(data.columns[-1], axis=1, inplace=True)

    # 修改列名匹配微信格式
    data.rename(columns={'交易分类': '类型', '商品说明': '商品', '收/付款方式': '支付方式', '交易订单号': '交易单号', '商家订单号': '商户单号'}, inplace=True)

    # 合并对方列
    data['交易对方'] = data['交易对方'] + ' ' + data['对方账号']
    data = data.drop('对方账号', axis=1)
    data.insert(10, '来源', "支付宝", allow_duplicates=True)  # tag

    # 设置值类型
    data['交易时间'] = data['交易时间'].astype('datetime64[ns]')
    data['金额'] = data['金额'].astype('float64')
    data['来源'] = data['来源'].astype('category')
    data = data.drop(data[data['收/支'] == '不计收支'].index)  # 删除不计收支项
    data = data.drop(data[data['交易状态'] == '交易关闭'].index) # 删除交易关闭项

    print("成功读取" + get_colored_str(BLUE, str(len(data))) + "条" + ALIPAY + "账单数据")
    return data


def add_year_month(df: pd.DataFrame):
    df.insert(df.columns.get_loc('交易时间') + 1, '年份', df['交易时间'].dt.year)
    df.insert(df.columns.get_loc('交易时间') + 2, '月份', df['交易时间'].dt.month)


def get_bill_data() -> pd.DataFrame:
    """
    该函数用于获取微信和支付宝账单数据并合并后返回, 统一格式为微信的账单格式。
    增加年份、月份和来源方便聚合统计
    Returns:
    pd.DataFrame: 包含合并后的微信和支付宝账单数据的数据框
    """
    print('请选择' + WECHAT + '账单文件, 若没有则点击取消')
    # 支持xlsx和csv格式
    # v1.1.0 支持多个文件
    wechat_bill_document_paths = filedialog.askopenfilenames(title='请选择要导入的微信账单',
                                                           filetypes=[('所有文件', '.*'), ('csv文件', '.csv'), ('xlsx文件', '.xlsx')])
    wechat_bill_data_list = []
    for i in range(len(wechat_bill_document_paths)):
        if wechat_bill_document_paths[i].endswith('.csv'):
            wechat_bill_data_list.append(get_wechat_bill_data_from_csv(wechat_bill_document_paths[i]))
        elif wechat_bill_document_paths[i].endswith('.xlsx'):
            wechat_bill_data_list.append(get_wechat_bill_data_from_xlsx(wechat_bill_document_paths[i]))

    wechat_bill_data = pd.concat(wechat_bill_data_list, axis=0)

    print('请选择' + ALIPAY + '账单文件, 若没有则点击取消')
    alipay_bill_document_paths = filedialog.askopenfilenames(title='请选择要导入的支付宝账单:',
                                                           filetypes=[('所有文件', '.*'), ('csv文件', '.csv')])
    alipay_bill_data_list = []
    for i in range(len(alipay_bill_document_paths)):
        alipay_bill_data_list.append(get_alipay_bill_data(alipay_bill_document_paths[i]))
    alipay_bill_data = pd.concat(alipay_bill_data_list, axis=0)

    result = pd.concat([wechat_bill_data, alipay_bill_data], axis=0)
    if len(result) == 0:
        read_input_exit('您未选择任何账单')
        exit(0)
    add_year_month(result)
    return result


def calculate_monthly_expense_by_year(data: pd.DataFrame) -> pd.DataFrame:
    # 根据年份、月份和收支进行分组，并计算每个月收入和支出的总额
    result = data.groupby(['年份', '月份', '来源', '收/支'])['金额'].sum().unstack(fill_value=0).reset_index()
    result['净收入'] = result['收入'] - result['支出']
    return result


def read_input_exit(exit_info: str):
    print(exit_info)
    print('按任意键退出程序')
    keyboard.read_key(suppress=True)


def output_result(data: pd.DataFrame, target: str = 'result', output_format: str = 'xlsx'):
    """将合并后的账单导出为一个文件：xlsx 为单文件双 sheet，csv 为明细与统计两个文件。输出到 output 文件夹并带时间后缀。"""
    if not os.path.exists('./output'):
        os.makedirs('./output')
    suffix = get_output_time_suffix()
    if output_format == 'xlsx':
        output_path = f'./output/{target}_{suffix}.xlsx'
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            data.to_excel(writer, sheet_name='明细', index=False)
            calculate_monthly_expense_by_year(data).to_excel(writer, sheet_name='统计', index=False)
        read_input_exit("数据已写入" + get_colored_str(YELLOW, output_path))
    else:
        base = f'./output/{target}_{suffix}'
        path_detail = f'{base}_明细.csv'
        path_stats = f'{base}_统计.csv'
        data.to_csv(path_detail, index=False, encoding='utf-8-sig')
        calculate_monthly_expense_by_year(data).to_csv(path_stats, index=False, encoding='utf-8-sig')
        read_input_exit("数据已写入" + get_colored_str(YELLOW, path_detail) + " 与 " + get_colored_str(YELLOW, path_stats))

def refactor_wechat_bill(path: str):
    data = get_wechat_bill_data_from_xlsx(path)
    output_path = './OriginalBills/wechat_output.xlsx'
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        data.to_excel(writer, sheet_name='微信账单格式化结果', index=False)
    read_input_exit("数据已写入" + get_colored_str(YELLOW, output_path))

def refactor_alipay_bill(path: str):
    data = get_alipay_bill_data(path)
    output_path = './OriginalBills/alipay_output.xlsx'
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        data.to_excel(writer, sheet_name='支付宝账单格式化结果', index=False)
    read_input_exit("数据已写入" + get_colored_str(YELLOW, output_path))

def combine_bills(paths: list):
    # 合并n个账单，都是已知格式，从第二行开始是数据
    data_frames = []
    for path in paths:
        data = pd.read_excel(path, header=0)
        data_frames.append(data)
    result = pd.concat(data_frames, axis=0)
    output_path = './OriginalBills/合并结果.xlsx'
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        result.to_excel(writer, sheet_name='明细', index=False)
    read_input_exit("数据已写入" + get_colored_str(YELLOW, output_path))

def expanse_by_type(path: str):
    data = pd.read_excel(path, header=0)
    # 根据D列的自定义类别进行分类，只统计H列的值为支出的行，按月份记录在表单‘类别支出’中
    # 先过滤出支出行
    data = data[data['收/支'] == '支出']
    # 然后月份和类别分组，计算每个月每个类别的支出总额
    result = data.groupby(['年份', '月份', '自定义类别'])['金额'].sum().unstack(fill_value=0).reset_index()
    # 最后一列添加一个总计
    result['总计'] = result.iloc[:, 3:].sum(axis=1)
    output_path = './OriginalBills/类别支出.xlsx'
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        result.to_excel(writer, sheet_name='类别支出', index=False)
    read_input_exit("数据已写入" + get_colored_str(YELLOW, output_path))

def export_income_expense_separately(bill_data: pd.DataFrame, output_format: str = 'xlsx'):
    """
    将收入和支出分开导出为两张表，排除金额为 0 的记录。
    支出表列：交易时间、类型、交易对方、商品、金额、支付方式、交易单号、商户单号、备注（前缀「支出 」）。
    收入表列：同上，备注前缀「收入 」。
    """
    if not os.path.exists('./output'):
        os.makedirs('./output')
    cols = ['交易时间', '类型', '交易对方', '商品', '金额', '支付方式', '交易单号', '商户单号', '备注']

    expense_data = bill_data[bill_data['金额'] != 0]
    expense_data = expense_data[expense_data['收/支'] == '支出']
    expense_data = expense_data[cols].copy()
    expense_data['备注'] = '支出 ' + expense_data['备注'].astype(str)
    expense_data['备注'] = expense_data['备注'].replace('支出 nan', '支出')

    income_data = bill_data[bill_data['金额'] != 0]
    income_data = income_data[income_data['收/支'] == '收入']
    income_data = income_data[cols].copy()
    income_data['备注'] = '收入 ' + income_data['备注'].astype(str)
    income_data['备注'] = income_data['备注'].replace('收入 nan', '收入')

    suffix = get_output_time_suffix()
    ext = output_format
    path_expense = f'./output/支出_{suffix}.{ext}'
    path_income = f'./output/收入_{suffix}.{ext}'
    if output_format == 'xlsx':
        with pd.ExcelWriter(path_expense, engine='openpyxl') as writer:
            expense_data.to_excel(writer, sheet_name='支出', index=False)
        with pd.ExcelWriter(path_income, engine='openpyxl') as writer:
            income_data.to_excel(writer, sheet_name='收入', index=False)
    else:
        expense_data.to_csv(path_expense, index=False, encoding='utf-8-sig')
        income_data.to_csv(path_income, index=False, encoding='utf-8-sig')
    read_input_exit("收支分离已写入 " + get_colored_str(YELLOW, path_expense) + " 与 " + get_colored_str(YELLOW, path_income))

def parse_args():
    parser = argparse.ArgumentParser(description='合并微信与支付宝个人对账单并导出')
    parser.add_argument(
        '-s', '--separate',
        action='store_true',
        help='开启收支分离：导出收入、支出两张表；默认不分离，导出为单文件（含明细与统计）'
    )
    parser.add_argument(
        '-f', '--format',
        choices=('xlsx', 'csv'),
        default='xlsx',
        help='输出格式：xlsx（默认）或 csv'
    )
    return parser.parse_args()


if __name__ == '__main__':
    args = parse_args()
    bill_data = get_bill_data()
    if args.separate:
        export_income_expense_separately(bill_data, output_format=args.format)
    else:
        output_result(bill_data, output_format=args.format)
