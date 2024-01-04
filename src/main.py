"""
Author: MrRedhat
Date: 2024-01-04
Description: 这个项目用于合并微信和支付宝的个人对账单
Version: 1.0.0
"""

from tkinter import filedialog

import keyboard as keyboard
import pandas as pd

GREEN = '\033[92m'
YELLOW = '\033[93m'
BLUE = '\033[94m'
END = '\033[0m'
WECHAT = GREEN + '微信' + END
ALIPAY = BLUE + '支付宝' + END


def get_colored_str(color: str, target: str) -> str:
    return color + target + END


# 微信账单格式：交易时间 交易类型 交易对方 商品 收/支 金额(元) 支付方式 当前状态 交易单号 商户单号 备注 共11列
def get_wechat_bill_data(path: str) -> pd.DataFrame:
    if path == '':
        return pd.DataFrame()

    data = pd.read_csv(path, header=16, skipfooter=0, encoding='utf-8')  # 微信的对账单格式从第17行开始
    data.rename(columns={'当前状态': '交易状态', '交易类型': '类型', '金额(元)': '金额'}, inplace=True)  # 修改列名称
    data['交易时间'] = data['交易时间'].astype('datetime64[ns]')  # 修改交易时间格式
    data['金额'] = data['金额'].map(lambda x: x.strip().strip('¥') if isinstance(x, str) else x)  # 去除金额符号
    data['金额'] = data['金额'].astype('float64')  # 修改金额格式
    data = data.drop(data[data['收/支'] == '/'].index)  # 删除不计收支项
    data['备注'].replace('/', '', inplace=True)  # 微信的空备注是'/'
    data.insert(10, '来源', "微信", allow_duplicates=True)  # tag
    print("成功读取" + get_colored_str(GREEN, str(len(data))) + "条" + WECHAT + "账单数据")
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
    wechat_bill_document_path = filedialog.askopenfilename(title='请选择要导入的微信账单',
                                                           filetypes=[('所有文件', '.*'), ('csv文件', '.csv')])
    wechat_bill_data = get_wechat_bill_data(wechat_bill_document_path)

    print('请选择' + ALIPAY + '账单文件, 若没有则点击取消')
    alipay_bill_document_path = filedialog.askopenfilename(title='请选择要导入的支付宝账单:',
                                                           filetypes=[('所有文件', '.*'), ('csv文件', '.csv')])
    alipay_bill_data = get_alipay_bill_data(alipay_bill_document_path)

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


def output_result(data: pd.DataFrame):
    output_path = '../合并结果.xlsx'
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        data.to_excel(writer, sheet_name='明细', index=False)
        calculate_monthly_expense_by_year(data).to_excel(writer, sheet_name='统计', index=False)
    read_input_exit("数据已写入" + get_colored_str(YELLOW, output_path))


if __name__ == '__main__':
    bill_data = get_bill_data()
    output_result(bill_data)
