# -*- coding: utf-8 -*-
import datetime
import os
import sys
import time

import pandas as pd
import xlwings as xw

USING_XLWINGS = True


# 加载支付宝账单数据
def load_alipay_bills(xlsx_path):
    print('load_alipay_bills')
    # 跳过头部16行信息
    df = pd.read_csv(xlsx_path, encoding='gb2312', skiprows=4)
    # print('df.columns\n', df.columns)
    # print('df content\n', df)
    # 删掉最后一列无效值
    df.drop(df.columns[[-1]], axis=1, inplace=True)

    # 1. 数据清洗（去除无效数据和不想要的数据）
    # 1.1 先删除含有 nan 的行
    df1 = df.dropna(axis=0, how='any', thresh=None, subset=None, inplace=False)
    # 1.2 筛选出空白行并删除
    # 删除 订单编号 是空的，删除交易状态不是交易成功的
    to_drop_index = []
    for index, row in df1.iterrows():
        if row['商家订单号               '].isspace():
            to_drop_index.append(True)
            continue
        if '交易关闭' in row['交易状态    '] or '退款成功' in row['交易状态    ']:
            to_drop_index.append(True)
            continue
        to_drop_index.append(False)
    # 从dataframe 中删除过滤出的行
    empty_series = pd.Series(to_drop_index)
    index_names = df1[empty_series].index
    df1.drop(index_names, inplace=True)

    # 另一种遍历筛选方式：只筛选订单号为空的行
    # to_drop_index = []
    # for order_id in df1['商家订单号               ']:
    #     to_drop_index.append(order_id.isspace())

    # 2. 保留需要的列
    # # 方法一 按照列名称选择
    # select_cols = ['付款时间                ', '交易对方            ', '商品名称                ', '金额（元）   ',
    #                '收/支     ']
    # df1 = df[select_cols]
    # print('df1.columns', df1.columns)
    #
    # # 方法二 选择连续列
    # select_cols = df.columns[3: 7]
    # df1 = df[select_cols]
    # print('df1.columns', df1.columns)

    # 方法三 选择所有行和指定列
    df_filtered = df1.iloc[:, [3, 7, 8, 9, 10]].copy()
    # print('df2.columns\n', df2.columns)

    # 3. 重命名列名称
    df_filtered.columns = ['时间', '交易对方', '商品名称', '金额', '类型']
    # 4. 补充（新增）需要的列
    # 新增一列并赋值
    df_filtered['账户1'] = '支付宝'
    df_filtered['账户2'] = ''
    df_filtered['分类'] = ' '
    # 新增一列并用另一列的值赋值
    df_filtered['备注'] = df_filtered['交易对方']
    df_filtered['账单图片'] = ''
    # 5. 排序
    df_x = df_filtered[['时间', '分类', '类型', '金额', '账户1', '账户2', '备注', '账单图片', '交易对方', '商品名称']]
    print('load_alipay_bills done')
    return df_x


# 加载微信账单
def load_wechat_bills(xlsx_path):
    print('load_wechat_bills_pandas')
    # 跳过头部16行信息
    df = pd.read_csv(xlsx_path, skiprows=16)
    # 删除不需要的列
    df_filtered = df.drop(columns=['支付方式', '当前状态', '交易单号', '商户单号', '备注'], inplace=False)
    # 列重命名
    df_filtered.columns = ['时间', '分类', '交易对方', '商品名称', '类型', '金额']
    # 新增一列并赋值
    df_filtered['账户1'] = '微信'
    df_filtered['账户2'] = ''
    # 新增一列并用另一列的值赋值
    df_filtered['备注'] = df['交易对方']
    df_filtered['账单图片'] = ''
    # 排序
    df_x = df_filtered[['时间', '分类', '类型', '金额', '账户1', '账户2', '备注', '账单图片', '交易对方', '商品名称']]
    print('load_wechat_bills_pandas done')
    return df_x


class QianJiHelper(object):

    def __init__(self, xlsx_name):
        self.xlsx_name = xlsx_name  # 日期
        self.create_new_qianji_xlsx()

    def create_new_qianji_xlsx(self):
        if USING_XLWINGS:
            print('create_new_xlsx:', self.xlsx_name)
            import os
            if os.path.exists(self.xlsx_name) is True:
                print('will overwrite exported file!')
            # 方法1：
            # 创建一个新的App，并在新App中新建一个Book
            wb = xw.Book()
            sheet1 = wb.sheets["sheet1"]
            sheet1.clear()
            sheet1.range('A1').value = [
                ['时间', '分类', '类型', '金额', '账户1', '账户2', '备注', '账单图片', '交易对方 / 对方名称',
                 '交易地点 / 商品名称']]
            sheet1.range('A1').column_width = 16
            sheet1.range('G1').column_width = 20
            sheet1.range('I1').column_width = 20
            sheet1.range('J1').column_width = 20
            wb.save(self.xlsx_name)
            wb.close()
        print('create_new_xlsx done')

    def write_data(self, bills_df):
        print('write_data in excel')
        if USING_XLWINGS:
            # 实例化一个工作表对象
            wb = xw.Book(self.xlsx_name)
            sheet1 = wb.sheets["sheet1"]
            max_row = sheet1.used_range.last_cell.row
            # 写入数据
            sheet1.range(1 + max_row, 1).value = bills_df.values
            wb.save()
            wb.close()
        else:
            bills_df.to_excel(self.xlsx_name)
        print('write_data done')


def getfiles():
    filenames = os.listdir(r'./')
    bill_file = []
    # 只返回扩展名是 csv 和 xlsx 的文件
    for filename in filenames:
        if str(filename).endswith('csv'):
            bill_file.append(filename)
        if str(filename).endswith('xlsx'):
            bill_file.append(filename)
    print('bill_file:', bill_file)
    return bill_file


if __name__ == '__main__':
    start_time = time.time()
    bill_files = getfiles()
    if len(bill_files) is 0:
        print('当前目录下没有账单文件')
        sys.exit()

    # 创建qianji 账单模板 excel
    # qianji_helper = QianJiHelper(xlsx_name=output_name)
    cols = ['时间', '分类', '类型', '金额', '账户1', '账户2', '备注', '账单图片', '交易对方 / 对方名称',
            '交易地点 / 商品名称']
    wechat_bills_df = pd.DataFrame(columns=cols, index=[0])
    alipy_bills_df = pd.DataFrame(columns=cols, index=[0])
    for file in bill_files:
        if file.startswith('微信'):
            wechat_bills_df = load_wechat_bills(xlsx_path=file)
        if file.startswith('alipay'):
            alipy_bills_df = load_alipay_bills(xlsx_path=file)
    # 组合df
    df_all = pd.concat([wechat_bills_df, alipy_bills_df])

    # todo: refine category

    # 创建qianji 账单模板 excel
    # 以年月为文件名
    ISO_TIME_FORMAT = '%Y-%m-%d %H:%M:%S'
    theTime = datetime.datetime.now().strftime(ISO_TIME_FORMAT)
    output_name = theTime[:7] + '.xlsx'
    qianji_helper = QianJiHelper(xlsx_name=output_name)

    qianji_helper.write_data(df_all)

    print("耗时: {:.3f}秒".format(time.time() - start_time))
