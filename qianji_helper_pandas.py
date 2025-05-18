# -*- coding: utf-8 -*-
import datetime
import os
import sys
import time
import json

import pandas as pd
import xlwings as xw

USING_XLWINGS = True


# 加载支付宝账单数据
def load_alipay_bills(xlsx_path):
    print(f'处理 支付宝账单: {xlsx_path}')
    # 跳过头部4行信息
    df = pd.read_csv(xlsx_path, encoding='gbk', skiprows=4)
    # print('df.columns\n', df.columns)
    # print('df content\n', df)
    # 删掉最后一列无效值
    df = df.drop(df.columns[[-1]], axis=1, inplace=False)

    # 1. 数据清洗（去除无效数据和不想要的数据）
    # 1.1 先删除含有 nan 的行
    df_lite = df.dropna(axis=0, how='any')
    # 1.2 筛选出空白行并删除
    # 删除 订单编号 是空的，删除交易状态不是交易成功的
    # 使用布尔索引进行过滤
    condition1 = df_lite['商家订单号               '].str.isspace() & df_lite['付款时间                '].str.isspace()
    condition2 = df_lite['交易状态    '].str.contains('交易关闭|退款成功', na=False)

    # 合并条件
    to_drop = condition1 | condition2

    # 删除符合条件的行
    df_lite = df_lite[~to_drop]

    # 另一种遍历筛选方式：只筛选订单号为空的行
    # to_drop_index = []
    # for order_id in df_lite['商家订单号               ']:
    #     to_drop_index.append(order_id.isspace())

    # 2. 保留需要的列
    # # 方法一 按照列名称选择
    # select_cols = ['付款时间                ', '交易对方            ', '商品名称                ', '金额（元）   ',
    #                '收/支     ']
    # dst_df = df_lite[select_cols].copy()
    # print('df_lite.columns', df_lite.columns)
    #
    # # 方法二 选择连续列
    # select_cols = df.columns[3: 7]
    # df_lite = df[select_cols]
    # print('df_lite.columns', df_lite.columns)

    # 方法三 选择所有行和指定列
    dst_df = df_lite.iloc[:, [2, 7, 8, 9, 10]].copy()
    # print('df2.columns\n', df2.columns)

    # 3. 重命名列名称
    dst_df.columns = ['时间', '交易对方', '商品名称', '金额', '类型']
    # 4. 补充（新增）需要的列
    # 新增一列并赋值
    dst_df['账户1'] = '支付宝'
    dst_df['账户2'] = ''
    dst_df['分类'] = ' '
    # 新增一列并用另一列的值赋值
    dst_df['备注'] = dst_df['交易对方']
    dst_df['账单标记'] = ''
    dst_df['账单图片'] = ''
    # 5. 排序
    dst_df = dst_df[
        ['时间', '分类', '类型', '金额', '账户1', '账户2', '备注', '账单标记', '账单图片', '交易对方', '商品名称']]
    print('done')
    return dst_df


# 加载微信账单
def load_wechat_bills(xlsx_path):
    print(f'处理 微信账单: {xlsx_path}')
    # 跳过头部16行信息
    df = pd.read_csv(xlsx_path, skiprows=16)
    # 删除不需要的列
    df_lite = df.drop(columns=['支付方式', '当前状态', '交易单号', '商户单号', '备注'], inplace=False)
    # 列重命名
    df_lite.columns = ['时间', '分类', '交易对方', '商品名称', '类型', '金额']
    # 去除 ￥ 符号
    # df_lite['金额'] = df_lite['金额'].apply(lambda x: x[1:]).astype('float')
    df_lite['金额'] = df_lite['金额'].str.slice(1).astype('float')
    # 新增一列并赋值
    df_lite['账户1'] = '微信'
    df_lite['账户2'] = ''
    # 覆盖 分类 内容
    df_lite['分类'] = ' '
    # 新增一列并用另一列的值赋值
    df_lite['备注'] = df['交易对方']
    df_lite['账单标记'] = ''
    df_lite['账单图片'] = ''
    # 排序 本质上应该是按列组合成新的 dataframe
    dst_df = df_lite[
        ['时间', '分类', '类型', '金额', '账户1', '账户2', '备注', '账单标记', '账单图片', '交易对方', '商品名称']]
    print('done')
    return dst_df


def load_ccbc_bills(xlsx_path):
    print(f'处理 建设银行 账单: {xlsx_path}')
    # 跳过头部5行信息
    df = pd.read_excel(xlsx_path, skiprows=5)
    # 删掉最后一行无效值
    df_lite = df.drop(df.index[[-1]], inplace=False)
    # 列名
    # '记账日          ', '交易日期          ', '交易时间                ',
    #        '支出                ', '收入                ', '账户余额          ',
    #        '币种          ', '摘要            ', '对方账号          ', '对方户名          ',
    #        '交易地点                '],
    # df_lite = df.drop(columns=['记账日          ', '账户余额          ', '币种          ', '对方账号          '],
    #                       inplace=False)
    df_lite['时间'] = df_lite['交易日期          '].astype(str).str[:4] + '/' + df_lite['交易日期          '].astype(
        str).str[4:6] + '/' + df_lite['交易日期          '].astype(str).str[6:8] + ' ' + df_lite[
                          '交易时间                ']
    df_lite['类型'] = df_lite['支出                '].apply(
        lambda x: '支出' if float(str(x).replace(',', '')) > 0.00 else '收入')
    df_lite['金额'] = df_lite.apply(
        lambda row: row['支出                '] if row['类型'] == '支出' else row['收入                '], axis=1)

    df_lite = df_lite.drop(columns=['交易日期          ', '交易时间                '], inplace=False)

    # 新增一列并赋值
    df_lite['账户1'] = '建设银行'
    df_lite['账户2'] = ''
    df_lite['分类'] = ' '
    df_lite['备注'] = df['对方户名          ']
    df_lite['账单标记'] = ''
    df_lite['账单图片'] = ''
    # 排序 本质上应该是按列组合成新的 dataframe
    dst_df = df_lite[
        ['时间', '分类', '类型', '金额', '账户1', '账户2', '备注', '账单标记', '账单图片', '对方户名          ',
         '交易地点                ']]
    # 列重命名
    dst_df.columns = ['时间', '分类', '类型', '金额', '账户1', '账户2', '备注', '账单标记', '账单图片', '交易对方',
                      '商品名称']
    print('done')
    return dst_df


def load_cib_bills(xlsx_path):
    print(f'处理 兴业银行 账单: {xlsx_path}')
    # 读入并跳过头部10行信息
    df = pd.read_excel(xlsx_path, skiprows=10)
    # 删掉最后一行无效值
    df_lite = df.drop(df.index[[-1]], inplace=False)
    df_new = pd.DataFrame()
    df_new['时间'] = df_lite['交易时间']
    df_new['类型'] = df_lite['支出'].apply(
        lambda x: '支出' if float(str(x).replace(',', '')) > 0.00 else '收入')
    df_new['金额'] = df_lite.apply(
        lambda row: row['支出'] if float(str(row['支出']).replace(',', '')) > 0 else row['收入'], axis=1)
    df_new['金额'] = df_new['金额'].apply(lambda x: float(str(x).replace(',', '')))

    # 新增列并赋值
    df_new['账户1'] = '兴业银行'
    df_new['账户2'] = ''
    df_new['分类'] = ' '
    df_new['备注'] = df['用途']
    df_new['账单标记'] = ''
    df_new['账单图片'] = ''
    print('done')
    return df_new


class QianJiHelper(object):

    def __init__(self, xlsx_name):
        self.xlsx_name = xlsx_name  # 日期
        self.create_new_qianji_xlsx()

    def create_new_qianji_xlsx(self):
        if USING_XLWINGS:
            print('create_new_xlsx:', self.xlsx_name)
            import os
            if os.path.exists(self.xlsx_name):
                print('will overwrite exported file!')
            # 方法1：
            # 创建一个新的App，并在新App中新建一个Book
            wb = xw.Book()
            sheet1 = wb.sheets["sheet1"]
            sheet1.clear()
            sheet1.range('A1').value = [
                ['时间', '分类', '类型', '金额', '账户1', '账户2', '备注', '账单标记', '账单图片',
                 '交易对方 / 对方名称',
                 '交易地点 / 商品名称']]
            sheet1.range('A1').column_width = 16
            sheet1.range('G1').column_width = 20
            sheet1.range('J1').column_width = 20
            sheet1.range('K1').column_width = 20
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
            sheet1.range((1 + max_row, 1)).value = bills_df.values
            # sheet1.range("a2").value = bills_df.values
            wb.save()
            wb.close()
        else:
            bills_df.to_excel(self.xlsx_name)
        print('write_data done')


def get_files(dir):
    filenames = os.listdir(dir)
    bill_file = [os.path.join(dir, filename) for filename in filenames if
                 filename.endswith(('.csv', '.xlsx', '.xls'))]
    print('bill_file:', bill_file)
    return bill_file


def get_bills(bill_files):
    wechat_bills_df = pd.DataFrame()
    alipy_bills_df = pd.DataFrame()
    ccbc_bills_df = pd.DataFrame()
    cib_bills_df = pd.DataFrame()
    for file in bill_files:
        if os.path.basename(file).startswith('微信'):
            wechat_bills_df = load_wechat_bills(xlsx_path=file)
        if os.path.basename(file).startswith('alipay'):
            alipy_bills_df = load_alipay_bills(xlsx_path=file)
        if os.path.basename(file).startswith('交易明细'):  # 建设银行
            ccbc_bills_df = load_ccbc_bills(xlsx_path=file)
        if os.path.basename(file).startswith('兴业银行'):
            cib_bills_df = load_cib_bills(xlsx_path=file)
    # 组合df
    df_all = pd.concat([wechat_bills_df, alipy_bills_df, ccbc_bills_df, cib_bills_df], ignore_index=True)
    return df_all


def load_keyword_mapping(file_path):
    with open(file_path, 'r', encoding='utf-8') as file:
        mapping = json.load(file)
    return mapping


def classify_text(text, keyword_to_category_mapping):
    if pd.isnull(text):
        return None
    text_lower = text.lower()
    for keyword, category in keyword_to_category_mapping.items():
        if keyword.lower() in text_lower:  # 不区分大小写匹配
            return category
    return None


if __name__ == '__main__':
    start_time = time.time()
    # 输入目录
    if len(sys.argv) > 1:
        input_dir = sys.argv[1]
    else:
        # 当前目录
        input_dir = os.path.dirname(os.path.abspath(__file__))
    print('input_dir:', input_dir)
    bill_files = get_files(input_dir)
    if len(bill_files) == 0:
        print('当前目录下没有账单文件')
        sys.exit()

    # 创建qianji 账单模板 excel
    # qianji_helper = QianJiHelper(xlsx_name=output_name)

    df_all = get_bills(bill_files)

    # 加载关键字到类别的映射
    keyword_to_category_mapping = load_keyword_mapping('category_mapping.json')

    for index, row in df_all.iterrows():
        remark = row['备注']
        if pd.isnull(remark):
            continue
        category = classify_text(remark, keyword_to_category_mapping)
        if category is not None:
            # print(f"备注: '{remark}' -> 类别: {category}")
            # row['分类'] = category # 这种方式不一定能修改 df_all 的内容
            df_all.loc[index, '分类'] = category  # 使用 .loc 直接修改 DataFrame,注意不能有重复index（concat 默认产生重复index）

    # 创建qianji 账单模板 excel
    # 以年月为文件名
    ISO_TIME_FORMAT = '%Y-%m-%d_%H%M%S'
    theTime = datetime.datetime.now().strftime(ISO_TIME_FORMAT)
    output_name = os.path.join(input_dir, theTime + '.xlsx')
    qianji_helper = QianJiHelper(xlsx_name=output_name)

    qianji_helper.write_data(df_all)

    print("耗时: {:.3f}秒".format(time.time() - start_time))
