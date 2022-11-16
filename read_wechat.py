import time

import pandas as pd
import xlwings as xw


def load_wechat_bills(xlsx_path):
    print('load_wechat_bills')
    wb = xw.Book(xlsx_path)
    # print(wb.sheets)
    # 只有当表格打开时才有active sheet
    # print(wb.sheets.active)
    sheet1 = wb.sheets[0]
    # 输出工作簿名称
    print('sheet1 name:', sheet1.name)

    # 工作表sheet中有数据区域最大的行数，法2
    max_row = sheet1.used_range.last_cell.row
    print('max row', max_row)
    # 工作表sheet中有数据区域最大的列数，法2
    max_col = sheet1.used_range.last_cell.column
    print('max col', max_col)

    date_time = sheet1.range((18, 1), (max_row, 1)).value
    goods_type = sheet1.range((18, 2), (max_row, 2)).value
    seller = sheet1.range((18, 3), (max_row, 3)).value
    goods = sheet1.range((18, 4), (max_row, 4)).value
    in_out = sheet1.range((18, 5), (max_row, 5)).value
    amount = sheet1.range((18, 6), (max_row, 6)).value
    wb.close()

    from bill import BillInfo

    bills_list = []
    for i in range(len(date_time)):
        new_b = BillInfo(date_time=date_time[i], category=goods_type[i], seller=seller[i], account1='微信',
                         goods=goods[i], in_out=in_out[i], amount=amount[i])
        # print('new bill:', new_b.to_str())
        bills_list.append(new_b)
    print('load_wechat_bills done. cnt:', len(bills_list))
    return bills_list


def load_wechat_bills_pandas(input_path):
    print('load_wechat_bills_pandas')
    df = pd.read_csv(input_path, skiprows=16)
    df.drop(columns=['支付方式', '当前状态', '交易单号', '商户单号', '备注'], inplace=True)
    df.columns = ['时间', '分类', '交易对方', '商品', '类型', '金额']
    df['账户1'] = '微信'
    df['备注'] = df['交易对方']
    df2 = df[['时间', '分类', '类型', '金额', '账户1', '备注', '交易对方', '商品']]
    print('load_wechat_bills_pandas done')
    return df2


if __name__ == '__main__':
    path = './微信支付账单(20221001-20221031).csv'
    start_time = time.time()
    wechat_bills = load_wechat_bills(path)
    print("耗时: {:.3f}秒".format(time.time() - start_time))
    if len(wechat_bills) > 0:
        print(wechat_bills[0].to_str())

    start_time = time.time()
    wechat_bills_df = load_wechat_bills_pandas(path)
    print("耗时: {:.3f}秒".format(time.time() - start_time))
    if wechat_bills:
        print(wechat_bills_df.iloc[0].values)
