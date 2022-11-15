import datetime

import xlwings as xw


def load_alipay_bills(xlsx_path):
    print('load_alipay_bills')
    wb = xw.Book(xlsx_path)
    # print(wb.sheets)
    # print(wb.sheets.active)
    sheet1 = wb.sheets[0]
    # 输出工作簿名称
    print('sheet1 name:', sheet1.name)

    # 工作表sheet中有数据区域最大的行数，法2
    max_row = sheet1.used_range.last_cell.row - 7
    print('max row', max_row)
    # 工作表sheet中有数据区域最大的列数，法2
    max_col = sheet1.used_range.last_cell.column
    print('max col', max_col)

    date_time = sheet1.range((6, 4), (max_row, 4)).value
    in_out = sheet1.range((6, 11), (max_row, 11)).value
    amount = sheet1.range((6, 10), (max_row, 10)).value
    seller = sheet1.range((6, 8), (max_row, 8)).value
    goods = sheet1.range((6, 9), (max_row, 9)).value

    from bill import BillInfo

    bills_list = []

    for i in range(len(date_time)):
        if isinstance(date_time[i], datetime.datetime) is False:
            continue
        new_b = BillInfo(date_time=date_time[i], category=None, in_out=in_out[i], amount=amount[i], account1='支付宝',
                         seller=seller[i], goods=goods[i])
        # print('new bill:', new_b.to_str())
        bills_list.append(new_b)
    wb.close()
    print('load_alipay_bills done. cnt', len(bills_list))
    return bills_list


if __name__ == '__main__':
    path = './alipay_record_20221113_1558_1.csv'
    wechat_bills = load_alipay_bills(path)
    if len(wechat_bills) > 0:
        print('bills_list', wechat_bills[0].to_str())
