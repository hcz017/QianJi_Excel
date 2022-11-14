import xlwings as xw


def load_wechat_bills(xlsx_path):
    print('load_wechat_bills')
    wb = xw.Book(xlsx_path)
    print(wb.sheets)
    # print(wb.sheets.active)
    sheet1 = wb.sheets[0]
    # 输出工作簿名称
    print(sheet1.name)
    # 写入值
    # sheet1.range('A1').value = 'python知识学堂'
    # 读值并打印
    print('value of A1:', sheet1.range('A1').value)

    # 工作表sheet中有数据区域最大的行数，法2
    max_row = sheet1.used_range.last_cell.row
    print('max row', max_row)
    # 工作表sheet中有数据区域最大的列数，法2
    max_col = sheet1.used_range.last_cell.column
    print('max col', max_col)

    time = sheet1.range((18, 1), (max_row, 1)).value
    goods_type = sheet1.range((18, 2), (max_row, 2)).value
    seller = sheet1.range((18, 3), (max_row, 3)).value
    goods = sheet1.range((18, 4), (max_row, 4)).value
    in_out = sheet1.range((18, 5), (max_row, 5)).value
    amount = sheet1.range((18, 6), (max_row, 6)).value

    from bill import BillInfo

    bills_list = []
    for i in range(len(time)):
        new_b = BillInfo(time=time[i], g_type=goods_type[i], seller=seller[i], account1='微信', goods=goods[i],
                         in_out=in_out[i], amount=amount[i])
        print('new bill:', new_b.to_str())
        bills_list.append(new_b)
    wb.close()
    return bills_list


if __name__ == '__main__':
    path = './微信支付账单(20221001-20221031).xlsx'
    wechat_bills = load_wechat_bills(path)
    if len(wechat_bills) > 0:
        print('bills_list', wechat_bills[0].to_str())
