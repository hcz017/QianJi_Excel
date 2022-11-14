import xlwings as xw


def create_new_xlsx(xlsx_name):
    print('create_new_xlsx')
    import os
    if os.path.exists(xlsx_name) is True:
        return
    # 方法1：
    # 创建一个新的App，并在新App中新建一个Book
    wb = xw.Book()
    sheet1 = wb.sheets["sheet1"]
    sheet1.range('A1').value = [
        ['时间', '分类', '类型', '金额', '账户1', '账户2', '备注', '账单图片', '交易对方 / 对方名称', '交易地点 / 商品名称']]
    sheet1.range('A1').column_width = 16
    sheet1.range('I1').column_width = 20
    sheet1.range('J1').column_width = 20
    wb.save(xlsx_name)
    wb.close()


def write_data(path, bills):
    print('write_data in excel')

    import os
    if os.path.exists(path) is False:
        create_new_xlsx(path)
    # 在A1单元格写入值
    # 实例化一个工作表对象
    wb = xw.Book(path)
    sheet1 = wb.sheets["sheet1"]
    # 输出工作簿名称
    print(sheet1.name)
    # 写入值
    # sheet1.range('A1').value = 'python知识学堂'
    # 读值并打印
    print('value of A1:', sheet1.range('A1').value)
    # 逐行写入数据
    for i in range(len(bills)):
        bill = bills[i]
        # sheet1.range('A2').options(transpose=True).value = bill.time
        # 从第2行第一列 逐行写入
        sheet1.range(i + 2, 1).value = [bill.time, bill.goods_type, bill.in_out, bill.amount, bill.account1,
                                        bill.account2, bill.remarks, '', bill.seller, bill.goods]
    wb.save()
    wb.close()


if __name__ == '__main__':
    name = '2022-11.xlsx'
    # create_new_xlsx(name)
    write_data(name, None)
