# -*- coding: utf-8 -*-
import datetime
import os
import sys
import time

import xlwings as xw


# 账单数据条
class BillInfo(object):
    # ['时间', '分类', '类型', '金额', '账户1', '账户2', '备注', '账单图片', '交易对方 / 对方名称', '交易地点 / 商品名称']]
    def __init__(self, date_time=None, category=None, in_out=None, amount=None, account1=None, account2=None,
                 remarks=None, seller=None, goods=None):
        self.date_time = date_time  # 日期
        self.category = category  # 分类
        self.in_out = in_out  # 类型（支出/收入）
        self.amount = amount  # 金额
        self.account1 = account1  # 账户1
        self.account2 = account2  # 账户2
        self.remarks = remarks  # 备注
        self.seller = seller  # 交易对方
        self.goods = goods  # 商品名称

    def to_str(self):
        return str(self.date_time) + str(self.category) + str(self.in_out) + str(self.amount) + str(
            self.account1) + str(self.seller) + str(self.goods)


# 加载支付宝账单数据
def load_alipay_bills(xlsx_path):
    print('load_alipay_bills')
    wb = xw.Book(xlsx_path)
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

    # 构建账单数据条（行），不构建账单数据条也可以，在写入数据的时候按列写入上面的列表， 但在写入的时候要注意写入的顺序。
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


# 加载微信账单
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

    # 构建账单数据条（行）
    bills_list = []
    for i in range(len(date_time)):
        new_b = BillInfo(date_time=date_time[i], category=goods_type[i], seller=seller[i], account1='微信',
                         goods=goods[i], in_out=in_out[i], amount=amount[i])
        # print('new bill:', new_b.to_str())
        bills_list.append(new_b)
    print('load_wechat_bills done. cnt:', len(bills_list))
    return bills_list


class QianJiHelper(object):

    def __init__(self, xlsx_name):
        self.xlsx_name = xlsx_name  # 日期
        self.create_new_qianji_xlsx()

    def create_new_qianji_xlsx(self, ):
        xlsx_name = self.xlsx_name
        print('create_new_xlsx', xlsx_name)

        # 方法1：
        # 创建一个新的App，并在新App中新建一个Book
        wb = xw.Book()
        sheet1 = wb.sheets["sheet1"]
        sheet1.clear()
        sheet1.range('A1').value = [
            ['时间', '分类', '类型', '金额', '账户1', '账户2', '备注', '账单图片', '交易对方 / 对方名称',
             '交易地点 / 商品名称']]
        sheet1.range('A1').column_width = 16
        sheet1.range('I1').column_width = 20
        sheet1.range('J1').column_width = 20
        wb.save(xlsx_name)
        wb.close()
        print('create_new_xlsx done')

    def write_data(self, bills):
        print('write_data in excel')

        # 实例化一个工作表对象
        wb = xw.Book(self.xlsx_name)
        sheet1 = wb.sheets["sheet1"]
        max_row = sheet1.used_range.last_cell.row
        print('max row', max_row)
        # 逐行写入数据
        for i in range(len(bills)):
            bill = bills[i]
            # sheet1.range('A2').options(transpose=True).value = bill.time
            # 从第2行第一列 逐行写入
            sheet1.range(i + max_row + 1, 1).value = [bill.date_time, bill.category, bill.in_out, bill.amount,
                                                      bill.account1, bill.account2, bill.remarks, '', bill.seller,
                                                      bill.goods]
        wb.save()
        wb.close()
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
    print(bill_file)
    return bill_file


if __name__ == '__main__':
    start_time = time.time()
    bill_files = getfiles()
    if len(bill_files) is 0:
        print('当前目录下没有账单文件')
        sys.exit()
    # 以年月为文件名
    ISO_TIME_FORMAT = '%Y-%m-%d %H:%M:%S'
    theTime = datetime.datetime.now().strftime(ISO_TIME_FORMAT)
    output_name = theTime[:7] + '.xlsx'

    # 创建qianji 账单模板 excel
    qianji_helper = QianJiHelper(xlsx_name=output_name)

    for file in bill_files:
        if file.startswith('微信'):
            wechat_bills = load_wechat_bills(xlsx_path=file)
            qianji_helper.write_data(wechat_bills)
        if file.startswith('alipay'):
            alipy_bills = load_alipay_bills(xlsx_path=file)
            qianji_helper.write_data(alipy_bills)

    # todo: refine category
    print("耗时: {:.3f}秒".format(time.time() - start_time))
