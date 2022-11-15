# -*- coding: utf-8 -*-
import datetime
import os

import qianji_helps
import read_alipay
import read_wechat


def getfiles():
    filenames = os.listdir(r'./')
    bill_file = []
    # 只返回扩展名是 csv 和 xlsx 的文件
    for file in filenames:
        if str(file).endswith('csv'):
            bill_file.append(file)
        if str(file).endswith('xlsx'):
            bill_file.append(file)
    print(bill_file)
    return bill_file


if __name__ == '__main__':
    # 以年月为文件名
    ISOTIMEFORMAT = '%Y-%m-%d %H:%M:%S'
    theTime = datetime.datetime.now().strftime(ISOTIMEFORMAT)
    output_name = theTime[:7] + '.xlsx'

    # create new  qianjiexcel
    qianji_helps.create_new_xlsx(xlsx_name=output_name)

    bill_files = getfiles()
    for file in bill_files:
        if file.startswith('微信'):
            wechat_bills = read_wechat.load_wechat_bills(xlsx_path=file)
            qianji_helps.write_data(output_name, wechat_bills)
        if file.startswith('alipay'):
            alipy_bills = read_alipay.load_alipay_bills(xlsx_path=file)
            qianji_helps.write_data(output_name, alipy_bills)

    # todo: refine type
