import qianji_helps
import read_wechat
import read_alipay

if __name__ == '__main__':
    name = '2022-11.xlsx'
    # create new qianji excel
    qianji_helps.create_new_xlsx(xlsx_name=name)
    # read wechat bills
    wechat_bills = read_wechat.load_wechat_bills(xlsx_path='./微信支付账单(20221001-20221031).csv')
    print('wechat bills count', len(wechat_bills))
    # write to qianji excel
    qianji_helps.write_data(name, wechat_bills)

    alipy_bills = read_alipay.load_alipay_bills(xlsx_path='./alipay_record_20221113_1558_1.csv')
    print('alipay bills count', len(alipy_bills))
    # write to qianji excel
    qianji_helps.write_data(name, alipy_bills)

    # todo: refine type
