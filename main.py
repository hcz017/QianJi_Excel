import qianji_helps
import read_wechat

if __name__ == '__main__':
    name = '2022-11.xlsx'
    # create new qianji excel
    qianji_helps.create_new_xlsx(xlsx_name=name)
    # read wechat bills
    wechat_bills = read_wechat.load_wechat_bills(xlsx_path='./微信支付账单(20221001-20221031).xlsx')
    # write to qianji excel
    qianji_helps.write_data(name, wechat_bills)

    # todo: refine type