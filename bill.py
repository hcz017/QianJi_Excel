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
