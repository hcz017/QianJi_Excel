class BillInfo(object):
    def __init__(self, time=None, g_type=None, seller=None, account1=None, account2=None, remarks=None, goods=None,
                 in_out=None, amount=None):
        self.time = time  # 日期
        self.goods_type = g_type  # 分类
        self.in_out = in_out  # 类型（支出/收入）
        self.amount = amount  # 金额
        self.account1 = account1  # 账户1
        self.account2 = account2  # 账户2
        self.remarks = remarks  # 备注
        self.seller = seller  # 交易对方
        self.goods = goods  # 商品名称

    def to_str(self):
        return str(self.time) + str(self.goods_type) + str(self.seller) + str(self.goods) + str(self.in_out) + str(
            self.amount)
