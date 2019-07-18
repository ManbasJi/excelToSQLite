from excelFileDispose import ExcelFileDispose


# 自定义回调函数
def insertData(self, i, row_data):
    sql = 'insert into STOCK_DETAIL(ID,CODE,NAME,SCATE_NAME,CATEGORY_NAME,AREA_NAME,TOTAL_MARKET_VALUE)values(?,?,?,?,?,?,?)'
    self.cursor.execute(
        sql, (i, row_data[0], row_data[1], '沪深', row_data[6], row_data[7], row_data[15]))
    self.conn.commit()


# 执行分析插入语句
efd = ExcelFileDispose("./stock_market.db")
efd.readFile('./沪深Ａ股20190717.xlsx', 'detail', insertData)
