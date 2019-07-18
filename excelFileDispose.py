import sqlite3
import xlrd

'''
  创建excel文件解析类
  由于不同的excel文件插入的字段和表不同，所以插入这部分通过回调函数的方式执行
'''


class ExcelFileDispose(object):
    def __init__(self, file):
        super(ExcelFileDispose, self).__init__()
        # 初始化数据库
        self.conn = sqlite3.connect(file)
        self.cursor = self.conn.cursor()

  # 释放数据库实例
    def __del__(self):
        self.cursor.close()
        self.conn.close()

  # 读取excel 文件
    def readFile(self, file, sheet_name, insertfunc):
        data = xlrd.open_workbook(file)
        table = data.sheet_by_name(sheet_name)
        for i in range(1, table.nrows):
            row_data = table.row_values(i)
            if row_data:
                insertfunc(self, i, row_data)
            else:
                print('row_data is null')
