__author__ = 'Luzaofa'
__date__ = '2018/10/12 13:57'

import pymssql


class DB_helper(object):

    def __init__(self):

        self.conn = pymssql.connect(
            host='192.168.2.12',
            user='sa',
            password='cytz@6666',
            database='datacenter2',
            charset='utf8')

        self.cursor = self.conn.cursor()

    def commit_sql(self, sql):
        try:
            self.cursor.execute(sql)
            self.conn.commit()
        except Exception as e:
            print(str(e))
            self.conn.rollback()  # 数据回滚，若一个插入失败都不做插入

    """
        自定义SQL
    """

    def insert(self):
        """数据插入"""
        sql = """"""
        self.commit_sql(sql)
