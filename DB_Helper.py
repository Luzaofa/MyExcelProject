__author__ = 'Luzaofa'
__date__ = '2018/10/12 13:57'

import pymssql
import sqlite3


class DB_helper(object):

    def __init__(self):

        # self.conn = pymssql.connect(
        #     host='host',
        #     user='user',
        #     password='password',
        #     database='database',
        #     charset='utf8')
        self.conn = sqlite3.connect('E:/MyExcelProject/MyExcelProject.db')

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

    def create_table_config(self):
        sql = """
            create table Config(
            FileType int ,
            Organization int ,
            Data int ,
            FundCode int ,
            FundName int ,
            UnitValue int ,
            Total int ,
            JudgeCol int ,
            JudgeValue int,
            TJC1 int ,
            TJC2 int,
            TJC3 int
            )"""
        self.commit_sql(sql)

    def create_table_filetype(self):
        sql = """
            create table FileType(
            FileType CHAR(10),
            FileName CHAR(50)
            )"""
        self.commit_sql(sql)

    def create_table_mass(self):
        sql = """
            create table Mass(
            Organization CHAR(50),
            Data CHAR(50),
            FundCode CHAR(50),
            FundName CHAR(50),
            UnitValue CHAR(50),
            Total CHAR(50)
            )"""
        self.commit_sql(sql)

    def select(self, fileName):
        sql = """
        select Organization, Data, FundCode, FundName, UnitValue, Total, JudgeCol, JudgeValue, TJC1, TJC2, TJC3 from Config
        where FileType = (select FileType from FileType where FileName = '%s')
        """ % (fileName)
        self.cursor.execute(sql)
        values = self.cursor.fetchall()
        Key = ['Organization', 'Data', 'FundCode', 'FundName', 'UnitValue', 'Total', 'JudgeCol', 'JudgeValue', 'TJC1',
               'TJC2', 'TJC3']
        dic, judge, total = {}, {}, {}
        i = 0
        for value in values[0]:
            if i < len(values[0][:-5]):
                dic[Key[i]] = value
            elif i >= len(values[0][:-5]) and i < len(values[0][:-3]):
                judge[Key[i]] = value
            else:
                total[Key[i]] = value
            i += 1
        return dic, judge, total

    def batch_insert(self, sql, param):  # 批量导入，sql为插入语句, param为插入值list
        try:
            self.cursor.executemany(sql, param)  # 批量执行
            self.conn.commit()
        except Exception as e:
            print(e)
            self.conn.rollback()  # 数据回滚，若一个插入失败都不做插入

    def insert_mass(self):
        sql = "insert into Mass(Organization, Data, FundCode, FundName, UnitValue, Total) values (?, ?, ?, ?, ?, ?)"
        return sql

    def select_mass(self, *param):
        sql = """
        select * from Mass
        where Organization = '%s' and Data = '%s' and FundCode = '%s' and FundName = '%s' and UnitValue = '%s' and Total = '%s'
        """ % (param)
        self.cursor.execute(sql)
        values = self.cursor.fetchall()
        return values

    def del_mass(self, *param):
        sql = """
        delete
        from Mass
        where Organization = '%s' and Data = '%s' and FundCode = '%s' and FundName = '%s' and UnitValue = '%s' and Total = '%s'
        """ % (param)
        self.commit_sql(sql)
