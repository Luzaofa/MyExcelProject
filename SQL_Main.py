__author__ = 'Luzaofa'
__date__ = '2018/10/12 14:30'

import os, shutil
from datetime import datetime
import time

import ExcelHelper
import DB_Helper


class MainService(object):

    def __init__(self, file_path):
        self.excelhelper = ExcelHelper.ExcelHelper(file_path, False)
        self.dbhelper = DB_Helper.DB_helper()

    def main(self, fileName):
        """
            业务处理模块
        """
        print(f'正在处理：{fileName}')
        self.excelhelper.find_sheet(1)
        try:
            ConfigValue, Judge, Total = self.dbhelper.select(fileName)

            judge_col = Judge['JudgeCol']
            judge_value = Judge['JudgeValue']

            col_values = self.excelhelper.get_col_values(judge_col)
            row, T_row = 1, 1
            param, T_part = [], {}

            for T_col in col_values:
                """获取时间、总合计"""
                try:
                    if Total['TJC1'] != 0:
                        Time = self.excelhelper.get_cell_value(T_row, Total['TJC1'])
                        if '日期' in str(Time):
                            T_part['Time'] = Time.replace('日期：', '')
                    if Total['TJC2'] != 0:
                        TotalJudge = self.excelhelper.get_cell_value(T_row, Total['TJC2'])
                        if '基金小计' in str(TotalJudge):
                            TotalValue = self.excelhelper.get_cell_value(T_row, Total['TJC3'])
                            T_part['TotalValue'] = TotalValue
                            break
                        T_row += 1
                except:
                    T_row += 1

            print(T_part)
            print('======================================')

            for col in col_values:
                """获取正常值"""
                try:
                    if judge_value in str(col[0]):
                        mass, part = {}, []
                        for key, value in ConfigValue.items():
                            if value == 0:
                                values = 0
                                if key == 'Data' and 'Time' in T_part.keys():
                                    values = T_part['Time']
                                elif key == 'Total' and 'TotalValue' in T_part.keys():
                                    values = T_part['TotalValue']
                                mass[key] = values
                            else:
                                values = self.excelhelper.get_cell_value(row, value)
                                if ',' in str(values):
                                    values = values.replace(',', '')
                                mass[key] = values
                            part.append(values)
                        if self.dbhelper.select_mass(*part):
                            self.dbhelper.del_mass(*part)
                        param.append(part)
                        print(mass)
                    row += 1
                except:
                    row += 1

            # try:
            #     # self.dbhelper.create_table_config()
            #     # self.dbhelper.create_table_filetype()
            #     self.dbhelper.create_table_mass()
            #
            # except:
            #     pass

            # 批量插入数据（一篇一篇）
            sql = self.dbhelper.insert_mass()
            self.dbhelper.batch_insert(sql, param)

            self.excelhelper.close()
            return True

        except IOError as e:
            print(str(e), '没有该文件配置')
            return False


def load_files(path):
    """加载Excel文件"""
    files = os.listdir(path)  # 返回子目录下所有文件名集合
    for file in files:
        if file.endswith('xls') or file.endswith('xlsx') and not file.startswith('~'):
            yield path + file


def move_file(path, file, move_path):
    """处理后移除文件"""
    try:
        shutil.move(file, move_path)
    except:
        os.remove(move_path + file.replace(path, ''))
        shutil.move(file, move_path)


def log(value, log_path):
    """保存日志"""
    with open(log_path, 'a+') as f:
        f.write(value + '\n')


def kill_process():
    """关闭后台Excel所占进程"""
    command = 'taskkill /F /IM {process_name}'
    for i in ['EXCEL.EXE']:
        os.system(command.format(process_name=i))


if __name__ == '__main__':

    # path = "//127.0.0.1/Data/"  # 服务器文件夹路径（正式）
    # deal_path = '//127.0.0.1/DealFiles/'  # 处理后移除文件夹
    # log_path = '//127.0.0.1/DealFiles/'    # 日志
    path = 'E:/MyExcelProject/Files/'  # 本地文件夹（测试）
    move_path = 'E:/MyExcelProject/DealFiles/'  # 处理后移除文件夹
    log_path = 'E:/MyExcelProject/log.txt'  # 日志

    while True:
        deal_file = []
        for file in load_files(path):
            filename = file.split(' ')[1]
            mainService = MainService(file)
            mass = mainService.main(fileName=filename.replace('.xlsx', '').replace('.xls', ''))
            kill_process()
            if mass:
                deal_file.append(file.replace(path, ''))
                move_file(path, file, move_path)
        now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')  # 当前时间
        sleep_time = 5
        LogMass = f'处理完毕，{sleep_time}分钟后获取新数据！当前时间：{now}，一共处理了 {len(deal_file)} 个文件：{deal_file}'
        print(LogMass)
        log(LogMass, log_path)
        time.sleep(sleep_time * 60)
