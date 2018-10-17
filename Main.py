__author__ = 'Luzaofa'
__date__ = '2018/10/12 14:30'

import os, shutil
from datetime import datetime
import time

import ExcelHelper
import DB_Helper
import Config


class MainService(object):

    def __init__(self, file_path):
        self.excelhelper = ExcelHelper.ExcelHelper(file_path, False)
        self.dbhelper = DB_Helper.DB_helper()

    def get_file_type(self, fileName):
        """获取文件类型"""
        file_types = Config.get_file_type
        for key, values in file_types.items():
            for value in values:
                if value == fileName:
                    return key

    def main(self, fileName):
        """
            业务处理模块
        """
        print(f'正在处理：{fileName}')
        self.excelhelper.find_sheet(1)
        file_type = self.get_file_type(fileName)
        config = Config.file_type[file_type]

        judge_col = 1
        judge_value = ''

        if file_type == 'type1':
            judge_value = 'Luzaofa'

        if file_type == 'type2':
            judge_col = 2
            judge_value = '卢造发'

        col_values = self.excelhelper.get_col_values(judge_col)
        row = 1
        for col in col_values:
            try:
                if judge_value in str(col[0]):
                    mass = {}
                    for key, value in config.items():
                        values = self.excelhelper.get_cell_value(row, value)
                        mass[key] = values
                    print(mass)
                row += 1
            except:
                row += 1

        self.excelhelper.close()
        return True


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
    with open(log_path, 'a+') as f:
        f.write(value + '\n')


def kill_process():
    """关闭后台Excel所占进程"""
    command = 'taskkill /F /IM {process_name}'
    for i in ['EXCEL.EXE']:
        os.system(command.format(process_name=i))


if __name__ == '__main__':

    path = 'XXX/ExcelProject/Files/'  # 本地文件夹
    move_path = 'XXX/ExcelProject/DealFiles/'  # 处理后移除文件夹
    log_path = 'XXX/ExcelProject/log.txt'  # 日志

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
