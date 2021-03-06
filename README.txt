"""
   功能说明：
   该程序主要功能有：自动化解析Excel文件（不同样式类型），生成日志文件，文件解析完成后移动文件到move_path
   提供如下两种处理思路
"""

配置说明：
    path = 'XXX/Files/'           # 文件夹路径（存放待解析的文件）
    move_path = 'XXX/DealFiles/'  # 处理后移除文件夹（处理后文件存放处）
    log_path = 'XXX/log.txt'      # 日志（日志文件）

    思路一：将所处理类型配置到Config.py

    Config.py：Excel模板解析配置说明

    get_file_type：文件类型归类（同一类按照字典样式配置）
    file_type：不同类型所需解析字段（key:列名 value:对应所在列值）
    ExcelHelper.py：封装有关Excel的所有操作（实现其他功能只需调用相应的功能即可）

    思路二：将所处理类型配置到数据库

    FileType表：
        FileType 文件类型（可自定义）
        FileName 文件名字
    Config表：
      FileType     文件类型（与FileType表对应）
      ....         (自定义所需字段)
      JudgeCol     判断列
      JudgeValue   判断列所对应值
      TJC1         时间所在列
      TJC2         日期判断列
      TJC3         日期所在值
    Mass表：
      ....         (与Config表....保持一致)