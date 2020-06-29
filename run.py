import datetime,pymysql,time
import xlwt
from config import *
import os, re
import sys
from time import sleep



class Dict(dict):
    '''
    Simple dict but support access as x.y style.
    '''
    def __init__(self, names=(), values=(), **kw):
        super(Dict, self).__init__(**kw)
        for k, v in zip(names, values):
            self[k] = v

    def __getattr__(self, key):
        try:
            return self[key]
        except KeyError:
            raise AttributeError(r"'Dict' object has no attribute '%s'" % key)

    def __setattr__(self, key, value):
        self[key] = value


# 做成dict类型
def toDict(d):
    D = Dict()
    for k, v in d.items():
        D[k] = toDict(v) if isinstance(v, dict) else v
    return D


#定义个方法执行查询sql操作
def get_data(db,sql):
    cur = db.cursor()
    # 使用cursor()方法获取操作游标
    try:
        cur.execute(sql)   #执行sql语句
        return cur.fetchall()  #获取查询的所有记录
    except Exception as e:
        raise e
    finally:
        cur.close()


# str 转 list
def str_to_list(now_str):
    # 多个参数解析成list
    if re.findall(r'\(',now_str):
        cur_str = now_str.replace('[','').replace(']','').replace('(','|').replace(')','|').replace(' ','').split('|')
        # print(cur_str)
        str_list = []
        for temporary_str in cur_str:
            if temporary_str != '' and temporary_str!=',':
                str_list.append(temporary_str.split(','))
        return str_list
    else:
        return  list(map(lambda x: x.strip(), now_str.replace('[','').replace(']','').split(',')))

if __name__ == '__main__':
    ThisMonthToday=datetime.date.today()
    configs = toDict(configs)
    host=configs['db']['host']
    port=configs['db']['port']
    db_name=configs['db']['database']
    user=configs['db']['user']
    password=configs['db']['password']
    # 连接数据库
    db= pymysql.connect(host=host,user=user,password=password,db=db_name,port=port)
    date_list=[]
    parameter_list=[]
    # 拿取sql.txt参数
    with open('date.txt', 'r') as f:
        # 切换sql目录
        os.chdir(r'sql')
        for line in f.readlines():
            
            line = line.replace('\n', '')
            resut_list = re.split(r" +",line,1)
            sql_file = resut_list[0]
            parameter_list = str_to_list(resut_list[1])
            try:
                # sql_file == xxx.sql
                with open(sql_file, 'r') as f:
                    sql = ''
                    for run_sql_line in f.readlines():
                        # print(run_sql_line)
                        if run_sql_line.strip().endswith(';') :
                            # 准备执行sql
                            sql += run_sql_line.replace('\n', ' ')
                                # 传执行时的参数
                            if re.findall(r'{',sql):
                                if (isinstance(parameter_list[0], str)):
                                    for i in parameter_list:
                                        # print(sql.format(i))
                                        date_list.append(get_data(db,sql.format(i)))
                                else:
                                    for i in parameter_list:
                                        # print(sql.format(*i))
                                        date_list.append(get_data(db,sql.format(*i) ))
                            sql = ''
                        elif not run_sql_line.strip().startswith('--'):
                            sql += run_sql_line.replace('\n', ' ')
                        else:
                            pass
            except:
                raise
    db.close()
    # print(date_list)
    # 创建一个xls文件对象
    wb = xlwt.Workbook()
    # 加入表单
    sh = wb.add_sheet('Last_month')
    # 制作表头
    os.chdir(r'../')
    with open('title.txt', 'r') as f:
        for line in f.readlines():
            # print(str(line))
            try:
               title_list = re.split(r'[|]+', str(line).strip() )
               i = 0
            #    print(title_list)
               for title in title_list:
                   sh.write(0,i,title)
                   i += 1
                #    print(i)
            except:
                raise
    data_1_lenth=len(date_list)
    start_row_num=1
    for date_1_list in date_list:
        for item in date_1_list:
            start_col_num=0
            for date in item:
                sh.write(start_row_num,start_col_num,date)
                start_col_num+=1
            start_row_num+=1
    filename=str(ThisMonthToday)
    wb.save(filename+'报表'+'.xls')
    print('报表生成完成！！！')