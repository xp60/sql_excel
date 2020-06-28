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
def str_to_list(str):
    return str.replace('[','').replace(']','').split(',')

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
    # 接受参数
    date = sys.argv
    date.remove('run.py')
    print(date)
    for i in date:
        parameter_list.append(str_to_list(i))
    print(parameter_list)
    print(os.getcwd())
    ########

    sql_list = []
    os.chdir(r'sql')
    li  = os.listdir()
    for i in li:
        if i.split('.')[-1]=='sql':
            sql_list.append(i)
    print('sql_list is :',sql_list)
    parameter_index = 0
    for sql_file in sql_list:
        with open(sql_file, 'r') as f:
            sql = ''
            for line in f.readlines():
                try:
                    print(line == '')
                    
              
                    if line.strip().endswith(';') :
                        
                        sql += line.replace('\n', ' ')
                        print(sql)
                        # 传执行时的参数
                        if re.findall(r'{',sql):
                            print('into re========',parameter_list[parameter_index])
                            if (len(parameter_list[parameter_index]) < 2):
                                sql = sql.format(parameter_list[parameter_index]) 
                            else:
                                sql = sql.format(*parameter_list[parameter_index]) 
                            print(sql)
                        print(sql)
                        date_list.append(get_data(db,sql))
                        sql = ''
                    elif not line.strip().startswith('--'):
                        sql += line.replace('\n', ' ')
                        print(sql)
                    else:
                        pass
                except:
                    raise
            parameter_index += 1
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