import datetime,pymysql,time
import xlwt
from config import *
import os, re



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
    with open('sql.txt', 'r') as f:
        sql = ''
        for line in f.readlines():
            try:
                # print(line.strip())
                if line.strip().endswith(';'):
                    sql += line.replace('\n', ' ')
                    # print(sql)
                    data_1=get_data(db,sql)
                    date_list.append(data_1)
                    sql = ''
                    # print('=====',data_1)
                else:
                    sql += line.replace('\n', ' ')
                    print(sql)
                

                
            except:
                raise
    db.close()
    # print(date_list)
    # 创建一个xls文件对象
    wb = xlwt.Workbook()
    # 加入表单
    sh = wb.add_sheet('Last_month')
    # 制作表头
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