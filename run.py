from multiprocessing import Pool
from time import sleep
from config import *
import datetime,pymysql,time
import os, re
import sys
import xlwt


import xlrd
import pandas as pd
from pathlib import Path


max_process = 10


def read_SQL_select(configs, sql,sql_title, p_count):
    print(sql)
    host=configs['db']['host']
    port=configs['db']['port']
    db_name=configs['db']['database']
    user=configs['db']['user']
    password=configs['db']['password']
    # 连接数据库
    db= pymysql.connect(host=host,user=user,password=password,db=db_name,port=port)
    cur = db.cursor()
    # 使用cursor()方法获取操作游标
    try:
        cur.execute(sql)   #执行sql语句
        date_list = cur.fetchall()  #获取查询的所有记录
    except:
        print('sql error---------',sql)
        raise
    finally:
        cur.close()
        db.close()
    # print(date_list)
    # 做表
    wb = xlwt.Workbook()
    # 加入表单
    sh = wb.add_sheet('date')
    print(sql_title)
    # 制作表头
    with open('title.txt', 'r') as f:
        for line in f.readlines():
            re_list = re.split(r"\s+",line,1)
            # print(str(line))
            if re_list[0] == sql_title:
                try:
                    title_list = re.split(r'[|]+', str(re_list[1]).strip() )
                    i = 0
                    # print(title_list)
                    for title in title_list:
                        sh.write(0,i,title)
                        i += 1
                        #    print(i)
                except:
                    raise
    data_1_lenth=len(date_list)
    start_row_num=1
    for date_1_list in date_list:
        start_col_num=0
        for item in date_1_list:
            sh.write(start_row_num,start_col_num,item)
            start_col_num+=1
        start_row_num+=1
    filename=str(ThisMonthToday)
    wb.save(sql_title+'-'+ str(p_count) + '-' +filename+'报表'+'.xls')
    print(sql_title + '报表生成完成！！！')


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
# 获取xls表
def file_name(file_dir):
    list=[]
    for file in os.listdir(file_dir):
        if os.path.splitext(file)[1] == '.xls' or os.path.splitext(file)[1] == '.xlsx':
            list.append(file)
    return list

def merge_xlsx(path,filenames,sheet_num,output_filename):
    data = []   #定义一个空list
    title = []
    path_folder = Path(path)
    for i in range(len(filenames)):
        read_xlsx = xlrd.open_workbook(path_folder / filenames[i])
        sheet_num_data = read_xlsx.sheets()[sheet_num] #查看指定sheet_num的数据
        title = sheet_num_data.row_values(0)   #表头
        for j in range(1,sheet_num_data.nrows): #逐行打印
            data.append(sheet_num_data.row_values(j))
    content= pd.DataFrame(data)
    #修改表头
    content.columns = title
    #写入excel文件
    output_path = path_folder / 'output'
    output_filename_xlsx = output_filename + '-'+ str(ThisMonthToday) + '.xlsx'
    if not os.path.exists(output_path):
        print("output folder not exist, create it")
        os.mkdir(output_path)
    content.to_excel((output_path / output_filename_xlsx), header=True, index=False)
    print("merge success")


if __name__ == '__main__':
    ThisMonthToday=datetime.date.today()
    # sql配置信息
    configs = toDict(configs)
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
            parameter_list_len = 0
            if not isinstance(parameter_list[0], str):
                parameter_list_len = len(parameter_list) 
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
                                p_count = 1
                                p = None
                                if (parameter_list_len == 0 ):
                                    for i in parameter_list:
                                        print(sql.format(i))
                                        if p_count % max_process == 1:
                                            print(i)
                                            if p:
                                                p.close()
                                                p.join()
                                            p = Pool(max_process)
                                        p.apply_async(func=read_SQL_select, args=(configs,sql.format(i),sql_file.split('.')[0], p_count))
                                        # read_SQL_select(configs,sql.format(i),sql_file.split('.')[0]+str(p_count))
                                        # date_list.append(get_data(db,))
                                        p_count += 1
                                    if p:
                                        p.close()
                                        p.join()
                                else:
                                    for i in parameter_list:
                                        print(sql.format(*i))
                                        if p_count % max_process == 1:
                                            if p:
                                                p.close()
                                                p.join()
                                            p = Pool(max_process)
                                        p.apply_async(func=read_SQL_select, args=(configs,sql.format(*i),sql_file.split('.')[0], p_count))
                                        # read_SQL_select(configs,sql.format(*i),sql_file.split('.')[0]+str(p_count))
                                        # date_list.append(get_data(db,))
                                        p_count += 1
                                    if p:
                                        p.close()
                                        p.join()
                                    # p_count = 1
                            sql = ''
                        elif not run_sql_line.strip().startswith('--'):
                            sql += run_sql_line.replace('\n', ' ')
                        else:
                            pass
                    # print(sql_file.split('.')[0])
                    # create_excel(date_list,sql_file.split('.')[0])
                    # date_list.clear()
            except:
                # raise
                raise
    # db.close()
    # print(date_list)
    # 创建一个xls文件对象


    os.chdir(r'../')
    print(file_name('sql'))
    filenames = file_name('sql')
    k_file = {}
    for name in filenames:
        if name.split('-')[0] not in k_file:
            k_file[name.split('-')[0]] = [name]
        else:
            k_file[name.split('-')[0]].append(name)
    print(k_file)
    for k, v in k_file.items():
        print(k,v)
        merge_xlsx('sql',v,0,k) #合并文件中第一个表的数据，输出到 output/sheet1.xlsx中
    # merge_xlsx(path,filenames,1,"sheet1") #合并文件中第二个表的数据，输出到 output/sheet2.xlsx中
