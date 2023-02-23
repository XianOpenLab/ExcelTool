import csv
import sys
import os
import chardet


def check_encoding(filepath):
    with open(filepath, 'rb') as input_file:
        raw_data = input_file.read()
        result = chardet.detect(raw_data)
        if result['encoding'] == 'utf-8':
            return True
        else:
            return False


def read_csv(file):
    # 声明一个空列表用来存放数据
    data_list = []

    # 打开csv文件
    if check_encoding(file):
        encoding = 'utf-8'
    else:
        encoding = 'gbk '
    with open(file, 'r', encoding=encoding, errors="ignore") as csv_file:
        # 逐行读取csv文件内容
        try:
            csv_reader = csv.reader(csv_file)
            for row in csv_reader:
                data_list.append(row)
        except Exception as e:
            print(e)
            print("《" + os.path.basename(file) + "》格式异常，请用wps或者office另存份文件，再试")


    # 输出读取的CSV文件内容
    csv_file.close()
    return data_list
