import os
import random
import pandas as pd
import re
import string
import sqlalchemy as sql
from win32com import client
import docx

def integrate(folder, all_txt):  # 将文件夹中txt合并
    name_list = os.listdir(folder)
    for name in name_list:
        file_path = folder + '/' + name
        with open(file_path, 'r', encoding='utf-8')as f:
            sentence_list = f.readlines()
        with open(all_txt, 'a+', encoding='utf-8')as w:
            for sentence in sentence_list:
                sentence = sentence.lstrip()
                w.writelines(sentence)
            w.write('\n')


def participle(file_name1, file_name2):  # 将分词后的句子转换为词 标签\n的形式
    sentence1_list = []
    with open(file_name1, 'r', encoding='utf-8')as f:
        sentence_list = f.readlines()
        for sentence in sentence_list:
            sentence1 = sentence.split()
            for i in sentence1:
                if i[-1] != '\n':
                    i = i + '\n'
                sentence1_list.append(i)
    with open(file_name2, 'w', encoding='utf-8')as w:
        for line in sentence1_list:
            w.writelines(line)


def sentence_label(file_name1, file_name2):  # 句子标签处理
    with open(file_name1, 'r', encoding='utf-8')as f, open(file_name2, 'a+', encoding='utf-8')as w:
        sentence_list = f.readlines()
        for sentence in sentence_list:
            if len(sentence) != 10:
                if len(sentence) != 10 and sentence != '\n':
                    sentence = sentence[4:-6] + '\t' + 'FD' + '\n'
                elif sentence[0:4] == '<FC>':
                    sentence = sentence[4:-6] + '\t' + 'FC' + '\n'
                elif sentence != '\n':
                    sentence = sentence.strip('\n') + '\t' + 'O' + '\n'
                w.writelines(sentence)
        w.write('\n')
        # for sentence in sentence_list:
        #     if len(sentence) != 10 and sentence != '\n':
        #         if sentence[0:4] == '<FD>':
        #             sentence = '1' + '\t' + sentence[4:-6].lstrip() + '\n'
        #         elif sentence[0:4] == '<FC>':
        #             sentence = '2' + '\t' + sentence[4:-6].lstrip() + '\n'
        #         elif sentence != '\n':
        #             sentence = '0' + '\t' + sentence
        #         w.writelines(sentence)


def devide(file_name1):  # 划分验证集训练集
    with open(file_name1, 'r', encoding='utf-8')as a:
        sentence_list = a.readlines()
    test_list = []
    a = len(sentence_list)
    num_list_test = [random.randint(0, a) for i in range(0, int(a/100))]
    for i in num_list_test:
        test_list.append(sentence_list[i])
    train_list = [sentence_list[i] for i in range(len(sentence_list)) if (i not in num_list_test)]
    with open('train.txt', 'w', encoding='utf-8')as tr:
        for line in train_list:
            tr.writelines(line)
    with open('test.txt', 'w', encoding='utf-8')as te:
        for line in test_list:
            te.writelines(line)


def devide_1(file_name1, file_name2, file_name3):  # 将按行划分的中英文本分为两个单语文本文档
    ch_list = []
    en_list = []
    name_list = os.listdir(file_name1)
    for name in name_list:
        path = file_name1 + '/' + name
        with open(path, 'r', encoding='utf-8') as raw:
            raw_list = raw.readlines()
            for i in range(len(raw_list)):
                if (raw_list[i][0] >= 'a' and raw_list[i][0] <= 'z') or (raw_list[i][0] >= 'A' and raw_list[i][0] <= 'Z') or raw_list[i][0] in [0, 9]:
                    en_list.append(raw_list[i])
                if raw_list[i][0] >= '\u4e00' and raw_list[i][0] <= '\u9fa5':
                    ch_list.append(raw_list[i])
                if raw_list[i] == '\n':
                    if (raw_list[i-2][0] >= 'a' and raw_list[i-2][0] <= 'z') or (raw_list[i-2][0] >= 'A' and raw_list[i-2][0] <= 'Z') or raw_list[i-2][0] in [0, 9]:
                        en_list.append(raw_list[i])
                    if raw_list[i-2][0] >= '\u4e00' and raw_list[i-2][0] <= '\u9fa5':
                        ch_list.append(raw_list[i])
    with open(file_name2, 'w', encoding="utf-8")as ch:
        for line in ch_list:
            ch.write(line)
    with open(file_name3, 'w', encoding="utf-8")as en:
        for line in en_list:
            en.write(line)


def devide_2(excel_name, out_txt):  # 从excel提取每列带固定标签的内容并写入txt
    data = pd.read_excel(excel_name)
    colum_list = data.columns.tolist()
    print(colum_list)
    for i in range(len(colum_list)):
        print(i)
        df = pd.read_excel(excel_name, usecols=[i])
        df_list = df.values.tolist()
        txt_list = []
        for line in df_list:
            if r'/t' in str(line) :
                txt_list.append(str(line).strip('[').strip(']').strip("'"))
        with open(out_txt, 'a+', encoding='utf-8')as f:
            for line in txt_list:
                f.write(line + '\n')


def devide_3(txt_name, txt_name1, txt_name2):
    with open(txt_name, 'r', encoding='utf-8')as f:
        word = f.read()
    line_list = word.split('\n\n')
    g_list = []
    b_list = []
    for i in range(len(line_list)):
        if i % 2 ==0:
            g_list.append(line_list[i])
        if i % 2 !=0:
            b_list.append(line_list[i])
    with open(txt_name1, 'w', encoding='utf-8')as g:
        for line in g_list:
            g.write(line + '\n\n')
    with open(txt_name2, 'w', encoding='utf-8')as b:
        for line in b_list:
            b.write(line + '\n\n')



def excel2txt(folder, txt_path):  # 将excel中分列写入excel
    result_list = []
    for name in os.listdir(folder):
        if name.split('.')[-1] != 'txt':
            name_path = folder + '/' + name
            data_ch = pd.read_excel(name_path, usecols=['中文'])
            data_en = pd.read_excel(name_path, usecols=['英文'])
            data_en_list = data_en.values.tolist()
            data_ch_list = data_ch.values.tolist()
            print(name)
            for i in range(len(data_ch_list)):
                if type(data_ch_list[i][0]) == str:
                    ch = data_ch_list[i][0].lstrip(string.digits).strip().strip('\n')
                if type(data_en_list[i][0]) == str:
                    en = data_en_list[i][0].strip('\n')
                with open(txt_path, 'a+', encoding='utf-8')as r:
                    r.write(ch + '\n')
                    r.write(en + '\n')


def txt_process(folder, out_folder):  # 修改txt文件内容格式
    for name in os.listdir(folder):
        name_path = folder + '/' + name
        with open(name_path, 'r', encoding='utf-8')as f:
            word_list = f.readlines()
        txt_path = out_folder + '/' + name
        with open(txt_path, 'w', encoding='utf-8')as s:
            for line in word_list:
                if line != '\n':
                    line = line.lstrip(string.digits).strip().strip('\t').replace('「', '“').replace('」', '”').replace('『', '‘').replace('』', '’')
                    if line != '':
                        s.write(line + '\n')


def seq_process(folder, out_folder):  # 将分词后的语料更改为字\t标签的格式
    result_list = []
    with open(folder, 'r', encoding='utf-8')as s:
        sentence_list = s.readlines()
    with open(out_folder, 'w', encoding='utf-8')as f:
        for sentence in sentence_list:
            word_list = sentence.split('/')
            for word in word_list:
                if word == '\n':
                    f.write(word)
                else:
                    if len(word) == 1:
                        f.write(word + '\t' + 'B-NN' + '\n')
                    elif len(word) > 1:
                        f.write(word[0] + '\t' + 'B-NN' + '\n')
                        for i in range(1, len(word)):
                            f.write(word[i] + '\t' + 'I-NN' + '\n')    


def sql_read():  # 读取数据库数据
    engine = sql.create_engine('mysql+pymysql://admin:748620@192.168.116.123:3306/term_disambiguation')
    sql1 = '''select * from cnki_allterms'''
    datas = pd.read_sql(sql1,engine)
    return datas


def doc_to_docx(word_name, save_file):  # doc转docx
    word = client.Dispatch("Word.Application")
    doc_path = 'F:\数据相关\读取word\docs\\' + word_name 
    docx_name = os.path.splitext(word_name)[0] + '.docx'
    out_path = save_file + '\\' + docx_name
    doc = word.Documents.Open(doc_path)
    doc.SaveAs("{}".format(out_path), 12, False, "", True, "", False, False, False, False)
    doc.Close()
    word.Quit()


def word_read(docx_path):  # 读取docx文档
    docx_text = docx.Document(docx_path)


if __name__ == '__main__':
    txt_process('1', '2')