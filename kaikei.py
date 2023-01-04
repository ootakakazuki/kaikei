import pandas as pd
import datetime
import os
import numpy as np
import re
import shutil


# 処理済みのファイル名
ALLFILES_COMPLETED = "allfiles_completed"

# 処理失敗したファイル名
PROCESS_FAILED_FILE_NAMES = "処理失敗ファイル"

ANS_0 = "合計金額がNULL または0です"

hidukeari = 0
hidukenashi = 0
dic = {
    "1月": [0, 0],
    "2月": [0, 0],
    "3月": [0, 0],
    "4月": [0, 0],
    "5月": [0, 0],
    "6月": [0, 0],
    "7月": [0, 0],
    "8月": [0, 0],
    "9月": [0, 0],
    "10月": [0, 0],
    "11月": [0, 0],
    "12月": [0, 0],
    "total_amount": [0, 0]
}

list_processfailed = []
list_reasons_for_fail = []

def filename_insert_totalamount(file, totalamount):
    p = r'(.*)\)'
    m = re.search(p, file)
    name1 = ""
    try:
        name1 = m.group(1)             
    except AttributeError as e:
        print("▲▼▲▼▲▼ファイル名に括弧がついていません▲▼▲▼▲▼", file=fo)
        pass
    p = r'\)(.*)'
    m = re.search(p, file)
    name2 = ""
    try:
        name2 = m.group(1)
    except AttributeError as e:
        pass
    if totalamount == 0:
        name1 = name1 + ")" + name2
    else:
        name1 = name1 + "_" + str(totalamount) + ")" + name2
    if name1 == "" or name2 == "":
        list_processfailed.append(file.title())
    os.rename(allfiles_copy + "/" + file, allfiles_copy + "/"+ name1)
    return name1

def dic_to_excel(dic):
    df2 = pd.DataFrame(dic)
    df2 = df2.transpose() # 転置
    df2 = df2.rename(columns={0: "金額合計", 1: "ファイル数"})
    print(df2.head(15), file=fo)
    #excel_writer = pd.ExcelWriter('seikyuusyo_data_per_month.xlsx')
    #df2.to_excel(excel_writer)
    df2.to_csv("done.csv")
    
def dic_calc_sum():
    total_amount = 0
    total_count = 0
    for key, val in dic.items():
        total_amount += dic[key][0]
        total_count += dic[key][1]
    dic["total_amount"][0] = total_amount
    dic["total_amount"][1] = total_count


def date_process(df, file, hidukeari, flag_shippai):
    month = 0
    flag_date = False
    for i in range(1, 5):
        if flag_date:break
        na_records = list(df.iloc[i, :].isnull())
        for j in range(df.shape[1]):
            if isinstance(df.iat[i, j], datetime.datetime) and na_records[j] == False :
            
                flag_date = True
                print(str(df.iat[i, j].month) + "月", file=fo)
                month = str(df.iat[i, j].month) + "月"
                break
        
    if flag_date == False:
        print("file名: " + file, file=fo)
        print("請求日が見つからなかった", file=fo)
        flag_shippai = True
    return month, hidukeari, flag_shippai


def amount_process(df, file, month, flag_shippai, new_filename):
    total = 0
    flag_dis = False
    for i in range(11, 14):
        if flag_dis: break
        na_records_goukeikingaku = list(df.iloc[i, :].isnull())
        for j in range(df.shape[1]):
            if df.iat[i, j] == "合計金額" or df.iat[i, j] == "合計金額(税込)" == False: continue
            if flag_dis: break
            for k in range(j, df.shape[1]):
                if isinstance(df.iat[i, k], int) and na_records_goukeikingaku[k] == False:
                    if df.iat[i, k] == 0:
                        print("合計金額が0円です", file=fo)
                        break
                    
                    flag_dis = True
                    print("合計金額: " + str(df.iat[i, k]), file=fo)
                    dic[month][0] += df.iat[i, k]
                    dic[month][1] += 1
                    total = df.iat[i, k]
                    #new_filename = filename_insert_totalamount(file, df.iat[i, k])
                    break            
        
    new_filename = filename_insert_totalamount(file, total)
                    
    if flag_dis == False:
        print(ANS_0, file=fo)  # 合計金額がNULL または0
        flag_shippai = True
        # list_processfailed.append(file.title())
    return flag_shippai, new_filename
                    
def print_filecount():
    print("ファイル数: " + str(len(files)), file=fo)
    print("-"*30, file=fo)

def ready_dir():
    if os.path.exists("allfiles_copy"): shutil.rmtree("allfiles_copy")
    taisyou_folder = input("処理対象となるエクセルファイルが入ったフォルダー名を教えてください:  ")
    #taisyou_folder = "allfiles"
    allfiles_copy = shutil.copytree(taisyou_folder, "allfiles_copy")
    files = os.listdir(allfiles_copy)
    #os.makedirs(ALLFILES_COMPLETED + "/1", ALLFILES_COMPLETED + "/2", exist_ok=True)
    list_month = ["1月", "2月", "3月", "4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月", PROCESS_FAILED_FILE_NAMES]

    for i in list_month:
        os.makedirs(ALLFILES_COMPLETED + "/" + i)
    return files, allfiles_copy

def hogehoge(hidukenashi):
    list_processfailed.append(file.title())
    hidukenashi += 1
    print("失敗　: " + str(hidukenashi) + "\n", file=fo)
    #print()
    return hidukenashi


def file_to_monthdirectory(month, new_filename, allfiles_copy):
    os.rename(allfiles_copy + "/" + new_filename, ALLFILES_COMPLETED + "/" + month + "/" + new_filename, )

def file_to_faileddirectory(file, allfiles_copy):
    os.rename(allfiles_copy + "/" + file.title(), ALLFILES_COMPLETED + "/" + PROCESS_FAILED_FILE_NAMES + "/" + file.title())


if __name__ == '__main__':
    #allfiles_copy = ""
    fo = open("log.txt", "w")
    if os.path.exists(ALLFILES_COMPLETED): 
        print("すでに" + ALLFILES_COMPLETED + "は存在しています", file=fo)
        exit()
    files, allfiles_copy = ready_dir()
    print_filecount()
    for file in files:
        new_filename = ""
        flag_shippai = False
        # ファイル名の先頭が「請求書」のみ処理する
        if file.title()[:3] != "請求書": 
            print(file.title() + "はタイトルが請求書ではないため、処理をスキップします", file=fo)
            hidukenashi = hogehoge(hidukenashi)
            #os.rename(allfiles_copy + "/" + file.title(), ALLFILES_COMPLETED + "/" + PROCESS_FAILED_FILE_NAMES + "/" + file.title() )
            file_to_faileddirectory(file, allfiles_copy)
            continue
        dirfile = allfiles_copy + "/" + file
        path = os.path.join(os.getcwd(), dirfile)
        print("処理中のファイル名:  " + file, file=fo)
        df = pd.read_excel(path)
        month, hidukeari, flag_shippai = date_process(df, file, hidukeari, flag_shippai)
        if month == 0:
            hidukenashi = hogehoge(hidukenashi)
            file_to_faileddirectory(file, allfiles_copy)
            continue
        flag_shippai, new_filename = amount_process(df, file, month, flag_shippai, new_filename)
        
        if flag_shippai:
            hidukenashi = hogehoge(hidukenashi)
            file_to_faileddirectory(file, allfiles_copy)
            
        else:
            hidukeari += 1
            print("成功　: " + str(hidukeari) + "\n", file=fo)
            #print()
            file_to_monthdirectory(month, new_filename, allfiles_copy)

    dic_calc_sum()
    shutil.rmtree(allfiles_copy)
    print("処理成功: " + str(hidukeari), file=fo)
    print("処理失敗: " + str(hidukenashi), file=fo)
    print("-"*30, file=fo)
    print("【処理失敗ファイル一覧】", file=fo)
    for i in list_processfailed:
        print(i, file=fo)
    print("-"*30, file=fo)
    dic_to_excel(dic)
    fo.close()
