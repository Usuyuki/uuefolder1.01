#起動メッセージ
print("========== ========== ========== ========== ========== ==========")
print("  宇大工学部生のためのフォルダ生成ツール(uuefolder)  ver1.01")
print("========== ========== ========== ========== ========== ==========")
print("Copyright @ 2020 Naofumi Obata. All rights reserved.")
print(" ")
print("ご利用いただき、ありがとうございます。")


import os
import pandas as pd
from time import sleep
import calendar
import datetime

#pathのための変数
slash ="/"
pro = "."

#月の最終曜日取得関数　　　　　→参考https://note.nkmk.me/python-datetime-first-last-date-last-week/
def get_day_of_last_week(year, month, dow): #day of the week =dow = 曜日
    '''dow: Monday(0) - Sunday(6)''' #月0,火1....日6
    n = calendar.monthrange(year, month)[1]
    l = range(n - 6, n + 1)
    w = calendar.weekday(year, month, l[0])
    w_l = [i % 7 for i in range(w, w + 7)]
    return l[w_l.index(dow)]

#日付リスト取得関数
def get_day_of_for_first_week(tuki_last,zenki_month):
    hizuke_all = []
    for tukihazime in range(tuki_last,1,-7):
        if tukihazime < 10:
            tukihazime = "0" + str(tukihazime)
        else:
            print("*")
        #0428など日付の前に
        month_with_hizuke = '0' + str(zenki_month) + str(tukihazime)
        hizuke_all.append(month_with_hizuke)
    return hizuke_all

#授業コマフォルダを作るための関数　　　　　→参考https://note.nkmk.me/python-listdir-isfile-isdir/
def coma_make(daylist,bango,path3):
    day_perfect = []
    print('現在のリストは')
    print(daylist)
    for daynow in daylist:
        bango += 1
        daynow =str(daynow)
        daynow =daynow.replace('\n', '')#エクセル内の改行はパスだとエラー出るので除く
        if daynow == "nan" or daynow == " ":
            print("空きコマ除去フィルター発動！")
        else:
            path4 =str(path3) + str(slash) + str(bango) +str(pro) + str(daynow)
            os.mkdir(path4)
            day_perfect.append(path4)
            print("授業ファイルの生成に成功しました。")
    return day_perfect

#日付フォルダを作るための関数
def Hizuke_folder_maker(day2_perfect,day2_all):
    for perfect_current in day2_perfect:
        for day2_current in day2_all:
                path7 = str(perfect_current) + slash + day2_current
                os.mkdir(path7)            




#もうすでにファイル作っていたら実行しない
maked = os.path.exists("1年前期")
if maked == True:
    print("すでに「1年前期」ファイルが生成されています。5秒後に終了します。")
    sleep(5)
    exit()
else:
    sleep(2)
    
    




print("1年前期の日付データを生成し、Excelデータを読み込みます。しばらくお待ち下さい。")
sleep(5)


"""
宇都宮大学の前期は4/20から8/7まで

曜日ごとの日付取得ゾーン↓

"""

#変数設定
first_month =4
first2_month = first_month + 1
last_month = 8
last2_manth = last_month 
year = 2020
#日付リスト生成(４月分は予め入力)
mon_all=["0420","0427"]
tue_all=["0421","0428"]
wed_all=["0422","0429"]
thi_all=["0423","0430"]
fri_all=["0424",]
#1年前期メイン処理
for zenki_month in range(first2_month,last2_manth,1):
    #月
    mon_last = get_day_of_last_week(year,zenki_month,0)#最終曜日の日付取得関数へ
    hizuke_all = get_day_of_for_first_week(mon_last,zenki_month)#日付リスト生成関数へ
    mon_all.extend(hizuke_all)#各曜日の日付リストに結合
    #日付フォルダ生成
    
    #火
    tue_last = get_day_of_last_week(year,zenki_month,1)
    hizuke_all = get_day_of_for_first_week(tue_last,zenki_month)#日付リスト生成関数へ
    tue_all.extend(hizuke_all)
    
    #水
    wed_last = get_day_of_last_week(year,zenki_month,2)
    hizuke_all = get_day_of_for_first_week(wed_last,zenki_month)#日付リスト生成関数へ
    wed_all.extend(hizuke_all)
    
    #木
    thi_last = get_day_of_last_week(year,zenki_month,3)
    hizuke_all = get_day_of_for_first_week(thi_last,zenki_month)#日付リスト生成関数へ
    thi_all.extend(hizuke_all)
    
    #金
    fri_last = get_day_of_last_week(year,zenki_month,4)
    hizuke_all = get_day_of_for_first_week(fri_last,zenki_month)#日付リスト生成関数へ
    fri_all.extend(hizuke_all)
    
    
#８月特別処理(8/1～7)
mon_all.append("0803")
tue_all.append("0804")
wed_all.append("0805")
thi_all.append("0806")
fri_all.append("0807")

print("日付リストの生成に成功しました。")
    


"""
メインプロジェクト↓

"""

#ファイル名の取得
path = "./ここに時間割Excelを入れてください"
file = os.listdir(path)

#ファイルが２つ以上ある場合は警告文
file2 = len(file)
if file2 >= 2:
    print('フォルダに２つ以上ファイルが入っているかExcelファイルが開かれてます。\nファイルを1つにする、もしくはExcelファイルを閉じてから、もう１度起動してください。')
    sleep(6)
    exit()
elif file2 == 0:
    print("フォルダにファイルがありません。\n先に時間割エクセルデータを入れてから再度起動してください。")
    sleep(6)
    exit()
else:
    refile = file[0]

    #Excelデータでなければ警告文
    if 'xlsx' not in refile:
        print('Excelデータを入れてください。\nPDFなど他のファイルは読み込めません。')
        sleep(6)
        exit()
    else:
        #エクセルの読み込み
        path2 = "./ここに時間割Excelを入れてください/" + refile
        #シート名は選択してもらう
        print( str(file) + "を読み込みました。")
        sleep(2)
        print("準備が完了しました。")
        sleep(2)
        print("次にExcel内のシートを読み込みます。")
        sleep(5)
        print("----------------------------------------------------------------------------")
        print("あなたのクラスを上記の番号の中から選んで数字を入力して下さい。")
        print("[1:Ⅰ-1],[2:Ⅰ-2],[3:Ⅱ-1],[4:Ⅱ-2],[5:Ⅲ-1],[6:Ⅲ-2],[7:Ⅳ-1],[8:Ⅳ-1]")
        print("----------------------------------------------------------------------------")
        sleep(2)
        print("例えばⅡ-1クラスの人は3と入力")
        un =input("ここに半角で数字のみを入力してください:") #選んだ番号を入れる変数usernumberの略
        un = int(un)
        if un == 1:
            sn ="Ⅰ-1"    #snはSheetNumberの略
        elif un == 2:
            sn ="Ⅰ-2"
        elif un == 3:
            sn ="Ⅱ-1"
        elif un == 4:
            sn ="Ⅱ-2"
        elif un == 5:
            sn ="Ⅲ-1"
        elif un == 6:
            sn ="Ⅲ-2"
        elif un == 7:
            sn ="Ⅳ-1"
        elif un == 8:
            sn ="Ⅳ-1"
        else:
            print("正しい数値(1～9の半角整数)を入力してください。")
            print("システムの都合上10秒後に終了します。")
            print("お手数をおかけ致しますが、もう１度最初からお願います。")
            sleep(10)
            exit()
        print(sn + "ですね。入力ありがとうございます！")
        
        #Exelシートを開く
        excelfile = pd.read_excel(path2, sheet_name=sn)
        
        #Excelの行や列が変動してないか調べる。
        check_co = excelfile.iloc[1,23] #[行→、列↓]
        check_ro = excelfile.iloc[10,0]
        if str(check_co) == "時間割ｺｰﾄﾞ":
            print("Excel内の行の変更がされてないことを確認しました。")
            print()
        else:
            print("Excel内の行が変更されています。\n大変申し訳ありませんが正しく読み込めません。")
            sleep(10)
            exit()
        if str(check_ro) == "9-10時限":
            print("Excel内の列の変更がされてないことを確認しました。")
            print()
        else:
            print("Excel内の列が変更されています。\n大変申し訳ありませんが正しく読み込めません。")
            sleep(10)
            exit()

        
                
        """
        授業リストの生成
        
        """
        youbi= 0 #曜日ごとに取り出すための変数
        your_class = 0 #授業を入れる変数
        #授業リストの生成
        mon_class= []
        tue_class= []
        wed_class= []
        thi_class= []
        fri_class= []
        for co in range(1,26,5):#列(月から金まで)　列=column
            youbi += 1
            for ro in range(2,12,2):#行(1から5限取までり出す)  行=row
                if youbi ==1:
                    your_class = excelfile.iloc[ro, co]
                    mon_class.append(your_class)
                elif youbi ==2:
                    your_class = excelfile.iloc[ro, co]
                    tue_class.append(your_class)
                elif youbi ==3:
                    your_class = excelfile.iloc[ro, co]
                    wed_class.append(your_class)
                elif youbi ==4:
                    your_class = excelfile.iloc[ro, co]
                    thi_class.append(your_class)
                elif youbi ==5:
                    your_class = excelfile.iloc[ro, co]
                    fri_class.append(your_class)
                
                print(str(youbi) +"回目取り出し完了")
        print(mon_class)
        print(tue_class)
        print(wed_class)
        print(thi_class)
        print(fri_class)
        os.mkdir("1年前期")
        countday =0

        
        """
        授業フォルダの生成
        
        """
        mon_perfect = []
        tue_perfect = []
        wed_perfect = []
        thi_perfect = []
        fri_perfect = []
        
        for youbifile in ["1.月","2.火","3.水","4.木","5.金"]:
            path3 = './1年前期/' +youbifile
            os.mkdir(path3)
            print("曜日フォルダの生成に成功しました")
            bango = 0
            countday += 1



            if countday == 1:
                daylist = mon_class
                day_perfect = coma_make(daylist,bango,path3)
                mon_perfect = day_perfect

            elif countday == 2:
                daylist =tue_class
                day_perfect = coma_make(daylist,bango,path3)
                tue_perfect = day_perfect
            elif countday == 3:
                daylist = wed_class
                day_perfect = coma_make(daylist,bango,path3)
                wed_perfect = day_perfect
            elif countday == 4:
                daylist = thi_class
                day_perfect = coma_make(daylist,bango,path3)
                thi_perfect = day_perfect
            elif countday == 5:
                daylist = fri_class
                day_perfect = coma_make(daylist,bango,path3)
                fri_perfect = day_perfect

        #特別日処理
        path_holiday='./1年前期/6.土日授業'
        pathyobi = './1年前期/7.補講・振替日'
        path_syutyu ='./1年前期/8.集中講義'
        
        os.mkdir(pathyobi)
        os.mkdir(path_holiday)
        os.mkdir(path_syutyu)
        pathyobi1 ='./1年前期/7.補講・振替日/0711(補講日)'
        pathyobi2 ='./1年前期/7.補講・振替日/0716(補講日)'
        pathyobi3 ='./1年前期/7.補講・振替日'
        
        os.mkdir(pathyobi1)
        os.mkdir(pathyobi2)
        os.chdir(pathyobi3)
        #水曜日振替授業についてのテキストファイルを生成
        with open("水曜日振替授業について.txt", "w") as f:
            f.write("4月28日現在の情報では7月4日、7月17日の土曜は水曜日時程で授業を行うとのことです。(令和2年度授業時間割より)\nそのため、これら2つの日付フォルダはここではなく、水曜日の授業の中に入れておきました。\n")
        os.chdir('../')
        os.chdir('../')
        #水曜日振替授業を水曜日の授業フォルダに生成
        for furikae in wed_perfect:
            path5 = str(furikae) + slash + "0704(振替日)"
            path6 = str(furikae) + slash + "0717(振替日)"
            os.mkdir(path5)
            os.mkdir(path6)
        print("水曜日振替授業対応の為の操作完了")
        print("補講・振替日フォルダ生成に成功しました")
        print("土日授業フォルダ生成に成功しました")
        print("集中講義フォルダ生成に成功しました")

        #ここから日付フォルダ生成
        print("これより日付フォルダの生成を開始します。")
        print(mon_perfect)
        print(mon_all)
        Hizuke_folder_maker(mon_perfect,mon_all)
        print(tue_perfect)
        print(tue_all)
        Hizuke_folder_maker(tue_perfect,tue_all)
        print(wed_perfect)
        print(wed_all)
        Hizuke_folder_maker(wed_perfect,wed_all)
        print(thi_perfect)
        print(thi_all)
        Hizuke_folder_maker(thi_perfect,thi_all)
        print(fri_perfect)
        print(fri_all)
        Hizuke_folder_maker(fri_perfect,fri_all)
        

                

        #終了処理
        print("---------------------------------------------------------------------")
        print("すべての処理が終了しました。お疲れさまでした。\n10秒後に終了します。")
        print("---------------------------------------------------------------------")
        sleep(10)
        exit()
