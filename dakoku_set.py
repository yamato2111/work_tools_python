#打刻時間になるとアラートをあげて打刻画面を表示する
#windowsのタスクスケジューラにバッチファイルを自動登録

from tkinter import *
import tkinter as tk
import tkinter.ttk as ttk
from tkinter import messagebox
import os
import sys
import subprocess

#ルートフレーム作成
root = tk.Tk()
root.title("打刻アラート")
root.geometry("300x280")

#ボタンクリック時のイベント
def btn_click():
    mon = select_mon.get()
    tue = select_tue.get()
    wed = select_wed.get()
    thu = select_thu.get()
    fri = select_fri.get()
    sat = select_sat.get()
    sun = select_sun.get()
    kakunin = messagebox.askquestion("確認","こちらでよろしいですか？\n\n月曜：%s \n火曜：%s \n水曜：%s \n木曜：%s \n金曜：%s \n土曜：%s \n日曜：%s " % (mon,tue,wed,thu,fri,sat,sun) )#メッセージボックス
    print("askquestion", kakunin)
    #タスク作成コマンド作成
    if kakunin == "yes":
        #アラート実行ファイル作成
        copy_command = "xcopy dakoku_alert C:\\Users\\%username%\\Documents\\dakoku_alert\\ /S /Y"
        os.system(copy_command)
        #古いタスクを削除
        command1 = "schtasks /delete /tn dakoku_asa_in /f"
        command2 = "schtasks /delete /tn dakoku_asa_out /f"
        command3 = "schtasks /delete /tn dakoku_oso_in /f"
        command4 = "schtasks /delete /tn dakoku_oso_out /f"
        command5 = "schtasks /delete /tn dakoku_naka_a_in /f"
        command6 = "schtasks /delete /tn dakoku_naka_a_out /f"
        command7 = "schtasks /delete /tn dakoku_naka_b_in /f"
        command8 = "schtasks /delete /tn dakoku_naka_b_out /f"
        command9 = "schtasks /delete /tn dakoku_sat_in /f"
        command10 = "schtasks /delete /tn dakoku_sat_out /f"
        command11 = "schtasks /delete /tn dakoku_sun_in /f"
        command12 = "schtasks /delete /tn dakoku_sun_out /f"
        os.system(command1)
        os.system(command2)
        os.system(command3)
        os.system(command4)
        os.system(command5)
        os.system(command6)
        os.system(command7)
        os.system(command8)
        os.system(command9)
        os.system(command10)
        os.system(command11)
        os.system(command12)
        os.system("cls")
        #タスク作成処理
        asa_list = []
        oso_list = []
        nak_list = []
        #月曜
        if mon == "朝番":
            asa_list.append("mon")
        elif mon == "遅番":
            oso_list.append("mon")
        elif mon == "中抜け":
            nak_list.append("mon")
        #火曜
        if tue == "朝番":
            asa_list.append("tue")
        elif tue == "遅番":
            oso_list.append("tue")
        elif tue == "中抜け":
            nak_list.append("tue")
        #水曜
        if wed == "朝番":
            asa_list.append("wed")
        elif wed == "遅番":
            oso_list.append("wed")
        elif wed == "中抜け":
            nak_list.append("wed")
        #木曜
        if thu == "朝番":
            asa_list.append("thu")
        elif thu == "遅番":
            oso_list.append("thu")
        elif thu == "中抜け":
            nak_list.append("thu")
        #金曜
        if fri == "朝番":
            asa_list.append("fri")
        elif fri == "遅番":
            oso_list.append("fri")
        elif fri == "中抜け":
            nak_list.append("fri")
        #タスク作成コマンド作成
        if len(asa_list) >= 1:
            mojiretsu = ""
            for youbi in asa_list:
                mojiretsu += youbi
                mojiretsu += ","
            add_youbi = mojiretsu.rstrip(",")
            asaIn_com = "schtasks /create /tn dakoku_asa_in /tr C:\\Users\\%username%\\Documents\\dakoku_alert\\run.bat /sc weekly /d " + add_youbi + " /st 08:45"
            asaOut_com = "schtasks /create /tn dakoku_asa_out /tr C:\\Users\\%username%\\Documents\\dakoku_alert\\run.bat /sc weekly /d " + add_youbi + " /st 18:00"
            os.system(asaIn_com)
            os.system(asaOut_com)
        if len(oso_list) >= 1:
            mojiretsu = ""
            for youbi in oso_list:
                mojiretsu += youbi
                mojiretsu += ","
            add_youbi = mojiretsu.rstrip(",")
            osoIn_com = "schtasks /create /tn dakoku_oso_in /tr C:\\Users\\%username%\\Documents\\dakoku_alert\\run.bat /sc weekly /d " + add_youbi + " /st 12:45"
            osoOut_com = "schtasks /create /tn dakoku_oso_out /tr C:\\Users\\%username%\\Documents\\dakoku_alert\\run.bat /sc weekly /d " + add_youbi + " /st 22:00"
            os.system(osoIn_com)
            os.system(osoOut_com)
        if len(nak_list) >= 1:
            mojiretsu = ""
            for youbi in nak_list:
                mojiretsu += youbi
                mojiretsu += ","
            add_youbi = mojiretsu.rstrip(",")
            nakAIn_com = "schtasks /create /tn dakoku_naka_a_in /tr C:\\Users\\%username%\\Documents\\dakoku_alert\\run.bat /sc weekly /d " + add_youbi + " /st 08:45"
            nakAOut_com = "schtasks /create /tn dakoku_naka_a_out /tr C:\\Users\\%username%\\Documents\\dakoku_alert\\run.bat /sc weekly /d " + add_youbi + " /st 13:30"
            nakBIn_com = "schtasks /create /tn dakoku_naka_b_in /tr C:\\Users\\%username%\\Documents\\dakoku_alert\\run.bat /sc weekly /d " + add_youbi + " /st 18:15"
            nakBOut_com = "schtasks /create /tn dakoku_naka_b_out /tr C:\\Users\\%username%\\Documents\\dakoku_alert\\run.bat /sc weekly /d " + add_youbi + " /st 22:00"
            os.system(nakAIn_com)
            os.system(nakAOut_com)
            os.system(nakBIn_com)
            os.system(nakBOut_com)
        if sat == "出勤":
            satIn_com = "schtasks /create /tn dakoku_sat_in /tr C:\\Users\\%username%\\Documents\\dakoku_alert\\run.bat /sc weekly /d sat /st 09:15"
            satOut_com = "schtasks /create /tn dakoku_sat_out /tr C:\\Users\\%username%\\Documents\\dakoku_alert\\run.bat /sc weekly /d sat /st 18:30"
            os.system(satIn_com)
            os.system(satOut_com)
        if sun == "出勤":
            sunIn_com = "schtasks /create /tn dakoku_sun_in /tr C:\\Users\\%username%\\Documents\\dakoku_alert\\run.bat /sc weekly /d sun /st 09:15"
            sunOut_com = "schtasks /create /tn dakoku_sun_out /tr C:\\Users\\%username%\\Documents\\dakoku_alert\\run.bat /sc weekly /d sun /st 18:30"
            os.system(sunIn_com)
            os.system(sunOut_com)
        fin = messagebox.showinfo("完了", "設定が完了しました。")
        if fin == "ok":
            sys.exit()
    return 0

#月曜
select_mon = StringVar()
lbl_mon = tk.Label(text="月曜：")
lbl_mon.place(x=30, y=20)
combo_mon = ttk.Combobox(root, state='readonly', textvariable=select_mon)
combo_mon["values"] = ("朝番","遅番","中抜け", "休み")
combo_mon.current(3)
combo_mon.place(x=90, y=20)
#火曜
select_tue = StringVar()
lbl_tue = tk.Label(text="火曜：")
lbl_tue.place(x=30, y=50)
combo_tue = ttk.Combobox(root, state='readonly', textvariable=select_tue)
combo_tue["values"] = ("朝番","遅番","中抜け", "休み")
combo_tue.current(3)
combo_tue.place(x=90, y=50)
#水曜
select_wed = StringVar()
lbl_wed = tk.Label(text="水曜：")
lbl_wed.place(x=30, y=80)
combo_wed = ttk.Combobox(root, state='readonly', textvariable=select_wed)
combo_wed["values"] = ("朝番","遅番","中抜け", "休み")
combo_wed.current(3)
combo_wed.place(x=90, y=80)
#木曜
select_thu = StringVar()
lbl_thu = tk.Label(text="木曜：")
lbl_thu.place(x=30, y=110)
combo_thu = ttk.Combobox(root, state='readonly', textvariable=select_thu)
combo_thu["values"] = ("朝番","遅番","中抜け", "休み")
combo_thu.current(3)
combo_thu.place(x=90, y=110)
#金曜
select_fri = StringVar()
lbl_fri = tk.Label(text="金曜：")
lbl_fri.place(x=30, y=140)
combo_fri = ttk.Combobox(root, state='readonly', textvariable=select_fri)
combo_fri["values"] = ("朝番","遅番","中抜け", "休み")
combo_fri.current(3)
combo_fri.place(x=90, y=140)
#土曜
select_sat = StringVar()
lbl_sat = tk.Label(text="土曜：")
lbl_sat.place(x=30, y=170)
combo_sat = ttk.Combobox(root, state='readonly', textvariable=select_sat)
combo_sat["values"] = ("出勤", "休み")
combo_sat.current(1)
combo_sat.place(x=90, y=170)
#日曜
select_sun = StringVar()
lbl_sun = tk.Label(text="日曜：")
lbl_sun.place(x=30, y=200)
combo_sun = ttk.Combobox(root, state='readonly', textvariable=select_sun)
combo_sun["values"] = ("出勤", "休み")
combo_sun.current(1)
combo_sun.place(x=90, y=200)

#ボタン作成

btn = tk.Button(root, text="確定", width=10, command=btn_click)
btn.place(x=110, y=240)

root.mainloop()
