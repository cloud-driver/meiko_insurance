#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import shutil
import os
import smtplib
import time
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication 
import xlwings as xw
import json

def send_email(subject, dir, files, body, to_email):
    from_email = 'cheng520justus@gmail.com'
    password = 'jjgh bbkj qmvn xrah'

    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))
    
    for file in files:
        filepath = os.path.join(dir, file)
        with open(filepath, 'rb') as f:
            part = MIMEApplication(f.read())
            part.add_header('Content-Disposition', 'attachment', filename=file)
            msg.attach(part)

    with smtplib.SMTP(host="smtp.gmail.com", port=587) as smtp:
        try:
            smtp.ehlo()
            smtp.starttls()
            smtp.login(from_email, password)
            smtp.send_message(msg)
            print("成功傳送")
        except Exception as e:
            print("Error message: ", e)

def copy_ex(classroom):
    t = time.time()
    tt = time.localtime(t)
    src_ad = 'example/加保.xlsx'
    src_re = 'example/退保.xlsx'
    dst_ad = os.path.join(classroom, f"{classroom}加保{time.strftime('%Y%m%d',tt)[2:]}.xlsx")
    dst_re = os.path.join(classroom, f"{classroom}退保{time.strftime('%Y%m%d',tt)[2:]}.xlsx")

    if not os.path.exists(classroom):
        os.mkdir(classroom)

    shutil.copyfile(src_ad, dst_ad)
    shutil.copyfile(src_re, dst_re)

def edit_ex(classroom, names):
    worst_msg = []
    files = os.listdir(classroom)
    with open(os.path.join("json", f"{classroom}data.json"), "r", encoding="utf-8") as f:
        users = json.load(f)
    for file in files:
        filepath = os.path.join(classroom, file)
        app = xw.App(visible=False, add_book=False)
        try:
            wb = app.books.open(filepath)
            sheet = wb.sheets[0]

            if "加保" in file:
                for i, name in enumerate(names):
                    for user in users:
                        if user["name"] == name:
                            sheet.cells(i + 2, 6).value = name
                            sheet.cells(i + 2, 7).value = user["id"]
                            sheet.cells(i + 2, 8).value = user["birth"]
                            sheet.cells(i + 2, 9).value = user["salary"]
                            sheet.cells(i + 2, 1).value = "4"
                            sheet.cells(i + 2, 2).value = user["type"]
                            sheet.cells(i + 2, 3).value = user["classroom"]
                            sheet.cells(i + 2, 4).value = user["check"]
                            sheet.cells(i + 2, 10).value = user["special_identity"]
                            sheet.cells(i + 2, 14).value = user["pay_identity"]
                            sheet.cells(i + 2, 15).value = user["rate"]
                            break
                    else:
                        worst_msg.append(f"找不到 {name} 的資料")
            else:
                for i, name in enumerate(names):
                    for user in users:
                        if user["name"] == name:
                            sheet.cells(i + 2, 5).value = name
                            sheet.cells(i + 2, 6).value = user["id"]
                            sheet.cells(i + 2, 8).value = user["birth"]
                            sheet.cells(i + 2, 1).value = "2"
                            sheet.range(i + 2, 1).color = (255, 255, 0)
                            sheet.cells(i + 2, 2).value = user["classroom"]
                            sheet.range(i + 2, 2).color = (255, 255, 0)
                            sheet.cells(i + 2, 3).value = user["check"]
                            sheet.range(i + 2, 3).color = (255, 255, 0)
                            break

            wb.save()
            wb.close()
        except Exception as e:
            print(f"處理 {file} 時發生錯誤：{e}")
        finally:
            app.quit()

    return worst_msg

def get_email(classroom):
    with open(r"json/emails.json", "r", encoding="utf-8") as f:
        emails = json.load(f)
    return emails[classroom]

if __name__ == "__main__":
    print("這裡是自建函式庫，你點錯了，請使用 app.py 發送資料測試")
