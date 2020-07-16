from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import pandas as pd
import datetime
import os
import win32com.client as win32
import xlwings as xw
import time
import tkinter as tk
from tkinter import filedialog
from tkinter import *
import logging as log
root = tk.Tk()
root.withdraw()
import os.path
file_path = filedialog.askdirectory()
file_pathre = file_path.replace("/","\\")

def Chrome():
    global driver
    log.info("Initializing Chrome browser")
    options = webdriver.ChromeOptions()
    options.add_experimental_option("prefs", {
        "download.default_directory": file_pathre,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    })
    driver = webdriver.Chrome(options=options)
    abcd(driver)
    exit()
def Mozilla_firefox():
    global driver
    log.info("Initializing Firefox browser")
    profile = webdriver.FirefoxProfile()
    profile.set_preference('browser.download.folderList', 2) # custom location
    profile.set_preference('browser.download.manager.showWhenStarting', False)
    profile.set_preference('browser.download.dir', file_pathre)
    profile.set_preference('browser.helperApps.neverAsk.saveToDisk', 'text/csv')
    driver = webdriver.Firefox(profile)
    abcd(driver)
    exit()
def main_screen():
    global screen
    # photo = PhotoImage(file="UCS.png")
    screen = Toplevel(screen2)
    screen.geometry("300x250")
    screen.title("Download Active List")
    # Label(screen , image = photo)
    # Label.pack()
    Label(screen,text= "Download Active List",bg = "Red" ,  width = "300" ,height = "2" ,font = ("calibri",13)).pack()
    Label(screen,text= "").pack()
    Button(screen,text = "Chrome",height = "2" , width = "25", command = Chrome).pack()
    Label(screen,text="").pack()
    Button(screen,text = "Mozilla firefox" ,height = "2" , width = "25" , command = Mozilla_firefox) .pack()
    Label(screen, text="").pack()
    Button(screen, text="Exit", height="2", width="25", command=exitall).pack()
def exitall():
    screen.destroy()
    screen2.destroy()
    exit()

def login():
    global screen2
    screen2 = Tk()
    screen2.title("Pls enter the Details")
    screen2.geometry("350x250")
    Label(screen2, text = "Please enter the username and Password of the Company").pack()
    Label(screen2, text = "").pack()

    global username_verify
    global password_verify
    username_verify = StringVar()
    password_verify = StringVar()

    global username_entry1
    global password_entry1
    Label(screen2, text = "Username * ").pack()
    username_entry1 = Entry(screen2, textvariable = username_verify)
    username_entry1.pack()
    Label(screen2, text = "").pack()
    Label(screen2, text = "Password * ").pack()
    password_entry1 = Entry(screen2, textvariable = password_verify)
    password_entry1.pack()
    Label(screen2, text= "").pack()
    Button(screen2, text = "Next",height = 1 , width = 10 , command = main_screen).pack()
    screen2.mainloop()



def delete3():
    screen2.destroy()



def abcd(driver):
    driver.get("https://unifiedportal-emp.epfindia.gov.in/epfo/");
    username_word = username_entry1.get()
    password_word = password_entry1.get()
    if os.path.isfile(file_pathre + r"\ActiveMember.csv"):
        finalpath = (file_pathre + r"\ActiveMember(1).csv")
        finalpath2 = (file_pathre + r"\ActiveMember(1) ")
    else:
        finalpath = (file_pathre + r"\ActiveMember.csv")
        finalpath2 = (file_pathre + r"\ActiveMember ")

    xpath = ("//input[@id='username']")
    if wait_for(driver, xpath) == 1:
        el = driver.find_element_by_xpath("//input[@id='username']")
        el.send_keys(username_word)
        el = driver.find_element_by_xpath("//input[@id='password']")
        el.send_keys(password_word)

    else:
        exit()

    el = driver.find_element_by_xpath("//button[@class='btn btn-success btn-logging']")
    el.click()

    time.sleep(2)

    xpath2 = ("(//a[@class='dropdown-toggle'])[4]")
    if wait_for(driver, xpath2) == 1:
        el = driver.find_element_by_xpath(xpath2).click()
        el = driver.find_element_by_xpath("//*[text()[contains(., 'ACTIVE MEMBER')]]").click()
    else:
        exit()

    time.sleep(10)
    xpath3 = ("//a[@title='Download all active members']")
    if wait_for(driver, xpath3) == 1:
        el2 = driver.find_element_by_xpath(xpath3).click()
    else:
        exit()

    radate = datetime.date.today()
    x = datetime.datetime(radate.year, radate.month, radate.day)
    lastdate = x.strftime("%d%m%Y")
    for j in range(1, 30, 2):
        try:
            os.rename(finalpath, finalpath2 + lastdate + ".csv")
            break
        except:
            time.sleep(1)
            print("Nothing")

    inputxl1 = finalpath2 + lastdate + ".csv"
    df1 = pd.read_csv(inputxl1)
    writer1 = pd.ExcelWriter(finalpath2 + lastdate + ".xlsx")
    df1.to_excel(writer1, sheet_name="Working", index=False)
    df1.to_excel(writer1, sheet_name="ActiveMember", index=False)
    workbook = writer1.book
    writer1.save()

    wb1 = xw.Book("Active Member.xlsm")
    sheet = wb1.sheets['Sheet1']
    sheet.range('A2').value = finalpath2 + lastdate + ".xlsx"
    wb1.save()
    wb1.close()

    xl = win32.Dispatch('Excel.Application')
    xl.Application.visible = True  # change to True if you are desired to make Excel visible

    try:
        wb = xl.Workbooks.Open(os.path.abspath("Active Member.xlsm"))
        xl.Application.run("'Active Member.xlsm'" + "!Module1.Active_Member")
        wb.Save()
        wb.Close()
    except Exception as ex:
        template = "An exception of type {0} occurred. Arguments:\n{1!r}"
        message = template.format(type(ex).__name__, ex.args)
        print(message)


def wait_for(driver, xpath, wait=30):
    waited = 0
    for i in range(1, wait+1, 2):
        try:
            driver.find_element_by_xpath(xpath)
            print("waited for " + str(waited) + " seconds")
            return 1
        except:
            time.sleep(2)
            waited = waited + 2
    return 0




login()
