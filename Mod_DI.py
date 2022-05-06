from tkinter.dialog import DIALOG_ICON



from calendar import month
from faulthandler import dump_traceback
from traceback import FrameSummary
import pandas as pd
import time 
from pywinauto import *
import glob
import os.path
import openpyxl
import sys
import pyperclip
from pywinauto.keyboard import send_keys

from openpyxl.styles import PatternFill
from datetime import date, timedelta
from datetime import datetime
import matplotlib
import math

import numpy as np
import matplotlib.pyplot as plt

from selenium import webdriver

import win32com.client
import os
import glob


#-------------------------------------------------------
def get_ultimacurva():
        #----------------Browser
        driver = webdriver.Chrome(executable_path=r'C:\Users\luter\OneDrive\Documents\Bloomberg\chromedriver.exe')
        driver.get('https://www2.bmf.com.br/pages/portal/bmfbovespa/boletim1/txref1.asp')
        curva = driver.find_element_by_xpath('/html/body/form/table[1]/tbody/tr[2]/td[2]/img')
        curva.click()
        #--------------Get .xls file

        folder_path = r'C:\Users\luter\Downloads'
        file_type = r'\*xls'
        files = glob.glob(folder_path + file_type)
        latest_v = max(files, key=os.path.getctime)

        #-----------------Change Directory
        o = win32com.client.Dispatch("Excel.Application")
        o.Visible = False
        input_dir = latest_v
        out_dir = r'C:\Users\luter\OneDrive\Documents\DI'
        file = os.path.basename(latest_v)
        output = out_dir + '/' + file.replace('.xls','.xlsx')
        wb = o.Workbooks.Open(latest_v)
        wb.ActiveSheet.SaveAs(output,51)
        wb.Close(True)


#--------------------------Done
def get_dfDI():

        folder_path = r'C:\Users\luter\OneDrive\Documents\DI'
        file_type = r'\*xlsx'
        files = glob.glob(folder_path + file_type)
        latest_v = max(files, key=os.path.getctime)
        data_curva = latest_v.lstrip(folder_path + "PRE")
        data_curva = data_curva.rstrip(" .xlsx")
        y = int(data_curva[0:4])
        m = int(data_curva[4:6])
        d = int(data_curva[6:8])
        df = pd.read_excel(latest_v)
        r = len(df)
        c =  df.shape[1]
        di = pd.DataFrame(df.iloc[:,1], columns = ['days', 'di_252','di_360'])
        dia_curva = datetime(year = y, month = m, day = d)
        for i in range(r-1):
                di = di.append(pd.Series(), ignore_index=True)

        for i in range(0,r-1):
                di.iat[i,1] = float(df.iat[i+1,1].replace(",","."))/100
                di.iat[i,2] = float(df.iat[i+1,2].replace(",","."))/100
                di.iat[i,0] = dia_curva + timedelta(days=(df.iat[i+1,0]))

        return di

di = get_dfDI()
#-----------------------Done v user
def interpola_di(di, dia):
        r = len(di)
        print("Please format date in DD/MM/YYYY \nIs your format correct? Y/N")
        res = input()
        if res != 'Y':
               sys.exit("Please enter dates with correct fomart")
        y = int(dia[6:10])
        m = int(dia[3:5])
        d = int(dia[0:2])
        data = datetime(year = y, month = m, day = d)
        c1 = (data- di.iat[0,0]).days
        for i in range(0,r-1):
                if abs(int((data - di.iat[i,0]).days)) <  c1:
                        c1 = int((data - di.iat[i,0]).days)
                        c2 = i
        d = di.iat[0,0]
        d1 = di.iat[c2,0]
        d2 = di.iat[c2+1,0]
        tx1 = di.iat[c2,2]
        tx2 = di.iat[c2+1,2]
        fct1 = (1+tx1)**((d1 - d).days/365)
        fct2 = (1+tx2)**((d2 - d).days/365)
        fct = fct2/fct1
        tx3 = fct**(365/(d2-d1).days) -1
        fct3 = (1+ tx3)**((data-d1).days/365)
        tx3 = (fct3*fct1)**(365/(data-d).days)-1
        return tx3
#-------------------------V Module Factor 360
def interpola_di_mod_fct(di, dia):
 
        r = len(di)
        data = dia
        c1 = (data - di.iat[0,0]).days               
        for i in range(0,r-1):
                if abs(int((data - di.iat[i,0]).days)) <  c1:
                        c1 = int((data - di.iat[i,0]).days)
                        c2 = i
        d = di.iat[0,0]
        d1 = di.iat[c2,0]
        d2 = di.iat[c2+1,0]
        tx1 = di.iat[c2,2]
        tx2 = di.iat[c2+1,2]
        fct1 = (1+tx1)**((d1 - d).days/365)
        fct2 = (1+tx2)**((d2 - d).days/365)
        fct = fct2/fct1
        tx3 = fct**(365/(d2-d1).days) -1
        fct3 = (1+ tx3)**((data-d1).days/365)
        tx3 = (fct3*fct1)**(365/(data-d).days)-1
        fct = fct3*fct1
        return fct
#--------------------------V Module Rate 360
def interpola_di_mod_tx(di, dia):
 
        r = len(di)
        data = dia
        c1 = (data - di.iat[0,0]).days               
        for i in range(0,r-1):
                if abs(int((data - di.iat[i,0]).days)) <  c1:
                        c1 = int((data - di.iat[i,0]).days)
                        c2 = i
        d = di.iat[0,0]
        d1 = di.iat[c2,0]
        d2 = di.iat[c2+1,0]
        tx1 = di.iat[c2,2]
        tx2 = di.iat[c2+1,2]
        fct1 = (1+tx1)**((d1 - d).days/365)
        fct2 = (1+tx2)**((d2 - d).days/365)
        fct = fct2/fct1
        tx3 = fct**(365/(d2-d1).days) -1
        fct3 = (1+ tx3)**((data-d1).days/365)
        tx3 = (fct3*fct1)**(365/(data-d).days)-1
        fct = fct3*fct1
        return tx3
#------------------------- V Module Df 360
def interpola_di_mod_df(di, dia):
 
        r = len(di)
        data = dia
        c1 = (data - di.iat[0,0]).days               
        for i in range(0,r-1):
                if abs(int((data - di.iat[i,0]).days)) <  c1:
                        c1 = int((data - di.iat[i,0]).days)
                        c2 = i
        d = di.iat[0,0]
        d1 = di.iat[c2,0]
        d2 = di.iat[c2+1,0]
        tx1 = di.iat[c2,2]
        tx2 = di.iat[c2+1,2]
        fct1 = (1+tx1)**((d1 - d).days/365)
        fct2 = (1+tx2)**((d2 - d).days/365)
        fct = fct2/fct1
        tx3 = fct**(365/(d2-d1).days) -1
        fct3 = (1+ tx3)**((data-d1).days/365)
        tx3 = (fct3*fct1)**(365/(data-d).days)-1
        fct = fct3*fct1
        return 1/tx3
 
#------------------------------------ V Module Generic 360

def interpola_di_mod(di, dia, output):
        if output == "exp252":
                return interpola_di_mod_tx(di, dia)
        elif output == "fct":
                return interpola_di_mod_fct(di, dia)
        else:
                return interpola_di_mod_df(di, dia)

#-------------------------V Module Factor 252
def interpola_di_mod_fct_252(di, dia):
 
        r = len(di)
        data = dia
        c1 = (data - di.iat[0,0]).days               
        for i in range(0,r-1):
                if abs(int((data - di.iat[i,0]).days)) <  c1:
                        c1 = int((data - di.iat[i,0]).days)
                        c2 = i
        d = di.iat[0,0]
        d1 = di.iat[c2,0]
        d2 = di.iat[c2+1,0]
        tx1 = di.iat[c2,1]
        tx2 = di.iat[c2+1,1]
        fct1 = (1+tx1)**((d1 - d).days/365)
        fct2 = (1+tx2)**((d2 - d).days/365)
        fct = fct2/fct1
        tx3 = fct**(365/(d2-d1).days) -1
        fct3 = (1+ tx3)**((data-d1).days/365)
        tx3 = (fct3*fct1)**(365/(data-d).days)-1
        fct = fct3*fct1
        return fct
#--------------------------V Module Rate 252
def interpola_di_mod_tx_252(di, dia):
 
        r = len(di)
        data = dia
        c1 = (data - di.iat[0,0]).days               
        for i in range(0,r-1):
                if abs(int((data - di.iat[i,0]).days)) <  c1:
                        c1 = int((data - di.iat[i,0]).days)
                        c2 = i
        d = di.iat[0,0]
        d1 = di.iat[c2,0]
        d2 = di.iat[c2+1,0]
        tx1 = di.iat[c2,1]
        tx2 = di.iat[c2+1,1]
        fct1 = (1+tx1)**((d1 - d).days/365)
        fct2 = (1+tx2)**((d2 - d).days/365)
        fct = fct2/fct1
        tx3 = fct**(365/(d2-d1).days) -1
        fct3 = (1+ tx3)**((data-d1).days/365)
        tx3 = (fct3*fct1)**(365/(data-d).days)-1
        fct = fct3*fct1
        return tx3
#------------------------- V Module Df 252
def interpola_di_mod_df_252(di, dia):
 
        r = len(di)
        data = dia
        c1 = (data - di.iat[0,0]).days               
        for i in range(0,r-1):
                if abs(int((data - di.iat[i,0]).days)) <  c1:
                        c1 = int((data - di.iat[i,0]).days)
                        c2 = i
        d = di.iat[0,0]
        d1 = di.iat[c2,0]
        d2 = di.iat[c2+1,0]
        tx1 = di.iat[c2,1]
        tx2 = di.iat[c2+1,1]
        fct1 = (1+tx1)**((d1 - d).days/365)
        fct2 = (1+tx2)**((d2 - d).days/365)
        fct = fct2/fct1
        tx3 = fct**(365/(d2-d1).days) -1
        fct3 = (1+ tx3)**((data-d1).days/365)
        tx3 = (fct3*fct1)**(365/(data-d).days)-1
        fct = fct3*fct1
        return 1/tx3
 
#------------------------------------ V Module Generic 252

def interpola_di_mod_252(di, dia, output):
        if output == "exp252":
                return interpola_di_mod_tx_252(di, dia)
        elif output == "fct":
                return interpola_di_mod_fct_252(di, dia)
        else:
                return interpola_di_mod_df_252(di, dia)



#--------------------------------------------- FRA module
def get_FRAS(dias, anos):

        di = get_dfDI()
        fras = pd.DataFrame(di.iloc[:,1], columns = ['days', 'FRAs'])
        r = len(di)


        d = anos * 365
        mod = d/dias
        mod= math.trunc(mod)

        for i in range(mod-1):
                fras = fras.append(pd.Series(), ignore_index=True)

        d = di.iat[0,0]
        d
        type(d)        
        for i in range(mod-1):
                if i == 0:
                        d = d + timedelta(days = dias) 
                        fras.iat[i,0] = d
                        fras.iat[i,1] = interpola_di_mod_tx(di,d)
                else:
                        d = d + timedelta(days = dias) 
                        fras.iat[i,0] = d
                        fras.iat[i,1] = (interpola_di_mod_fct(di,d)/interpola_di_mod_fct(di,(d-timedelta(days=dias))))**(360/dias)-1


        return fras



#--------------------------------------------- FRA module 252
def get_FRAS_252(dias, anos):

        di = get_dfDI()
        fras = pd.DataFrame(di.iloc[:,1], columns = ['days', 'FRAs'])
        r = len(di)


        d = anos * 365
        mod = d/dias
        mod= math.trunc(mod)

        for i in range(mod-1):
                fras = fras.append(pd.Series(), ignore_index=True)

        d = di.iat[0,0]
        d
        type(d)        
        for i in range(mod-1):
                if i == 0:
                        d = d + timedelta(days = dias) 
                        fras.iat[i,0] = d
                        fras.iat[i,1] = interpola_di_mod_tx_252(di,d)
                else:
                        d = d + timedelta(days = dias) 
                        fras.iat[i,0] = d
                        fras.iat[i,1] = (interpola_di_mod_fct_252(di,d)/interpola_di_mod_fct_252(di,(d-timedelta(days=dias))))**(360/dias)-1


        return fras


get_ultimacurva()

dFRA = get_FRAS_252(125,5)

dFRA
#dFRA["A"] = pd.Series(list(range(len(dFRA))))
#dFRA.plot(x="days", y="FRAs").lines
plt.plot(dFRA.days, dFRA.FRAs)
plt.show()