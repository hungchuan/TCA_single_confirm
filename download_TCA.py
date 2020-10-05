#!/usr/bin/env python

import os
import sys
import re
import xlrd
#import splinter
import tca as TCA #URT
import time
import pandas as pd
import pygsheets
import numpy as np
import shutil 
import pdb
from selenium import webdriver
from googletrans import Translator
from datetime import datetime, date, timedelta
from print_log import log_print,Emptyprintf
from fnmatch import fnmatch, fnmatchcase

#PROJECT_IDS = [1429]
# PROJECT_IDS = [712, 746, 750]

ALL_PROBLEMS = "/media/sf_linux-share/all.txt"
PROBLEM = "~/Downloads/TCA.xlsx"
log = Emptyprintf

def isNaN(num):
    return num != num
    
def Google_Translator(src_in,dest_in,df_in):
    #df_in=df_in.fillna("")
    #print("df_in=",df_in)
    translator = Translator()
    # src來源語言，dest要翻譯語言，如果要找其他語言可以參考說明文件
    n=0 
    str="sorun : tv açılmamaktadır cevap : servise yönlendirildi"   
    str_out=translator.translate(str, src = src_in, dest = dest_in)
    print(str_out.text)
    for i in df_in:         
        #if (isNaN(df_in[n])==False):   
        if (len(df_in[n])>1):  
            print(df_in[n])  
            try:
                str_out=translator.translate(df_in[n], src = src_in, dest = dest_in)
            except:    
                n=n;            

            df_in[n] = str_out.text
        n=n+1    
    #print(str_out.text)
    #sel = input("pause")
    return df_in
    


def DF2List(df_in):
    train_data = np.array(df_in)#np.ndarray()
    df_2_list=train_data.tolist()#list
    return df_2_list

def waiting_for_update(br,text1):
    log('text1=',text1)
    count = 0
    result_info = br.find_element_by_name(text1)
    log('result_info=',result_info)
    result_info = br.find_element_by_name('nalndlaknfsln')
    log('result_info2=',result_info)
    
    while (result_info.text != text1):
        log ('waiting for update Count: %d' % count)
        count += 1
        time.sleep (1)
        result_info = br.find_element_by_name(text1)
    return result_info
    
def waiting_for_TCA_update(br,xpath):
    button = False
    for i in range(0,10):
        time.sleep(1) 
        print ('waiting for update Count: %d' % i)
        try:
            button = br.find_element_by_xpath(xpath)
        except:
            continue
        break
    print(xpath)
    print("waiting_for_TCA_update button=",button)
    return button    
    
def download_TCA (br, filename, PACKAGE_DIRECTORY, config):

    button = waiting_for_TCA_update(br,config ['xpath_cashout'])
    time.sleep(5)  
    button.click() #click cashout   
    
    button = waiting_for_TCA_update(br,config ['xpath_domain'])
    button.click() #click domain   
    print('wait 5 seconds................')
    time.sleep(5)              
    button = waiting_for_TCA_update(br,config ['xpath_domain_iframe'])
    print(button.text)
    print('wait 5 seconds................')
    time.sleep(5)  
    br.switch_to.frame(1)
    print('wait 5 seconds................')
    time.sleep(5)           
    button = waiting_for_TCA_update(br,config ['xpath_claims'])
    
    #br.get ('https://gsp.tpv-tech.com/RedirectURL.aspx?pg=c&srcgo=Cashout%2fGCS_HomepageForRDDomain.aspx')
    #button = br.find_element_by_xpath('//*[@id="aRDDomain"]')
    log('button=',button)
    print('button=',button)
    print('button=',button.text)
    if (int(button.text)==0):
        return False
    
    button.click() #click 數量   
    
    button = waiting_for_TCA_update(br,config ['xpath_download_all'])
    log('download all=',button)
    button.click() #click download all

    button = waiting_for_TCA_update(br,config ['xpath_yes'])
    log('yes = ',button)
    button.click() #click yes   
    
    #sel = input("pause")    
    
    return button     
    
    '''    
    #log('TCA_df=',TCA_df)  
    En_df = pd.DataFrame(columns=['En','En1'])
    En_df['En']=TCA_df['Fault Description Text']
    En_df["En1"] = En_df["En"].str.replace("\n"," ")
    En_df["En1"] = En_df["En1"].str.replace("\r"," ")
    En_df["En1"] = En_df["En1"].str.replace("<br/>"," ")
    log('En=',En_df)   
    '''            


def TCA_get_filename(PACKAGE_DIRECTORY):
    count = 0
    # 取得所有檔案與子目錄名稱
    files = []
    while ( len(files) == 0):
        print("waiting for download Count: %d" % count)
        try:
            files = os.listdir(PACKAGE_DIRECTORY)
        except:
            files = []            
        count += 1
        time.sleep (1)
    
    filename = os.path.join(PACKAGE_DIRECTORY, files[0])
    
    return filename

def TCA_backup(filename):
    current_directory = os.path.abspath('.')
    backup_directory = current_directory + '\\backup'
    
    now = datetime.now()  
    date= now.strftime("%Y")+now.strftime("%m")+now.strftime("%d")
    #newfilename = backup_directory+date+'.xlsx'
    writefile=date+'.xlsx'    
    
    newfilename=os.path.join(backup_directory, writefile)
        
    shutil.move(filename,newfilename) 
    '''
    TCA_analysis_file='TCA_analysis.xlsx'
    newfilename=os.path.join(backup_directory, TCA_analysis_file)        
    shutil.move(filename,newfilename)     
    '''
    return newfilename    
    
def TCA_read(filename):
    df = pd.read_excel(filename)
    new_df=df.fillna("")
    TCA_Items = new_df.shape[0]
    
    log('TCA_Items=',TCA_Items)
    TCA_df = pd.DataFrame(columns=['Unique Reference Number','Claim RD Review Date','Workshop Name','Model No','Product No','Repair Result','Regular SW number IN','Regular SW number OUT','Total Spare Part Cost (1-4 + Small) (Local Currency)','Symptom Code 1','Workshop Comment','Fault Description Text'])
    df_column=TCA_df.columns.values.tolist()
    for i in df_column:
        TCA_df[i]=new_df[i]  
    
    TCA_df['Claim RD Review Date'] = pd.to_datetime(TCA_df['Claim RD Review Date'],format='%m-%d-%Y')
    return TCA_df    

def TCA_upload_to_google(file, sheet, TCA_df):
    try:
        gc = pygsheets.authorize(service_file='PythonUpload-cfde37284cdc.json')
    except:
        return filename
    
    sh = gc.open(file)

    try:
        wks = sh.worksheet_by_title(sheet)
    except:
        wks = sh.add_worksheet(sheet,rows=1,cols=30,index=0)    
     
    wks.set_dataframe(TCA_df, (1, 1))
    '''
    sh = gc.open('Translation_Symptom')

    try:
        wks = sh.worksheet_by_title('data in')
    except:
        wks = sh.add_worksheet(gh_name,rows=510,cols=80,index=0)    
     
    wks.set_dataframe(En_df, (1, 1))
    '''
    return wks    


def TCA_download_from_google(file, sheet):

    try:
        gc = pygsheets.authorize(service_file='PythonUpload-cfde37284cdc.json')
    except:
        log("can not find json file")
    
    sh = gc.open(file)

    try:
        wks = sh.worksheet_by_title(sheet)
    except:
        wks = sh.add_worksheet(sheet,rows=1,cols=30,index=0)    
     
    df = wks.get_as_df()   
    return df    
    '''
    new_df = pd.DataFrame(columns=['Unique Reference Number','Claim RD Review Date','Workshop Name','Model No','Product No','Repair Result','Regular SW number IN','Regular SW number OUT','Total Spare Part Cost (1-4 + Small) (Local Currency)','Symptom Code 1','Workshop Comment','Fault Description Text'])
    df_column=new_df.columns.values.tolist()
    for i in df_column:
        new_df[i]=df[i]  
    
    #sh.del_worksheet(wks) # delete this wroksheet

    return new_df 
    '''

def download_from_google(file, sheet):
    try:
        gc = pygsheets.authorize(service_file='PythonUpload-cfde37284cdc.json')
    except:
        log("can not find json file")
    
    sh = gc.open(file)

    try:
        wks = sh.worksheet_by_title(sheet)
    except:
        wks = sh.add_worksheet(sheet,rows=1,cols=30,index=0)    
     
    df = wks.get_as_df()

    return df 
    
def download_issues (br, proj, filename):
    log('proj=',proj)
    #proj = int(proj)
    log('Downloading issues for project: %d' % (proj))
    log('urt=', TCA.TCA_URL + "/Pts/issuelist.aspx?project=%d" % proj)
    
    br.visit (TCA.TCA_URL + "/Pts/issuelist.aspx?project=%d" % proj)
    
    project = br.find_by_xpath('//*[@id="ctl00_pnlHeader"]/table/tbody/tr/td')
    project_name = project.text
    log('project_name = ',project_name)
    name_find= project_name.find('項目列表 »')
    len_title = len('項目列表 » ')
    log('len_title =',len_title)
    gh_name=project_name[name_find+len_title:]
    log('gh_name =',gh_name)
    newfilename = gh_name+'.xls'
    log('newfilename =',newfilename)
    
    result_info = br.find_by_xpath('//*[@id="ctl00_CP1_lblResultInfo"]')
    log('result_info = ',result_info.text)
    
    button= br.find_by_text('所有').click() #click 所有    
    #button = br.find_by_xpath('//*[@id="ctl00_CP1_tvNavt11"]').click() #click 所有
    #URT.wait_for_xpath (br, '//*[@id="ctl00_CP1_tvNavt11"]')
    
    count = 0
    result_info = br.find_by_xpath('//*[@id="ctl00_CP1_lblResultInfo"]')

    while (result_info.text != '所有事務'):
        log ('waiting for update Count: %d' % count)
        count += 1
        time.sleep (1)
        result_info = br.find_by_xpath('//*[@id="ctl00_CP1_lblResultInfo"]')

        
    button= br.find_by_text('導出').click() #click 導出    
    #button = br.find_by_xpath('//*[@id="ctl00_CP1_lnkExport"]').click()#click 導出
    #URT.wait_for_xpath (br, '//*[@id="ctl00_CP1_lnkExport"]')

    #URT.wait_for_xpath (br, "//*[@id='ctl00_CP1_lnkExport']")
    
    filename = os.path.expanduser (filename)
    log('filename_aaa = ',filename)
    
    if (os.path.isfile(filename)==True):
        log('remove file')
        os.remove(filename)
    
    time.sleep(1) 
    #URT.wait_for_xpath (br, "//*[@class='ctl00_CP1_tvNav_0']")
    # //*[@id="ctl00_CP1_btnExport"]
    
   
    #br.find_by_xpath ('//*[@id="ctl00_CP1_btnExport"]').first.click ()

    #URT.wait_for_xpath (br, '//*[@id="ctl00_CP1_btnExport"]')

    #filename = os.path.expanduser (filename)
    #log('filename = ',filename)
    #if os.path.exists (filename):
    #    os.unlink (filename)


    
    count = 0
    result_info = br.find_by_xpath('//*[@id="ctl00_CP1_Label1"]')

    while (result_info.text != '導出事務列表'):
        print ('waiting for update Count: %d' % count)
        count += 1
        time.sleep (1)
        result_info = br.find_by_xpath('//*[@id="ctl00_CP1_Label1"]')

        
    #br.find_by_xpath ('//*[@id="ctl00_CP1_btnExport"]').first.click ()
    br.find_by_xpath ('//*[@id="ctl00_CP1_btnExport"]').click () #click 導出
    #URT.wait_for_xpath (br, '//*[@id="ctl00_CP1_btnExport"]')
    #print('filename111 = ',filename)
    #time.sleep(2) 
    #filename = URT.complete_download (filename, proj)
    
    count = 0
    while (os.path.exists (filename) == False):
        log ('waiting for download Count: %d' % count)
        count += 1
        time.sleep (1)
    
    
    df = pd.read_excel(filename)
    new_df=df.fillna(0)
    
    log('newfilename = ',newfilename)
    PACKAGE_DIRECTORY = os.path.abspath('.')
    newfilename = os.path.join(PACKAGE_DIRECTORY,newfilename)
    print('newfilename222 = ',newfilename)
    shutil.move(filename,newfilename) 
    #new_df.to_excel(newfilename)
    
    try:
        gc = pygsheets.authorize(service_file='PythonUpload-cfde37284cdc.json')
    except:
        return filename
    
    sh = gc.open('URTracker')

    try:
        wks = sh.worksheet_by_title(gh_name)
    except:
        wks = sh.add_worksheet(gh_name,rows=510,cols=80,index=0)    
     
    wks.set_dataframe(new_df, (1, 1))
    
    return filename
   
'''    
    elem = None
    for e in br.find_by_xpath ("//*[@class='ctl00_CP1_tvNav_0']"): 
        if re.search ("Open", e.value):
            elem = e

    if elem != None:
        elem.click ()

        URT.wait_for_update_progress (br, "//*[@id='ctl00_UpdateProgress1']")

        br.find_by_xpath ("//*[@id='ctl00_CP1_lnkExport']").first.click ()

        URT.wait_for_xpath (br, "//*[@id='ctl00_CP1_btnExport']")

        filename = os.path.expanduser (filename)
        if os.path.exists (filename):
            os.unlink (filename)

        br.find_by_xpath ("//*[@id='ctl00_CP1_btnExport']").first.click ()

        filename = URT.complete_download (filename, proj)
        return filename
    else:
        raise Exception ("Couldn't download file for proj: %d" % (proj))

'''
def combine_problems (prdataset, target):
    SEP = "$;"
    columns = ["#", "Issue Code", "Assignee", "Last Process User", "Pillar", "Subsystem", "APK", "Subject", "State", "PR Due Date Initial", "PR Due Date Revised"]

    all_prs = []
    
    for prdata in prdataset:
        book = xlrd.open_workbook (prdata)
        sheet = book.sheets() [0]

        colmap = {}

        for colid in range (0, sheet.ncols):
            header = sheet.cell_value (0, colid).strip ()
            if header in columns:
                colmap [header] = colid

        for rowid in range (1, sheet.nrows):
            prval = {}
            for colkey in columns:
                if sheet.cell_type (rowid, colmap [colkey]) == xlrd.XL_CELL_NUMBER:
                    prval [colkey] = int (sheet.cell_value (rowid, colmap [colkey]))
                else:
                    prval [colkey] = unicode (sheet.cell_value (rowid, colmap [colkey]))
            all_prs.append (prval)

    tf = open (target, "wb")

    tf.write (SEP.join (columns) + "\n")

    for r in all_prs:
        for c in columns:
            if c == "#":
                tf.write ("%d" % (r [c]))
            else:
                if len (r [c]) != 0:
                    tf.write (r [c].encode ("utf-8"))
                else:
                    tf.write (" ")
            tf.write (SEP)
        tf.write ("\n");

def file_download (directory,config):

    if os.path.exists(directory):
        shutil.rmtree(directory)
    else:       
        # 使用 try 建立目錄
        try:
            os.makedirs(directory)
        # 檔案已存在的例外處理
        except FileExistsError:
            print("目錄已存在。")    
    
    config = TCA.read_config (config)
	
    #CURRENT_PACKAGE_DIRECTORY = os.path.abspath('.')   
    #selenium_driver_chrome = CURRENT_PACKAGE_DIRECTORY + '\chromedriver.exe'     	

    options = webdriver.ChromeOptions()
    prefs = {'profile.default_content_settings.popups': 0, 'download.default_directory': directory}
    options.add_experimental_option('prefs', prefs)
    #br = webdriver.Chrome(executable_path='D:\Python\Python37-32\chromedriver.exe', chrome_options=options)
    if (os.path.isfile ('c:\chromedriver.exe')):
        print('c:\chromedriver.exe')
        br = webdriver.Chrome(executable_path='c:\chromedriver.exe', chrome_options=options)
    else:
        print('False')    
        br = webdriver.Chrome(chrome_options=options)
    
    TCA.login (br, config)
    #time.sleep(30) 
    download_result = download_TCA(br,PROBLEM,directory,config)  
    if (download_result==False):
        download_file = False
    else:
        download_file = TCA_get_filename(directory)    
    time.sleep(3) 
    br.quit ()
    return download_file

def filter_fault_description(df_in):
    df_in = df_in.str.replace("\n"," ")
    df_in = df_in.str.replace("\r"," ")
    df_in = df_in.str.replace("<br/>"," ")
    df_in = df_in.str.replace("&"," ")
    
    find_index = df_in.str.find("[symptom description]")
    
    n=0 
    for i in df_in:    
        if (find_index[n]>0):
            df_in[n]=df_in[n][find_index[n]+22:]
        n=n+1
            
    find_index2 = df_in.str.find("[")
    n=0 
    for i in df_in:       
        if (find_index[n]>=0):
            df_in[n]=df_in[n][:find_index2[n]]
            print('new_str=',df_in[n])
        n=n+1
    
    return df_in           

def filter_Workshop_Comment(df_in):
    df_in = df_in.str.replace("\n"," ")
    df_in = df_in.str.replace("\r"," ")
    df_in = df_in.str.replace("<br/>"," ")
    
    find_index = df_in.str.find("workdscr")

    n=0 
    for i in df_in:    
        if (find_index[n]>=0):
            df_in[n]=df_in[n][find_index[n]+9:]
        n=n+1
    
    return df_in       

def TCA_is_analysis_file_exist():
    now = datetime.now()
    date= now.strftime("%Y")+now.strftime("%m")+now.strftime("%d") 
    
    CURRENT_PACKAGE_DIRECTORY = os.path.abspath('.')   
    Backup_DIRECTORY = CURRENT_PACKAGE_DIRECTORY + '\\backup'     
    #analysis_filename ='TCA_analysis.xlsx'    
    analysis_filename ='TCA_analysis_'+date+'.xlsx'
    
    analysis_filename=os.path.join(Backup_DIRECTORY,analysis_filename)
    print("analysis_filename=",analysis_filename)
    
    is_exist=os.path.exists (analysis_filename) 
    if (is_exist==True):
        return analysis_filename    
    else:
        return False    
    
def TCA_read_TCA_analysis(analysis_filename):

    df = pd.read_excel(analysis_filename)
    #print("df=",df)    
    #sel = input("pause")    
    return df      
    
def TCA_check_symptom(df_in):

    symptom_keyword_df=download_from_google('TCA2','symptom keyword')    

    symptom_keyword_df2 = symptom_keyword_df["symptom keyword"]
    symptom_df = symptom_keyword_df["symptom"]
    symptom_out_df = df_in.copy() #複製dataframe

    #find_index = df_in.str.find("audio")
    
    n=0 
    for i in symptom_out_df:         
        symptom_out_df[n] = ""
        n=n+1  
    
    for i in range(0, len(df_in)):  
        print("df_in=",df_in[i])
        for j in range(0, len(symptom_keyword_df)):
            srt = symptom_keyword_df2[j]
            srt=str(srt)
            srt="*"+srt+"*"
            #print("srt=",srt)
            #print("m,n=",j,i)
            #if (df_in[i].find(srt)>=0):
            if (fnmatch(df_in[i], srt)):
                print("j=",j)
                print("symptom_df=",symptom_df[j])
                symptom_out_df[i]=symptom_df[j]
                #sel = input("pause")
                break

    return symptom_out_df 
    
def TCA_check_Workshop_Comment(df_in):
    comment_keyword_df=download_from_google('TCA2','workshop comment keyword')    
    comment_keyword_df2 = comment_keyword_df["comment keyword"]
    check_tca_df = df_in['check tca'].copy() #複製dataframe
    comment_df=df_in['Workshop Comment']
   
    for i in range(0, len(df_in)):  
        for j in range(0, len(comment_keyword_df)):
            srt = comment_keyword_df2[j]
            srt=str(srt)
            comment_str=str(comment_df[i])
            if (comment_str.find(srt)>=0):
                check_tca_df[i]="Comment"
                break

    df_in["check tca"]=check_tca_df             
    return df_in     
    
def TCA_check_not_SW_symptom_code(df_in):    
    not_sw_df=download_from_google('TCA2','Not SW Symptom Code')    
    comment_keyword_df2 = not_sw_df["not sw"]
    check_tca_df = df_in['check tca'].copy() #複製dataframe
    symptom_code_df=df_in['Symptom Code 1']
   
    for i in range(0, len(df_in)):  
        for j in range(0, len(not_sw_df)):
            srt = comment_keyword_df2[j]
            srt=str(srt)
            comment_str=str(symptom_code_df[i])
            if (comment_str.find(srt)>=0):
                check_tca_df[i]="Symptom"
                break

    df_in["check tca"]=check_tca_df             
    return df_in      

def TCA_check_tca_in_symptom(df_in):

    symptom_df = df_in["symptom"]
    check_tca_df = df_in["symptom"].copy()
 
    n=0 
    for i in check_tca_df:         
        check_tca_df[n] = ""
        n=n+1  
        
    mask = symptom_df.str.contains("ambilight")
    
    for i in range(0, len(mask)):  
        if (mask[i]==True):
            check_tca_df[i]="Fault Description"
 
    mask = symptom_df.str.contains("ci")
    
    for i in range(0, len(mask)):  
        if (mask[i]==True):
            check_tca_df[i]="Fault Description"

    mask = symptom_df.str.contains("lvds")
    
    for i in range(0, len(mask)):  
        if (mask[i]==True):
            check_tca_df[i]="Fault Description"
            
    mask = symptom_df.str.contains("me")
    
    for i in range(0, len(mask)):  
        if (mask[i]==True):
            check_tca_df[i]="Fault Description"       
            
    mask = symptom_df.str.contains("nff")
    
    for i in range(0, len(mask)):  
        if (mask[i]==True):
            check_tca_df[i]="Fault Description"

    mask = symptom_df.str.contains("panel")
    
    for i in range(0, len(mask)):  
        if (mask[i]==True):
            check_tca_df[i]="Fault Description"

    #df_in["check tca"]=check_tca_df
    return check_tca_df   
    
def main (args):
    global log
    config = 'config'
    print('args = ',args)
    projdata = []
    
    try:
        if (args[1]=="debug"):
           log = log_print
    except:
        log = Emptyprintf

    try:
        if (args[1]=="confirm"):
            TCA_confirm_all(config)
            return False    
    except:
        print("")
        
    #sel = input("pause")
                
    now = datetime.now()
    date= now.strftime("%Y")+now.strftime("%m")+now.strftime("%d")    
    time_begin = now - timedelta(days=60)  # time_begin
    time_end = now - timedelta(days=1)  # time_begin
    #time_end = time_end.date()
    
    CURRENT_PACKAGE_DIRECTORY = os.path.abspath('.')    
    PACKAGE_DIRECTORY = CURRENT_PACKAGE_DIRECTORY + '\download' 
    Backup_DIRECTORY = CURRENT_PACKAGE_DIRECTORY + '\\backup' 
	
    if os.path.exists(Backup_DIRECTORY):
        print("backup目錄已存在。")   	
    else:       
        # 使用 try 建立目錄
        try:
            os.makedirs(Backup_DIRECTORY)
        # 檔案已存在的例外處理
        except FileExistsError:
            print("backup目錄已存在目錄已存在。")   	
    
    download_file=False
    is_analysis_file_exist=TCA_is_analysis_file_exist()
    if (is_analysis_file_exist==False):
        download_file = file_download(PACKAGE_DIRECTORY,config) #下載TCA檔案
    else:    
        download_file=True
    
    #sel = input("pause-680")    
    
    if (download_file==False):
        print("=================No claim in TCA=================")
        return False    

    current_df=TCA_download_from_google('TCA2','raw data2') 
    TCA_upload_to_google('TCA2','raw data2_backup',current_df) # upload to google sheet

    df_zero=current_df.copy()
    for col in df_zero.columns: 
        df_zero[col]="" 
    TCA_upload_to_google('TCA2','raw data2',df_zero) # upload to google sheet    

    current_df['Claim RD Review Date']=pd.to_datetime(current_df['Claim RD Review Date'])
    log("current_df=",current_df['Claim RD Review Date'])    
    log("time_end = ",time_end)    
    current2_df = current_df[ current_df['Claim RD Review Date'] <= time_end]
    log("TCA_read2_df=",current2_df['Claim RD Review Date'])   
    
    if (is_analysis_file_exist==False):
        download_file = TCA_get_filename(PACKAGE_DIRECTORY)    
        TCA_read_df = TCA_read(download_file)       
    
    ###翻譯================================================
    if (is_analysis_file_exist==False):
        #print("Workshop Comment = ",TCA_read_df['Workshop Comment'])
        TCA_read_df['Workshop Comment'] = TCA_read_df['Workshop Comment'].str.lower() #全部字串都先轉成小寫
        TCA_read_df['Fault Description Text'] = TCA_read_df['Fault Description Text'].str.lower() #全部字串都先轉成小寫
        
        Workshop_Comment_df=TCA_read_df['Workshop Comment']
        Workshop_Comment_df = filter_Workshop_Comment(Workshop_Comment_df)  
        Workshop_Comment_df = Google_Translator("auto","en",Workshop_Comment_df)   
        TCA_read_df['Workshop Comment']=Workshop_Comment_df      
        
        Description_df=TCA_read_df['Fault Description Text']
        Description_df = filter_fault_description(Description_df)    
        Description_df = Google_Translator("auto","en",Description_df)   
        TCA_read_df['Fault Description Text']=Description_df 
    else:         
        TCA_read_df = TCA_read_TCA_analysis(is_analysis_file_exist)
        TCA_read_df = TCA_read_df.drop(['symptom', 'check tca'], axis=1)
    #TCA_read_df['Fault Description Text'] = read_excel_df['Fault Description Text']
    #TCA_read_df['Workshop Comment'] = read_excel_df['Workshop Comment']
    ###===============================================================
    ###判斷symptom====================================================
    out_df=TCA_check_symptom(TCA_read_df['Fault Description Text'])
    TCA_read_df.insert(TCA_read_df.shape[1],'symptom',out_df)
    ###===============================================================
    ###判斷check tca====================================================
    out_df=TCA_check_tca_in_symptom(TCA_read_df)   
    TCA_read_df.insert(TCA_read_df.shape[1],'check tca',out_df)   
    TCA_read_df=TCA_check_not_SW_symptom_code(TCA_read_df)       
    TCA_read_df=TCA_check_Workshop_Comment(TCA_read_df)           
    #sel = input("pause==================================")        
    ###===============================================================    
    analysis_filename ='TCA_analysis_'+date+'.xlsx'
    print("analysis_filename=",analysis_filename)
    analysis_filename=os.path.join(Backup_DIRECTORY,analysis_filename)
    TCA_read_df.to_excel(analysis_filename,index=False)   
    
    TCA_new_df=pd.concat([current2_df,TCA_read_df],ignore_index=True)   
    TCA_new_df=TCA_new_df.fillna("")    
    #transfer date format
    TCA_new_df['Claim RD Review Date']=pd.to_datetime(TCA_new_df['Claim RD Review Date'])
    log(TCA_new_df['Claim RD Review Date'])
    
    log(TCA_new_df.info())
    #delete old data
    log('type=',type(TCA_new_df['Claim RD Review Date'][1:1]))
    TCA_new2_df = TCA_new_df[ TCA_new_df['Claim RD Review Date'] > time_begin]
    log(TCA_new2_df['Claim RD Review Date'])
    TCA_new2_df=TCA_new2_df.fillna("")   
    #pdb.set_trace() 
    #TCA_upload_to_google('TCA2','raw data2',df_zero) # upload to google sheet
    #pdb.set_trace() 
    TCA_upload_to_google('TCA2','raw data2',TCA_new2_df) # upload to google sheet
    
    if (is_analysis_file_exist==False):
        TCA_backup(download_file)
    print("==============Done==============")
        
def TCA_confirm_all(config):
    print("==============TCA_confirm_all==============")
    config = TCA.read_config (config)

    options = webdriver.ChromeOptions()
    #prefs = {'profile.default_content_settings.popups': 0, 'download.default_directory': directory}
    #options.add_experimental_option('prefs', prefs)
    #br = webdriver.Chrome(executable_path='D:\Python\Python37-32\chromedriver.exe', chrome_options=options)
    if (os.path.isfile ('c:\chromedriver.exe')):
        print('c:\chromedriver.exe')
        br = webdriver.Chrome(executable_path='c:\chromedriver.exe', chrome_options=options)
    else:
        print('False')    
        br = webdriver.Chrome(chrome_options=options)

    TCA.login (br, config) #log in

    button = waiting_for_TCA_update(br,config ['xpath_cashout'])
    time.sleep(5)       
    button.click() #click cashout   
    button = waiting_for_TCA_update(br,config ['xpath_domain'])
    button.click() #click domain   
    print('wait 5 seconds................')
    time.sleep(5)          
    button = waiting_for_TCA_update(br,config ['xpath_domain_iframe'])
    print(button.text)
    print('wait 5 seconds................')
    time.sleep(5)           
    #sel = input("pause-792")    
    br.switch_to.frame(1)
    print('wait 5 seconds................')
    time.sleep(5)       
    button = waiting_for_TCA_update(br,config ['xpath_claims'])
    #sel = input("pause-796")     
    print('claims=',button.text)
    #sel = input("pause-798")    
    #br.switch_to.frame(1)
    
    #br.get ('https://gsp.tpv-tech.com/RedirectURL.aspx?pg=c&srcgo=Cashout%2fGCS_HomepageForRDDomain.aspx')
    #button = br.find_element_by_xpath('//*[@id="aRDDomain"]')
    log('button=',button)
    if (int(button.text)==0):
        return False
    
    button.click() #click 數量     
   
    button = waiting_for_TCA_update(br,config ['xpath_download_all'])
    log('download all=',button)
    if (button==False):
        return False
        
    button = waiting_for_TCA_update(br,config ['xpath_row1']) # check row 1 exist
    log('row 1=',button)
    if (button==False):
        return False      
    
    while (button!=False):      
        TCA_click_select_all(br,config)        

        button = waiting_for_TCA_update(br,config ['xpath_row1']) # check row 1 exist
        log('row 1=',button)
        
        if (button.is_selected()):
            print("click confirm")
        else:
            TCA_click_select_all(br,config)        
            
        
        button = waiting_for_TCA_update(br,config ['xpath_confirm']) # btnConfirm  
        log('btnConfirm=',button)   
        #sel = input("pause")        
        button.click() #click confirm 
        time.sleep (5)         
        button = waiting_for_TCA_update(br,config ['xpath_row1']) # check row 1 exist 
            
    br.quit ()
    print("==============TCA_confirm_all_finish==============")
    return True

def TCA_click_select_all(br,config):
    button = waiting_for_TCA_update(br,config ['xpath_select_all']) # select all
    log('select all=',button)
    if (button==False):
        return False  
    button.click() #click select all           

    if (button.is_selected()):
        log("selected")
    else:
        button.click() #click select all  
    
        
if __name__ == '__main__':
    main (sys.argv)
