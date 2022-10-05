import os
import re
import time
import xlrd
from selenium import webdriver
from time import sleep
import tca as TCA
import inspect
#import pyperclip as pc


#GT_URL="https://translate.google.com/"
GT_URL="https://translate.google.com/?sl=auto&tl=en&text="

def printLineFileFunc():
    callerframerecord = inspect.stack()[1]   
    frame = callerframerecord[0]
    info = inspect.getframeinfo(frame)
    print ("File = ",info.filename,",Function = ",info.function,",Line = ",info.lineno)
    #print (info.filename)                      # __FILE__    
    #print (info.function)                       # __FUNCTION__
    #print (info.lineno)                          # __LINE__


def waiting_for_TCA_update(br,xpath):
    print("waiting_for_TCA_update button start")
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

def Translator (br,data_in):
    GT_URL_Text = GT_URL+data_in
    br.get(GT_URL_Text) 
    
    config = 'config'    
    config = TCA.read_config (config)
    
    button = False
    
    #sel = input("pause 58")    
    
    for i in range(0,10):
        time.sleep(1) 
        print ('waiting for update Count: %d'% i)
        printLineFileFunc()
        try:
            button = br.find_element_by_xpath(config ['xpath_translate'])
                                               
        except:
            continue
        break
        
    return button

def Google_Translator(src_in,dest_in,df_in):
    print(df_in)
    options = webdriver.ChromeOptions()
    prefs = {'profile.default_content_settings.popups': 0, 'download.default_directory': 'd:\\'}
    options.add_experimental_option('prefs', prefs)

    if (os.path.isfile ('c:\chromedriver.exe')):
        br = webdriver.Chrome(executable_path='c:\chromedriver.exe', chrome_options=options)
    else:
        br = webdriver.Chrome(chrome_options=options)
        
    #br.get('https://translate.google.com/')

    # src來源語言，dest要翻譯語言，如果要找其他語言可以參考說明文件
    n=0 
    str="sorun : tv açılmamaktadır cevap : servise yönlendirildi"   
    
    GT_URL_Text = GT_URL+str
    print(GT_URL_Text)
    br.get(GT_URL)
    time.sleep (5)

    #button = br.find_element_by_xpath('//*[@id="yDmH0d"]/c-wiz/div/div[2]/c-wiz/div[2]/c-wiz/div[1]/div[2]/div[2]/c-wiz[2]/div[5]/div/div[1]').text 
    #print(button)    

    for i in df_in:         
        #if (isNaN(df_in[n])==False): 
        try:
            str_len=len(df_in[n])
        except:  
            str_len=0
        
        if (str_len>1):
            print ('i= %d' % n)
            print(df_in[n])  

            str_out = Translator(br,df_in[n])
            if (str_out!=False):    
                print ('n_out= %d' % n)
                print(str_out.text)  
                df_in[n] = str_out.text
        n=n+1    
    
    '''
    button = br.find_element_by_xpath('/html/body/div[2]/div[1]/div[2]/div[1]/div[1]/div[1]/div[1]/div[4]/div[1]/div[2]/div[2]') 
    button.click() #click 英文        
    time.sleep (1)    
    
    str_out = Translator(br,str)
    print(str_out.text)
    for i in df_in:         
        #if (isNaN(df_in[n])==False): 
        try:
            str_len=len(df_in[n])
        except:  
            str_len=0
        
        if (str_len>1):
            print ('i= %d' % n)
            print(df_in[n])  
            try:
                str_out = Translator(br,df_in[n])
            except:    
                n=n;            
            print ('n_out= %d' % n)
            print(str_out.text)  
            df_in[n] = str_out.text
        n=n+1    
    '''
    br.quit ()
    
    return df_in


'''  
def Google_Translator(src_in,dest_in,df_in):
    
    #df_in=df_in.fillna("")
    #print("df_in=",df_in)
    translator = Translator()
    # src來源語言，dest要翻譯語言，如果要找其他語言可以參考說明文件
    n=0 
    str="sorun : tv açılmamaktadır cevap : servise yönlendirildi"   
    #str_out=translator.translate(str, src = src_in, dest = dest_in)
    #str_out=translator.translate(str, dest="en")
    #print(str_out.text)
    translated_flag=False
    for i in df_in:         
        #if (isNaN(df_in[n])==False):   
        if (len(df_in[n])>1):  
            print(df_in[n])  
            
            while (translated_flag==False):
                try:
                    #str_out=translator.translate(df_in[n], src = src_in, dest = dest_in)
                    str_out=translator.translate(df_in[n], dest = dest_in)
                    df_in[n] = str_out.text
                    translated_flag=True
                except:    
                    translated_flag=False  
                    print("Retry...")                    
                    time.sleep (1) 

                   

            #df_in[n] = str_out.text
            n=n+1    
        
        print ("n=",n)
        
    #print(str_out.text)
    #sel = input("pause")
     
    return df_in
 ''' 
 
