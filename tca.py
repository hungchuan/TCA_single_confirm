#!/usr/bin/env python

# encoding: utf-8
#
# Copyright 2011, Apurva Sharan
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#       http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.

import os
import re
import time
import xlrd
import SliderVerificationCode as Slider
import pyautogui

#import Slider

TCA_URL="https://gsp.tpv-tech.com"

def login (br, config):

    print ('Entering Login')
    br.get (TCA_URL + "/Login.aspx")
    br.set_window_size(1280, 800)
    elem_user = br.find_element_by_xpath(config ['xpath_username'])
    elem_user.send_keys(config ['username']) #瀏覽器版本不匹配的時候這裡可能報錯
    elem_pwd = br.find_element_by_xpath(config ['xpath_password'])
    elem_pwd.send_keys(config ['password'])
    button = br.find_element_by_xpath(config ['xpath_log_in'])   
    button.click()
    time.sleep (3)
    
    done = False
    ReTry = 0
    while (False == done) and (ReTry<10):
        time.sleep (3)
        ReTry=ReTry+1
        Verify = Slider.SliderVerificationCode()
        Verify.set_broswer(br)
        Verify.verify()    
        done = wait_for_xpath(br,config ['xpath_cashout'])
        if (False == done):
            pyautogui.press('enter')
           
   
    print ('Leaving Login')


def read_config (config):
    print('read_config')
    retmap = {}
    configs = open (config).readlines ()
    print('configs = ',configs)
    for config in configs:
        (k,v) = config.split ("==")
        if "USERNAME" == k.strip ():
            retmap ['username'] = v.strip ()
        if "PASSWORD" == k.strip ():
            retmap ['password'] = v.strip ()
        if "TIMEDELTA" == k.strip ():
            retmap ['timedelta'] = v.strip ()            
        if "MOZILLA_PROFILE_PATH" == k.strip ():
            retmap ['mozilla_profile_path'] = v.strip ()
        if "PROJECT_IDS" == k.strip ():
            retmap ['project_ids'] = v.strip ()
        if "XPath_username" == k.strip ():
            retmap ['xpath_username'] = v.strip ()            
        if "XPath_password" == k.strip ():
            retmap ['xpath_password'] = v.strip ()             
        if "XPath_log_in" == k.strip ():
            retmap ['xpath_log_in'] = v.strip ()             
        if "XPath_cashout" == k.strip ():
            retmap ['xpath_cashout'] = v.strip ()                         
        if "XPath_domain" == k.strip ():
            retmap ['xpath_domain'] = v.strip ()                         
        if "XPath_domain_iframe" == k.strip ():
            retmap ['xpath_domain_iframe'] = v.strip ()             
        if "XPath_claims" == k.strip ():
            retmap ['xpath_claims'] = v.strip ()          
        if "XPath_download_all" == k.strip ():
            retmap ['xpath_download_all'] = v.strip ()    
        if "XPath_yes" == k.strip ():
            retmap ['xpath_yes'] = v.strip ()    
        if "XPath_row1" == k.strip ():
           retmap ['xpath_row1'] = v.strip ()    
        if "XPath_select_all" == k.strip ():
           retmap ['xpath_select_all'] = v.strip ()    
        if "XPath_confirm" == k.strip ():
           retmap ['xpath_confirm'] = v.strip ()    
        if "XPath_translate" == k.strip ():
           retmap ['xpath_translate'] = v.strip ()                
    return retmap
    

def xpath_select_dropdown (br, xpath, text):
    for e in br.find_by_xpath (xpath):
        if re.search (text, e.text) != None:
            e.click ()
            return
    raise Exception ("Couldn't find required element: %s" % (text))

def xpath_click_first_checkbox (br, xpath):
    br.find_by_xpath (xpath).first.click ()

'''
def wait_for_xpath (br, xpath):
    done = False
    count = 0
    while (done == False):
        try:
            if (len (br.find_by_xpath (xpath)) != 0):
                done = True
        except e:
            print ('Retrying for xpath %s .. count : %d' % (xpath, count))
            count += 1
            time.sleep (1)
'''

def wait_for_xpath(br,xpath):
    done = False
    for i in range(0,10):
        time.sleep(1) 
        print ('waiting for update Count: %d' % i)
        try:
            done = br.find_element_by_xpath(xpath)
        except:
            continue
        break
    print(xpath)
    print("waiting_for_TCA_update button=",done)
    return done  


def wait_for_update_progress (br, xpath):
    # Ensure that update spinner shows up and goes off
    datafill = br.find_by_xpath (xpath)
    while (datafill.visible == False): time.sleep (1)
    while (datafill.visible == True): time.sleep (1)


def ensure_file_doesnt_exist (filename):
    if (os.path.exists (filename) == True):
        os.unlink (filename)


def complete_download (filename, proj):
    #PACKAGE_DIRECTORY = os.path.abspath('.')
    
    #filename= os.path.join(PACKAGE_DIRECTORY,filename)   
    #print('filename =',filename)    
    
    filename = os.path.expanduser (filename)
    print('filename = ',filename)
    
    count = 0
       
    while (os.path.exists (filename) == False):
        print ('Count 1: %d' % count)
        count += 1
        time.sleep (1)

    while (os.path.exists (filename + ".part") == True):
        print ('Count 2: %d' % count)
        count += 1
        time.sleep (1)

    while (os.path.getsize (filename) == 0):
        print ('Count 3: %d' % count)
        count += 1
        time.sleep (1)

    # download complete- rename the file and return the new name
    target = "%s.%d" % (filename, proj)
    os.rename (filename, target)
    return target


