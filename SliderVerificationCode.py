# -*- coding:utf-8 -*-
from PIL import Image
from time import sleep
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver import ActionChains
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
import cv2
import numpy
import pyautogui


TCA_URL="https://gsp.tpv-tech.com"

headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/76.0.3809.132 Safari/537.36"
}
chrome_options = webdriver.ChromeOptions()
chrome_options.add_experimental_option('w3c', False)
caps = DesiredCapabilities.CHROME
caps['loggingPrefs'] = {'performance': 'ALL'}

class SliderVerificationCode(object):
    def __init__(self):  # 初始化一些信息
        print('__init__')       
        self.left = 60  # 定義一個左邊的起點 缺口一般離圖片左側有一定的距離 有一個滑塊
        self.url = TCA_URL
        #self.driver = webdriver.Chrome(executable_path='c:\chromedriver.exe',desired_capabilities=caps,options=chrome_options)
        #self.wait = WebDriverWait(self.driver, 20)  # 設置等待時間20秒
        self.phone = "ryan.hung "
        self.passwd = "Aa111111"
        
    def set_broswer(self,br):  
        self.driver = br
    

    def input_name_password(self):  # 輸入賬號密碼
        self.driver.get(self.url)
        #self.driver.maximize_window()
        self.driver.set_window_size(1280, 800)
        
        input_name = self.driver.find_element_by_xpath('//*[@id="name"]')
        input_pwd = self.driver.find_element_by_xpath('//*[@id="password"]')
        input_name.send_keys("ryan.hung")
        input_pwd.send_keys("Aa111111")
        sleep(3)

    def click_login_button(self):  # 點擊登錄按鈕,出現驗證碼圖片
        login_btn = self.driver.find_element_by_xpath('//*[@id="btnlogin"]')
        login_btn.click()
        sleep(3)

    def brightness( im_file ):
       im = Image.open(im_file).convert('L')
       stat = ImageStat.Stat(im)
       return stat.rms[0]

    def get_geetest_image(self):  # 獲取驗證碼圖片    
    
        element = self.driver.find_element_by_xpath('//*[@id="xy_img"]')
        attributeValue = element.get_attribute("style")
        print(attributeValue)
        print(attributeValue.find('top:'))
        print(attributeValue.find('px;'))
        top=attributeValue[attributeValue.find('top:')+5:attributeValue.find('px;')]
        iTop=int(top)+3
        print(iTop)    
        #sel = input("pause top")
        
        #width: 300px; height: 234px;
        self.driver.get_screenshot_as_file(r".\1.png")
        #sel = input("pause")    
        img = Image.open("./1.png")
        cropped = img.crop((531, 197+iTop, 831, 197+iTop+40))  # (left, upper, right, lower)
        cropped.save("./captcha1.png")
        sleep(1)
        #sel = input("pause captcha1")    
        
        img = cv2.imread('./captcha1.png')
        print(img.shape)
        print(type(img))


    def is_similar(self, image1, image2, x, y):
        '''判斷兩張圖片 各個位置的像素是否相同
        #image1:帶缺口的圖片
        :param image2: 不帶缺口的圖片
        :param x: 位置x
        :param y: 位置y
        :return: (x,y)位置的像素是否相同
        '''
        # 獲取兩張圖片指定位置的像素點
        pixel1 = image1.load()[x, y]
        #pixel2 = image2.load()[x, y]
        pixel2 = image2.load()[0, 0]
        # 設置一個閾值 允許有誤差
        threshold = 60
        # 彩色圖 每個位置的像素點有三個通道
        if abs(pixel1[0] - pixel2[0]) < threshold and abs(pixel1[1] - pixel2[1]) < threshold and abs(
                pixel1[2] - pixel2[2]) < threshold:
            return True
        else:
            return False
    '''
    def get_diff_location(self):  # 獲取缺口圖起點
        captcha1 = Image.open('captcha1.png')
        captcha2 = Image.open('captcha2.png')
        for x in range(self.left, captcha1.size[0]):  # 從左到右 x方向
            for y in range(captcha1.size[1]):  # 從上到下 y方向
                if not self.is_similar(captcha1, captcha2, x, y):
                    return x  # 找到缺口的左側邊界 在x方向上的位置
    '''

    def get_diff_location(self):  # 獲取缺口圖起點
        img = cv2.imread('./captcha1.png')
        print(img.shape)
        print(type(img))
        
        contrast = 210
        brightness = -200
        output = img * (contrast/127 + 1) - contrast + brightness # 轉換公式
        # 轉換公式參考 https://stackoverflow.com/questions/50474302/how-do-i-adjust-brightness-contrast-and-vibrance-with-opencv-python

        # 調整後的數值大多為浮點數，且可能會小於 0 或大於 255
        # 為了保持像素色彩區間為 0～255 的整數，所以再使用 np.clip() 和 np.uint8() 進行轉換
        output = numpy.clip(output, 0, 255)
        output = numpy.uint8(output)

        #cv2.imshow('oxxostudio1', img)    # 原始圖片
        #cv2.imshow('oxxostudio2', output) # 調整亮度對比的圖片
        #cv2.waitKey(0)                    # 按下任意鍵停止
        #cv2.destroyAllWindows()      

        Min=255
        pos=0
        for i in range(260):
            cropped2 = output[0:40, 0+i:40+i]
            #cv2.imshow('image', img)
            #cv2.imshow('cropped', cropped2)
            #cv2.waitKey(0)      

            #cols, rows = cropped2.shape
            brightness = numpy.sum(cropped2) / (255 * 40 * 40)   
            if (brightness<Min):
                Min = brightness
                pos = i
            print('i=',i)
            print('brightness = ',brightness)        

        print('pos=',pos)
        #sel = input("pause get_diff_location")    
        return pos


    def get_move_track(self, gap):
        print('gap=',gap)
        track = []  # 移動軌跡
        current = 0  # 當前位移
        # 減速閾值
        mid = gap * 4 / 5  # 前4/5段加速 后1/5段減速
        t = 0.2  # 計算間隔
        v = 0  # 初速度
        while current < gap:
            if current < mid:
                a = 5  # 加速度為+5
            else:
                a = -5  # 加速度為-5
            v0 = v  # 初速度v0
            v = v0 + a * t  # 當前速度
            move = v0 * t + 1 / 2 * a * t * t  # 移動距離
            current += move  # 當前位移
            track.append(round(move))  # 加入軌跡
        return track

    def move_slider(self, track):
        #slider = self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '.handler handler_bg')))
        slider = self.driver.find_element_by_xpath('//*[@id="drag"]/div[3]')
        ActionChains(self.driver).click_and_hold(slider).perform()
        for x in track:  # 只有水平方向有運動 按軌跡移動
            ActionChains(self.driver).move_by_offset(xoffset=x, yoffset=0).perform()
            #sleep(0.2)
        sleep(1)
        ActionChains(self.driver).release().perform()  # 松開鼠標

    def verify(self):
        #self.input_name_password()
        #self.click_login_button()
        self.get_geetest_image()
        gap = self.get_diff_location()  # 缺口左起點位置
        #sel = input("main pause") 
        #pyautogui.press('enter')
        #sleep(2)
        #sel = input("main pause2")   
        #gap = 100
        gap = gap - 4  # 減去滑塊左側距離圖片左側在x方向上的距離 即為滑塊實際要移動的距離
        track = self.get_move_track(gap)
        print('track=',track)
        self.move_slider(track)

if __name__ == "__main__":
    br = webdriver.Chrome(executable_path='c:\chromedriver.exe',desired_capabilities=caps,options=chrome_options)
    springAutumn = SliderVerificationCode()
    springAutumn.set_broswer(br)
    springAutumn.input_name_password()
    springAutumn.click_login_button()    
    springAutumn.verify()