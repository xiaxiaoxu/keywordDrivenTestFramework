#encoding=utf-8
from selenium import webdriver
import time
from Util.ObjectMap import *
from ProjectVar.Var import *
import traceback
from selenium.webdriver.chrome.options import Options
from Util.FormatTime import *
from Util.DirandFile import *
#定义全局的浏览器driver变量
driver=None

def open_browser(browserName,*arg):
    global driver
    try:
        if browserName.lower().strip() == 'ie':
            driver = webdriver.Ie(executable_path = ieDriverFilePath)
        elif browserName.lower().strip() == 'chrome':

            driver = webdriver.Chrome(
                executable_path = chromeDriverFilePath,
            )
        else:
            driver = webdriver.Firefox(executable_path = firefoxDriverFilePath)
    except Exception,e:
        raise e

def visit_url(url,*arg):
    global driver
    try:
        driver.get(url)
    except Exception,e:
        raise e

def pause(seconds,*arg):
    time.sleep(float(seconds))

def close_browser(*arg):
    global driver
    try:
        driver.quit()
    except Exception,e:
        raise e

def enter_frame(locatorMethod,locatorExpression,*arg):
    global driver
    try:
        driver.switch_to.frame(getElement(driver,locatorMethod,locatorExpression))
    except Exception,e:
        raise e
        print "can not enter frame!"

def input_string(locatorMethod,locatorExpression,content,*arg):
    try:
        getElement(driver, locatorMethod, locatorExpression).clear()
        getElement(driver, locatorMethod, locatorExpression).send_keys(content)
    except Exception,e:
        raise e

def click(locatorMethod,locatorExpression,*arg):
    try:
        getElement(driver, locatorMethod, locatorExpression).click()
    except Exception,e:
        raise e

def login(usernameAndpassword,*arg):
    username,password=usernameAndpassword.split("||")
    open_browser("ie")
    visit_url("http://mail.126.com")
    pause(3)
    enter_frame("id", "x-URS-iframe")
    pause(2)
    input_string("xpath", "//input[@name='email']", username)
    input_string("xpath", "//input[@name='password']", password)
    pause(3)
    click("id", "dologin")
    pause(3)
    assert_word(u"退出")

def assert_word(expected_word,*arg):
    try:
        assert  True == (expected_word in driver.page_source)
    except AssertionError,e:
        raise  e
    except Exception,e:
        raise e

def capture_screen():
    global driver
    createDir(project_path + "\\ScreenPictures\\CapturePicture\\", dates())
    filename="%s\\ScreenPictures\\CapturePicture\\%s\\%s.png" %(project_path,dates(),times())
    print filename
    try:
        driver.get_screenshot_as_file(filename)
    except Exception,e:
        raise e
    return filename

def capture_error_screen():
    global driver
    createDir(project_path + "\\ScreenPictures\\ErrorPicture\\", dates())
    filename="%s\\ScreenPictures\\ErrorPicture\\%s\\%s.jpg" %(project_path,dates(),times())
    print filename
    try:
        driver.get_screenshot_as_file(filename)
    except Exception,e:
        raise e
    return filename

def add_contact_info(name,email,mobile,other_info):
    pause(5)
    # 点击“通讯录”按钮
    click("xpath", "//div[text()='通讯录']")
    pause(2)
    # 点击“新建联系人”按钮
    click("xpath", "//span[text()='新建联系人']")
    pause(2)
    # 输入联系人姓名
    input_string("xpath", "//a[@title='编辑详细姓名']/preceding-sibling::div/input",name)
    # 输入联系人电子邮箱
    input_string("xpath", "//*[@id='iaddress_MAIL_wrap']//input", email)

    click("xpath", "//span[text()='设为星标联系人']/preceding-sibling::span/b")
    pause(2)
    # 输入联系人手机号
    input_string("xpath", "//*[@id='iaddress_TEL_wrap']//dd//input", mobile)
    pause(2)
    # 输入备注信息
    input_string("xpath", "//textarea", other_info)
    pause(2)
    # 点击“确认”按钮
    click("xpath", "//span[text()='确 定']")
    pause(5)


if __name__=="__main__":
    '''
    open_browser("ie")
    visit_url("http://mail.126.com")
    pause(3)
    capture_screen()
    #driver.get_screenshot_as_file(ur"E:\KeyWordFrameWork_wulaoshi\ScreenPictures\CapturePicture\17时42分55秒.jpg")
    close_browser()
'''
    '''
    open_browser("ie")
    visit_url("http://mail.126.com")
    pause(3)
    enter_frame("id","x-URS-iframe")
    pause(2)
    input_string("xpath","//input[@name='email']","testman1980")
    input_string("xpath", "//input[@name='password']", "wulaoshi1978")
    pause(3)
    click("id","dologin")
    pause(3)
    #close_browser()


    pause(5)
    # 点击“通讯录”按钮
    click("xpath","//div[text()='通讯录']")
    pause(2)
    # 点击“新建联系人”按钮
    click("xpath", "//span[text()='新建联系人']")
    pause(2)
    # 输入联系人姓名
    input_string("xpath","//a[@title='编辑详细姓名']/preceding-sibling::div/input","xxx")
    # 输入联系人电子邮箱
    input_string("xpath", "//*[@id='iaddress_MAIL_wrap']//input", "2055739@qq.com")

    click("xpath","//span[text()='设为星标联系人']/preceding-sibling::span/b")
    pause(2)
    # 输入联系人手机号
    input_string("xpath", "//*[@id='iaddress_TEL_wrap']//dd//input", "135xxxxxxx")
    pause(2)
    # 输入备注信息
    input_string("xpath", "//textarea", u"朋友")
    pause(2)
    # 点击“确认”按钮
    click("xpath", "//span[text()='确 定']")
    pause(5)
'''


    login_info=("testman1980||wulaoshi1978","testman1981||wulaoshi1978")
    contact_info=[("g1","2ds30033@qq.com","1393333333333",u"朋友"),("g2","2ds30033@qq.com","1393333333333",u"朋友")]
    for i in login_info:
        login(i)
        for j in  contact_info:
            add_contact_info(j[0], j[1], j[2], j[3])
    # 等待10秒,以便登录成功后的页面加载完成
    #add_contact_info("gloryroad","2ds30033@qq.com","1393333333333",u"朋友")
    driver.quit()
