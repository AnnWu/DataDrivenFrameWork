# -*- coding: utf-8 -*-
# created by wmin3
from util.ObjectMap import *
from  util.KeyBoardUtil import KeyBoardKeys
from util.ClipboardUtil import Clipboard
from  util.WaitUtil import WaitUtil
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
from action.PageAction import *

def TestSendMailWithAttachment():

    #driver = webdriver.Firefox(executable_path="C:\\WebDriver\\geckodriver")
    #driver.maximize_window()
    open_browser("firefox")
    maximize_browser()

    visit_url("http://mail.126.com") #driver.get("http://mail.126.com")
    time.sleep(5)
    assert_string_in_pagesource(u"126网易免费邮")#assert u"126网易免费邮" in driver.page_source

    #wait = WaitUtil(driver)
    #wait.frameToBeAvailableAndSwitchToIt('xpath',"//iframe[@name='']")
    waitFrameToBeAvailableAndSwitchToIt('xpath',"//iframe[@name='']")

    #username = getElement(driver,"xpath","//input[@name='email']")
    #username.send_keys("minwu126")
    input_string("xpath","//input[@name='email']","minwu126")

    #passwd = getElement(driver, "xpath", "//input[@name='password']")
    #passwd.send_keys("wuli8228680")
    #passwd.send_keys(Keys.ENTER)
    input_string("xpath", "//input[@name='password']","wuli8228680")
    click("id","dologin")

    time.sleep(15)
    assert_title(u"网易邮箱") #assert u"网易邮箱" in driver.title

    #driver.switch_to.default_content()
    switch_to_default_content()
    time.sleep(1)

    #addressBook = wait.visibilityOfElementLocated("xpath","//div[text()='通讯录']")
    #addressBook.click()
    waitVisibilityOfElementLocated("xpath", "//div[text()='通讯录']")#.click()
    click("xpath", "//div[text()='通讯录']")
    #newContact = wait.visibilityOfElementLocated("xpath","//span[text()='新建联系人']")
    #newContact.click()
    waitVisibilityOfElementLocated("xpath", "//span[text()='新建联系人']")#.click()
    click("xpath", "//span[text()='新建联系人']")

    #contactName = wait.visibilityOfElementLocated("xpath","//a[@title='编辑详细姓名']/preceding-sibling::div/input")
    #contactName.send_keys("LILY")
    input_string("xpath", "//a[@title='编辑详细姓名']/preceding-sibling::div/input", "LIYLY")

    #email = getElement(driver,"xpath","//*[@id='iaddress_MAIL_wrap']//input")
    #email.send_keys("lily@qq.com")
    input_string("xpath", "//*[@id='iaddress_MAIL_wrap']//input", "lily@qq.com")

    #getElement(driver,"xpath","//span[text()='设为星标联系人']/preceding-sibling::span/b").click()
    click("xpath","//span[text()='设为星标联系人']/preceding-sibling::span/b")

    #mobile=getElement(driver,"xpath","//*[@id='iaddress_TEL_wrap']//dd//input")
    #mobile.send_keys("18888888891")
    input_string("xpath", "//*[@id='iaddress_TEL_wrap']//dd//input", "18888888855")

    #getElement(driver,"xpath","//textarea").send_keys(u"朋友") #备注
    input_string("xpath", "//textarea", u"朋友")

    #getElement(driver,"xpath","//span[text()='确 定']").click()
    click("xpath","//span[text()='确 定']")
    time.sleep(2)

    assert_string_in_pagesource(u"lily@qq.com")#assert u"lily@qq.com" in driver.page_source
    print(u"添加成功")
    time.sleep(3)
    #getElement(driver,"xpath","//div[.='首页']").click()
    click("xpath","//div[.='首页']")
    time.sleep(2)

    #element = wait.visibilityOfElementLocated("xpath","//span[text()='写 信']")
    #element.click()
    waitVisibilityOfElementLocated("xpath", "//span[text()='写 信']")#.click()
    click("xpath", "//span[text()='写 信']")
    print(u"写信")

    #receiver = getElement(driver,"xpath","//div[contains(@id,'_mail_emailinput')]//input")
    #receiver.send_keys("757693255@qq.com")
    input_string("xpath","//div[contains(@id,'_mail_emailinput')]//input","757693255@qq.com")

    #subject = getElement(driver,"xpath","//div[@aria-label='邮件主题输入框，请输入邮件主题']/input")
    #subject.send_keys(u"新邮件")
    input_string("xpath", "//div[@aria-label='邮件主题输入框，请输入邮件主题']/input", u"新邮件")

    #attachment=getElement(driver,"xpath","//div[@title='点击添加附件']/input[@size='1' and @type='file']")
    #attachment.send_keys("d:\\a.txt")
    input_string("xpath", "//div[@title='点击添加附件']/input[@size='1' and @type='file']", u"d:\\a.txt")

    time.sleep(5)

    #wait.frameToBeAvailableAndSwitchToIt("xpath","//iframe[@tabindex=1]")
    waitFrameToBeAvailableAndSwitchToIt("xpath", "//iframe[@tabindex=1]")
    #body= getElement(driver,"xpath","/html/body")
    #body.send_keys(u"发给自己的一封信")
    input_string("xpath","/html/body",u"发给自己的一封信")

    #driver.switch_to_default_content()
    switch_to_default_content()

    #getElement(driver,"xpath","//header//span[text()='发送']").click()
    click("xpath","//header//span[text()='发送']")
    print(u"开始发送邮件")
    time.sleep(3)
    assert_string_in_pagesource(u"发送成功")#assert u"发送成功"in driver.page_source
    close_browser() #driver.quit()

if __name__ =="__main__":
    TestSendMailWithAttachment()