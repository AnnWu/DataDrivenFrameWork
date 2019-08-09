# -*- coding: utf-8 -*-
# created by wmin3
from util.ObjectMap import *
from  util.KeyBoardUtil import KeyBoardKeys
from util.ClipboardUtil import Clipboard
from  util.WaitUtil import WaitUtil
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time

def TestSendMailWithAttachment():

    driver = webdriver.Firefox(executable_path="C:\\WebDriver\\geckodriver")
    driver.maximize_window()

    driver.get("http://mail.126.com")
    time.sleep(5)
    assert u"126网易免费邮" in driver.page_source

    wait = WaitUtil(driver)
    wait.frameToBeAvailableAndSwitchToIt('xpath',"//iframe[@name='']")
    username = getElement(driver,"xpath","//input[@name='email']")
    username.send_keys("minwu126")
    passwd = getElement(driver, "xpath", "//input[@name='password']")
    passwd.send_keys("wuli8228680")
    passwd.send_keys(Keys.ENTER)
    time.sleep(10)
    assert u"网易邮箱" in driver.title

    driver.switch_to.default_content()
    time.sleep(5)

    addressBook = wait.visibilityOfElementLocated("xpath","//div[text()='通讯录']")
    addressBook.click()
    newContact = wait.visibilityOfElementLocated("xpath","//span[text()='新建联系人']")
    newContact.click()

    contactName = wait.visibilityOfElementLocated("xpath","//a[@title='编辑详细姓名']/preceding-sibling::div/input")
    contactName.send_keys("LILY")
    email = getElement(driver,"xpath","//*[@id='iaddress_MAIL_wrap']//input")
    email.send_keys("lily@qq.com")
    getElement(driver,"xpath","//span[text()='设为星标联系人']/preceding-sibling::span/b").click()

    mobile=getElement(driver,"xpath","//*[@id='iaddress_TEL_wrap']//dd//input")
    mobile.send_keys("18888888891")

    getElement(driver,"xpath","//textarea").send_keys(u"朋友") #备注
    getElement(driver,"xpath","//span[text()='确 定']").click()
    time.sleep(2)

    assert u"lily@qq.com" in driver.page_source
    print(u"添加成功")
    time.sleep(3)
    getElement(driver,"xpath","//div[.='首页']").click()
    time.sleep(2)
    element = wait.visibilityOfElementLocated("xpath","//span[text()='写 信']")
    element.click()
    print(u"写信")

    receiver = getElement(driver,"xpath","//div[contains(@id,'_mail_emailinput')]//input")

    receiver.send_keys("757693255@qq.com")
    subject = getElement(driver,"xpath","//div[@aria-label='邮件主题输入框，请输入邮件主题']/input")
    subject.send_keys(u"新邮件")
    attachment=getElement(driver,"xpath","//div[@title='点击添加附件']/input[@size='1' and @type='file']")
    attachment.send_keys("d:\\a.txt")

    time.sleep(5)

    wait.frameToBeAvailableAndSwitchToIt("xpath","//iframe[@tabindex=1]")
    body= getElement(driver,"xpath","/html/body")
    body.send_keys(u"发给自己的一封信")

    driver.switch_to_default_content()

    getElement(driver,"xpath","//header//span[text()='发送']").click()
    print(u"开始发送邮件")
    time.sleep(3)
    assert u"发送成功"in driver.page_source
    driver.quit()

if __name__ =="__main__":
    TestSendMailWithAttachment()