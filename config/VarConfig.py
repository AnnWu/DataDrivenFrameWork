# -*- coding: utf-8 -*-
#用于定义整个框架中所需的一些全局常量值，方便维护
import os
ieDriverFilePath="C:\WebDriver\geckodriver"
chromeDriverFilePath="C:\WebDriver\geckodriver"
firefoxDriverFilePath="C:\WebDriver\geckodriver"

parentDirPath = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
#异常图片存放目录
screenPictureDir = parentDirPath + "\\exceptionpictures\\"

dataFilePath = parentDirPath+ u"\\testData\\126邮箱创建联系人发邮件.xlsx"

#用例表部分列对应的序号
testCase_testCaseName = 1
testCase_frameWorkName = 3
testCase_testStepSheetName = 4
testCase_dataSourceSheetName = 5
testCase_isExcute = 6
testCase_runTime = 7
testCase_testResult = 8

#步骤表中，部分列对应的数字序号
testStep_testStepDescribe = 1
testStep_keyWords = 2
testStep_loactionType = 3
testStep_locatorExpression = 4
testStep_operateValue = 5
testStep_runTime = 6
testStep_testResult = 7
testStep_errorInfo = 8
testStep_errorPic = 9

#数据源表中，是否执行列对应的数字编号
dataSource_isExcute=6
dataSource_email = 2
dataSource_runTime = 7
dataSource_result = 8
