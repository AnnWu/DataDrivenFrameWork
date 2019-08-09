# -*- coding: utf-8 -*-
#用于编写具体的测试逻辑代码
from . import *
from . import CreateContacts
from WriteTestResult import writeTestResult
from util.Log import *

def TestSendMailAndCreateContacts():

    try:
        caseSheet = excelObj.getSheetByName(u"测试用例")

        isExecuteColumn= excelObj.getColumn(caseSheet,testCase_isExcute)

        successfulCase = 0
        requiredCase = 0
        for idx,i in enumerate(isExecuteColumn[1:]):
            #用例sheet中的第一行为标题行，无须执行

            caseName = excelObj.getCellofValue(caseSheet,rowNo=idx+2,colNo=testCase_testCaseName)
            if i.value.lower()=='y':
                requiredCase +=1
                useFrameWorkName = excelObj.getCellofValue(caseSheet,rowNo=idx+2,colNo=testCase_frameWorkName)
                stepSheetName = excelObj.getCellofValue(caseSheet,rowNo=idx+2,colNo=testCase_testStepSheetName)
                print("----%s"%stepSheetName)
                #caseRow = excelObj.getRow(caseSheet,idx+2)
                #caseStepSheetName = caseRow[testCase_testStepSheetName-1].value
                if useFrameWorkName == u"数据":
                    print (u"********调用数据驱动*******")
                    dataSheetName = excelObj.getCellofValue(caseSheet,rowNo=idx+2,colNo=testCase_dataSourceSheetName)
                    dataSheetObj = excelObj.getSheetByName(dataSheetName)
                    stepSheetObj = excelObj.getSheetByName(stepSheetName)
                    result = CreateContacts.dataDriverFun(dataSheetObj,stepSheetObj)
                    if result:
                        print(u"用例%s执行成功"%caseName)
                        successfulCase+=1
                        writeTestResult(caseSheet,rowNo=idx+2,colsNo="testCase",testResult="pass")
                    else :
                        print(u"用例%s执行失败" % caseName)
                        writeTestResult(caseSheet,rowNo=idx+2,colsNo="testCase",
                                        testResult="faild")
                elif useFrameWorkName == u"关键字":
                    print (u"********调用关键字驱动*******")
                    stepSheet = excelObj.getSheetByName(stepSheetName)
                    stepNums = excelObj.getRowsNumber(stepSheet)
                    #记录测试用例i的步骤成功数
                    successfulSteps = 0
                    print(u"测试用例共%s步"%stepNums)
                    print (u"开始执行用例'%s'"%caseName)
                    logging.info(u"开始执行用例'%s'"%caseName)

                    for step in xrange(2,stepNums+1):
                        stepRow = excelObj.getRow(stepSheet,step)

                        #获取步骤的关键字作为调用的函数名
                        keyWord = stepRow[testStep_keyWords-1].value
                        #print keyWord
                        locationType = stepRow[testStep_loactionType-1].value
                        locatorExpression = stepRow[testStep_locatorExpression - 1].value

                        operteValue =  stepRow[testStep_operateValue-1].value

                        if isinstance(operteValue,long):
                            operteValue=str(operteValue)

                        expressionStr = "" #构造需要执行的python 表达式
                        if keyWord and operteValue and\
                                        locationType is None and locatorExpression is None:
                            expressionStr = keyWord.strip()+"(u'"+operteValue+"')"

                        elif keyWord and operteValue is None and \
                                        locationType is None and locatorExpression is None:
                            expressionStr = keyWord.strip() + "()"

                        elif keyWord and operteValue and \
                             locationType and locatorExpression is None:
                            expressionStr = keyWord.strip() + \
                                            "('"+locationType.strip()+"',u'"+operteValue+"')"
                        elif keyWord and operteValue and \
                             locationType and locatorExpression:
                            expressionStr = keyWord.strip() + \
                                        "('" + locationType.strip() + "','"+\
                                        locatorExpression.replace("'",'"').strip()+\
                                        "',u'" + operteValue + "')"

                        elif keyWord and locationType \
                                and locatorExpression and operteValue is None:
                            expressionStr = keyWord.strip() + \
                                        "('" + locationType.strip() + "', '" + \
                                        locatorExpression.replace("'", '"').strip()+ "')"

                        print expressionStr

                        try:
                            eval(expressionStr)
                        except Exception ,e:
                            #截取异常屏幕截图
                            capturePic=capture_screen()
                            #获取详细的异常堆栈信息
                            errorInfo = traceback.format_exc()

                            writeTestResult(stepSheet,step,"caseStep","faild",errorInfo,capturePic)
                            #这里失败的步骤是否加1？
                            print (u"步骤'%s'执行失败！"%stepRow[testStep_testStepDescribe-1].value)
                            logging.info(u"步骤'%s'执行失败,错误信息：%s"%(stepRow[testStep_testStepDescribe-1].value,errorInfo))
                        else:
                            writeTestResult(stepSheet, step, "caseStep", "pass")
                            successfulSteps+=1
                            print (u"步骤'%s'执行通过！" % stepRow[testStep_testStepDescribe - 1].value)
                            logging.info(u"步骤'%s'执行通过！" % stepRow[testStep_testStepDescribe - 1].value)
                    if successfulSteps == stepNums-1:
                        writeTestResult(caseSheet,idx+2,"testCase","pass")
                        successfulCase += 1
                    else:
                        writeTestResult(caseSheet, idx + 2, "testCase", "faild")
            else:
                writeTestResult(caseSheet, idx + 2, "testCase", "")

        print u"共%d 条用例，%d条需要被执行，本次通过%d条"\
                  %(len(isExecuteColumn)-1,requiredCase,successfulCase)
        logging.info(u"共%d 条用例，%d条需要被执行，本次通过%d条."
                         %(len(isExecuteColumn)-1,requiredCase,successfulCase))
    except Exception,e:
        logging.debug(u"程序本身发生异常\n%s"%traceback.format_exc())
        print traceback.print_exc()

#if __name__=="__main__":
#    TestSendMailAndCreateContacts()