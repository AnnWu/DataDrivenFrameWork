# -*- coding: utf-8 -*-
# created by wmin3
from . import *
from WriteTestResult import writeTestResult
from util.Log import *
def dataDriverFun(dataSourceSheetObj,stepSheetObj):
    try:
        dataIsExcuteColumn = excelObj.getColumn(dataSourceSheetObj,dataSource_isExcute)

        #获取数据源表中的 电子邮箱 列对象
        emailColumn = excelObj.getColumn(dataSourceSheetObj,dataSource_email)
        #获取测试步骤表中存在数据区域的行数
        stepRowNums = excelObj.getRowsNumber(stepSheetObj)
        successDatas=0
        requireDatas=0

        for idx,data in enumerate(dataIsExcuteColumn[1:]):
            if data.value=="y":
                print u"开始添加联系人'%s'"%emailColumn[idx+1].value
                logging.info(u"开始添加联系人'%s'"%emailColumn[idx+1].value)
                requireDatas+=1
                successStep=0
                for index in xrange(2,stepRowNums+1):
                    #获取数据驱动测试步骤表中 第index行对象
                    rowObj= excelObj.getRow(stepSheetObj,index)
                    keyWord = rowObj[testStep_keyWords-1].value
                    locationType = rowObj[testStep_loactionType - 1].value
                    locatorExpression = rowObj[testStep_locatorExpression - 1].value
                    operteValue = rowObj[testStep_operateValue - 1].value
                    if isinstance(operteValue,long):
                        operteValue=str(operteValue)
                    if operteValue and operteValue.isalpha():
                        coordinate = operteValue+str(idx+2)
                        operteValue=excelObj.getCellofValue(dataSourceSheetObj,coordinate=coordinate)
                    tmpStr="'%s','%s'"%(locationType.lower(),locatorExpression.replace("'",'"')
                                            )if locationType and locatorExpression else ""

                    if tmpStr:
                        tmpStr+=",u'"+operteValue+"'" if operteValue else ""

                    else:
                        tmpStr+="u'"+operteValue+"'" if operteValue else ""
                    runStr = keyWord+"("+tmpStr+")"
                    try:
                        if operteValue!=u"否":
                            eval(runStr)
                    except Exception,e:
                        print (u"执行步骤'%s'发生异常"%rowObj[testStep_testStepDescribe-1].value)
                        print (traceback.print_exc())
                        logging.info(u"执行步骤'%s'发生异常\n"%rowObj[testStep_testStepDescribe-1].value,traceback.format_exc())
                    else :
                        successStep+=1
                        logging.info(u"执行步骤'%s'成功" % rowObj[testStep_testStepDescribe - 1].value)
                        print (u"执行步骤'%s'成功" % rowObj[testStep_testStepDescribe - 1].value)
                if stepRowNums==successStep+1:
                    successDatas+=1
                    writeTestResult(sheetObj=dataSourceSheetObj,rowNo=idx+2,colsNo="dataSheet",testResult="pass")
                else :

                    writeTestResult(sheetObj=dataSourceSheetObj, rowNo=idx + 2, colsNo="dataSheet", testResult="faild")

            else:
                writeTestResult(sheetObj=dataSourceSheetObj,rowNo=idx+2,colsNo="dataSheet",testResult="")
        if requireDatas==successDatas:
            return 1
        return 0
    except Exception,e:
        raise e