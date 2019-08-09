# -*- coding: utf-8 -*-
# created by wmin3
from . import*
from util.Log import *
def writeTestResult(sheetObj,rowNo,colsNo,testResult,errorInfo=None,picPath=None):
    colorDict={'pass':"green",'faild':"red","":None}

    colsDict ={
        "testCase":[testCase_runTime,testCase_testResult],
        "caseStep":[testStep_runTime,testStep_testResult],
        "dataSheet": [dataSource_runTime, dataSource_result]}

    try:
        #excelObj.writeCellCurrenTime(sheetObj,rowNo=rowNo,colNo=colsDict[colsNo][0])
        #在步骤sheet中写入测试结果
        excelObj.writeCell(sheetObj,content=testResult,rowNo=rowNo,\
                           colNo=colsDict[colsNo][1],style= colorDict[testResult])
        if testResult == "":
            #清空时间单元格内容
            excelObj.writeCell(sheetObj,content="",rowNo=rowNo,\
                           colNo=colsDict[colsNo][0])
        else:
            excelObj.writeCellCurrenTime(sheetObj, rowNo=rowNo, colNo=colsDict[colsNo][0])
        if errorInfo and picPath:
            #写入异常信息，测试步骤执行失败时
            excelObj.writeCell(sheetObj,content=errorInfo,rowNo=rowNo,\
                           colNo=testStep_errorInfo)
            #写异常截图
            excelObj.writeCell(sheetObj, content=picPath, rowNo=rowNo, \
                       colNo=testStep_errorPic)

        else:
            if colsNo=="caseStep":
            #
                excelObj.writeCell(sheetObj,content="",rowNo=rowNo,\
                           colNo=testStep_errorInfo)
            #
                excelObj.writeCell(sheetObj, content="", rowNo=rowNo, \
                       colNo=testStep_errorPic)

    except Exception,e:
        print u"写excel 出错",traceback.print_exc()
        logging.debug(u"写excel出错%s"%traceback.print_exc())
