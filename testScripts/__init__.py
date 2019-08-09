# -*- coding: utf-8 -*-
# created by wmin3
from action.PageAction import *
from util.ParseExcel import ParseExcel
from config.VarConfig import *
import time
import traceback
import sys
reload(sys)
sys.setdefaultencoding("utf-8")

excelObj = ParseExcel()
excelObj.loadWorkBook(dataFilePath)