# -*- coding: utf-8 -*-
#用于实现模拟键盘单个或者组合按键
import win32api
import win32con

class KeyBoardKeys(object):
    '''
    模拟键盘按键类
    '''
    def __init__(self):
        pass
    VK_CODE = {
        'enter':0x0D,
        'ctrl':0x11,
        'v':0x56
    }
    @staticmethod
    def keyDown(keyName):
        win32api.keyd_event(KeyBoardKeys.VK_CODE[keyName],0,0,0)


    @staticmethod
    def keyUp(keyName):
        win32api.keyd_event(KeyBoardKeys.VK_CODE[keyName], 0, win32con.KEYEVENT_KEYUP, 0)

    @staticmethod
    def oneKey(key):
        KeyBoardKeys.keyDown(key)
        KeyBoardKeys.keyUp(key)

    @staticmethod
    def twoKeys(key1,key2):
        KeyBoardKeys.keyDown(key1)
        KeyBoardKeys.keyDown(key2)
        KeyBoardKeys.keyUp(key2)
        KeyBoardKeys.keyUp(key1)