# -*- coding: utf-8 -*-

import sys, os
import datetime, time
from math import ceil, floor # ceil : 소수점 이하를 올림, floor : 소수점 이하를 버림
import math

import pickle
import uuid
import base64
import subprocess
from subprocess import Popen

import PyQt5
from PyQt5 import QtCore, QtGui, uic
from PyQt5 import QAxContainer
from PyQt5.QtGui import *
from PyQt5.QtCore import *
from PyQt5.QtWidgets import (QApplication, QLabel, QLineEdit, QMainWindow, QDialog, QMessageBox, QProgressBar)
from PyQt5.QtWidgets import *
from PyQt5.QAxContainer import *

import numpy as np
from numpy import NaN, Inf, arange, isscalar, asarray, array

import pandas as pd
import pandas.io.sql as pdsql
from pandas import DataFrame, Series
# from pandas.lib import Timestamp

# Google SpreadSheet Read/Write
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from df2gspread import df2gspread as d2g
from string import ascii_uppercase # 알파벳 리스트

import logging
import logging.handlers

import sqlite3

import telepot

# SQLITE DB Setting *****************************************
DATABASE = 'stockdata.db'
def sqliteconn():
    conn = sqlite3.connect(DATABASE)
    return conn

# DB에서 종목명으로 종목코드 반환
def get_code(종목명):
    query = """
                select 종목코드
                from 종목코드
                where (종목명 = '%s')
            """ % (종목명)
    conn = sqliteconn()
    df = pd.read_sql(query, con=conn)
    conn.close()
    return list(df['종목코드'].values)[0]



# Google Spreadsheet Setting *******************************
scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
json_file_name = './secret/xtrader-276902-f5a8b77e2735.json'

credentials = ServiceAccountCredentials.from_json_keyfile_name(json_file_name, scope)
gc = gspread.authorize(credentials)

# XTrader-Stocklist URL
# spreadsheet_url = 'https://docs.google.com/spreadsheets/d/1pLi849EDnjZnaYhphkLButple5bjl33TKZrCoMrim3k/edit#gid=0' # Test Sheet
spreadsheet_url = 'https://docs.google.com/spreadsheets/d/1XE4sk0vDw4fE88bYMDZuJbnP4AF9CmRYHKY6fCXABw4/edit#gid=0' # Sheeet

# spreadsheet 연결 및 worksheet setting
doc = gc.open_by_url(spreadsheet_url)

# stock_sheet = doc.worksheet('test') # Test Sheet
stock_sheet = doc.worksheet('종목선정') # Sheet

strategy_sheet = doc.worksheet('ST bot')

# spreadsheet_key = '1pLi849EDnjZnaYhphkLButple5bjl33TKZrCoMrim3k' # Test Sheet
spreadsheet_key = '1XE4sk0vDw4fE88bYMDZuJbnP4AF9CmRYHKY6fCXABw4' # Sheet

# 스프레드시트 매수 매도 색상 업데이트를 위한 알파벳리스트(열 이름 얻기위함)
alpha_list = list(ascii_uppercase)

# 구글 스프레드 시트 Import후 DataFrame 반환
def import_googlesheet():
    try:
        row_data = stock_sheet.get_all_values()
        row_data[0].insert(2, '종목코드') # 번호, 종목명, 매수모니터링, 매수전략, 매수가1, 매수가2, 매수가3, 매도전략, 매도가

        for row in row_data[1:]:
            try:
                code = get_code(row[1])  # 종목명으로 종목코드 받아서(get_code 함수) 추가
            except Exception as e:
                code = ''
                Telegram('[XTrader]종목명 입력 오류 : %s' % row[1])
            row.insert(2, code)

        Telegram('[XTrader]구글 시트 확인 완료')

        data = pd.DataFrame(data=row_data[1:], columns=row_data[0])
        # 사전 데이터 정리
        data = data[(data['매수모니터링'] == '1') & (data['종목코드']!= '')]
        data = data[['번호', '종목명', '종목코드', '매수모니터링', '매수전략', '매수가1', '매수가2', '매수가3', '매도전략', '매도가']]
        del data['매수모니터링']

        return data

    except Exception as e:
        print("import_googlesheet Error : %s", e)
        logger.info("import_googlesheet Error : %s", e)

# Telegram Setting *****************************************
with open('secret/telegram_token.txt', mode='r') as tokenfile:
    TELEGRAM_TOKEN = tokenfile.readline().strip()
with open('secret/chatid.txt', mode='r') as chatfile:
        CHAT_ID = int(chatfile.readline().strip())
bot = telepot.Bot(TELEGRAM_TOKEN)
def Telegram(str):
    bot.sendMessage(CHAT_ID,str)

# 매수 후 보유기간 계산 *****************************************
today = datetime.date.today()
def periodcal(base_date): # 2018-06-23
    yy = int(base_date[:4]) # 연도
    mm = int(base_date[5:7]) # 월
    dd = int(base_date[8:10]) # 일
    base_d = datetime.date(yy, mm, dd)

    delta = (today - base_d).days # 날짜 차이 수 계산

    week = floor(delta / 7) # 몇 주 차이인지 계산하기 위해 7을 나누고 소수점 이하 버림
    if base_d.weekday() > 4: # 주말일 경우 금요일이라고 변경함
        first_count = 4
    else:
        first_count = base_d.weekday() # 평일은 그냥 해당 요일 사용

    if today.weekday() > 4: # 주말일 경우 금요일이라고 변경함
        last_count = 4
    else:
        last_count = today.weekday() # 평일은 그냥 해당 요일 사용

    delta = (4 - first_count) + ((week - 1) * 5) + (last_count + 1)
    # 계산 방식 : A + B + C
    # 과거 일의 요일을 계산해서 해당 주에서의 워킹데이 카운트(월 ~ 금으로 최대 4일임) : A = 4 - first_count
    # 일주일에 워킹데이가 5일이므로 계산 당일(today)과 과거 일 사이에 몇 주가 있는지를 계산해서 하나 작은 수의 주가 있다고 계산 : B = (week - 1) * 5
    # 마지막으로 현재의 요일에 해당하는 카운트(월요일0면 1, 화요일1이면 2) : C = last_count + 1

    return delta


로봇거래계좌번호 = None

주문딜레이 = 0.25
초당횟수제한 = 5

## 키움증권 제약사항 - 3.7초에 한번 읽으면 지금까지는 괜찮음
주문지연 = 3700 # 3.7초

로봇스크린번호시작 = 9000
로봇스크린번호종료 = 9999


class PandasModel(QtCore.QAbstractTableModel):
    def __init__(self, data=None, parent=None):
        QtCore.QAbstractTableModel.__init__(self, parent)
        self._data = data
        if data is None:
            self._data = DataFrame()

    def rowCount(self, parent=None):
        # return len(self._data.values)
        return len(self._data.index)

    def columnCount(self, parent=None):
        return self._data.columns.size

    def data(self, index, role=Qt.DisplayRole):
        if index.isValid():
            if role == Qt.DisplayRole:
                # return QtCore.QVariant(str(self._data.values[index.row()][index.column()]))
                return str(self._data.values[index.row()][index.column()])
        # return QtCore.QVariant()
        return None

    def headerData(self, column, orientation, role=Qt.DisplayRole):
        if role != Qt.DisplayRole:
            return None
        if orientation == Qt.Horizontal:
            return self._data.columns[column]
        return int(column + 1)

    def update(self, data):
        self._data = data
        self.reset()

    def reset(self):
        self.beginResetModel()
        # unnecessary call to actually clear data, but recommended by design guidance from Qt docs
        # left blank in preliminary testing
        self.endResetModel()

    def flags(self, index):
        return QtCore.Qt.ItemIsEnabled


## 포트폴리오에 사용되는 주식정보 클래스
class CPortStock(object):
    """
    def __init__(self, 매수일, 종목코드, 종목명, 매수가, 매도가1차=0, 매도가2차=0, 손절가=0, 수량=0, 매수단위수=1, STATUS=''):
        self.매수일 = 매수일
        self.종목코드 = 종목코드
        self.종목명 = 종목명
        self.매수가 = 매수가
        self.매도가1차 = 매도가1차
        self.매도가2차 = 매도가2차
        self.손절가 = 손절가
        self.수량 = 수량
        self.매수단위수 = 매수단위수
        self.STATUS = STATUS

        self.이전매수일 = 매수일
        self.이전매수가 = 0
        self.이전수량 = 0
        self.이전매수단위수 = 0
    """
    def __init__(self, 매수일, 종목코드, 종목명, 매수가, 매수조건, 보유일, 매도가=0,손절가=0, 수량=0):
        self.매수일 = 매수일
        self.종목코드 = 종목코드
        self.종목명 = 종목명
        self.매수가 = 매수가
        self.매수조건 = 매수조건
        self.매도가 = 매도가
        self.손절가 = 손절가
        self.수량 = 수량
        self.보유일 = 보유일

    def 평균단가(self):
        if self.이전매수단위수 > 0:
            return ((self.매수가 * self.수량) + (self.이전매수가 * self.이전수량)) // (self.수량 + self.이전수량)
        else:
            return self.매수가


## CTrade 거래로봇용 베이스클래스 : OpenAPI와 붙어서 주문을 내는 등을 하는 클래스
class CTrade(object):
    def __init__(self, sName, UUID, kiwoom=None, parent=None):
        """
        :param sName: 로봇이름
        :param UUID: 로봇구분용 id
        :param kiwoom: 키움OpenAPI
        :param parent: 나를 부른 부모 - 보통은 메인윈도우
        """
        print("CTrade : __init__")

        self.sName = sName
        self.UUID = UUID

        self.sAccount = None  # 거래용계좌번호
        self.kiwoom = kiwoom
        self.parent = parent

        self.running = False  # 실행상태

        self.portfolio= dict()  # 포트폴리오 관리 {'종목코드':종목정보}
        self.현재가 = dict()  # 각 종목의 현재가

    """
    # 조건 검색식 종목 읽기
    def GetCodes(self, Index, Name):
        print("CTrade : GetCodes")
        # logger.info("조건 검색식 종목 읽기")
        # self.kiwoom.OnReceiveTrCondition[str, str, str, int, int].connect(self.OnReceiveTrCondition)
        # self.kiwoom.OnReceiveConditionVer[int, str].connect(self.OnReceiveConditionVer)
        # self.kiwoom.OnReceiveRealCondition[str, str, str, str].connect(self.OnReceiveRealCondition)

        try:
            self.getConditionLoad()
            self.sendCondition("0156", Name, Index, 0) # 선정된 검색조건식으로 바로 종목 검색

        except Exception as e:
            print("GetCondition_Error")
            print(e)

    def getConditionLoad(self):
        self.kiwoom.dynamicCall("GetConditionLoad()")
        print("CTrade : getConditionLoad")
        # receiveConditionVer() 이벤트 메서드에서 루프 종료
        self.ConditionLoop = QEventLoop()
        self.ConditionLoop.exec_()

    def getConditionNameList(self):
        print("CTrade : getConditionNameList")
        data = self.kiwoom.dynamicCall("GetConditionNameList()")

        conditionList = data.split(';')
        del conditionList[-1]

        conditionDictionary = {}

        for condition in conditionList:
            key, value = condition.split('^')
            conditionDictionary[int(key)] = value

        return conditionDictionary

    def sendCondition(self, screenNo, conditionName, conditionIndex, isRealTime):
        print("CTrade : sendCondition")
        isRequest = self.kiwoom.dynamicCall("SendCondition(QString, QString, int, int",
                                     screenNo, conditionName, conditionIndex, isRealTime)

        # receiveTrCondition() 이벤트 메서드에서 루프 종료
        self.ConditionLoop = QEventLoop()
        self.ConditionLoop.exec_()
    """
    # 계좌 보유 종목 받음
    def InquiryList(self, _repeat=0):
        print("CTrade : InquiryList")
        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, Qstring)', "계좌번호", self.sAccount)
        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, Qstring)', "비밀번호입력매체구분", '00')
        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, Qstring)', "조회구분", '1')
        ret = self.kiwoom.dynamicCall('CommRqData(QString, QString, int, QString)', "계좌평가잔고내역요청", "opw00018", _repeat, '{:04d}'.format(self.sScreenNo))

        self.InquiryLoop = QEventLoop()  # 로봇에서 바로 쓸 수 있도록하기 위해서 계좌 조회해서 종목을 받고나서 루프해제시킴
        self.InquiryLoop.exec_()

    # 포트폴리오의 상태
    def GetStatus(self):
        """
        :return: 포트폴리오의 상태
        """
        print("CTrade : GetStatus")
        result = []
        for p, v in self.portfolio.items():
            result.append('%s(%s)[P%s/V%s/D%s]' % (v.종목명.strip(), v.종목코드, v.매수가, v.수량, v.매수일))

        return [self.__class__.__name__, self.sName, self.UUID, self.sScreenNo, self.running, len(self.portfolio), ','.join(result)]

    def GenScreenNO(self):
        """
        :return: 키움증권에서 요구하는 스크린번호를 생성
        """
        print("CTrade : GenScreenNO")
        self.SmallScreenNumber += 1
        if self.SmallScreenNumber > 9999:
            self.SmallScreenNumber = 0

        return self.sScreenNo * 10000 + self.SmallScreenNumber

    def GetLoginInfo(self, tag):
        """
        :param tag:
        :return: 로그인정보 호출
        """
        print("CTrade : GetLoginInfo")
        return self.kiwoom.dynamicCall('GetLoginInfo("%s")' % tag)

    def KiwoomConnect(self):
        """
        :return: 키움증권OpenAPI의 CallBack에 대응하는 처리함수를 연결
        """
        print("CTrade : KiwoomConnect")
        try:
            self.kiwoom.OnEventConnect[int].connect(self.OnEventConnect)
            self.kiwoom.OnReceiveMsg[str, str, str, str].connect(self.OnReceiveMsg)
            self.kiwoom.OnReceiveTrData[str, str, str, str, str, int, str, str, str].connect(self.OnReceiveTrData)
            self.kiwoom.OnReceiveChejanData[str, int, str].connect(self.OnReceiveChejanData)
            self.kiwoom.OnReceiveRealData[str, str, str].connect(self.OnReceiveRealData)

            # self.kiwoom.OnReceiveTrCondition[str, str, str, int, int].connect(self.OnReceiveTrCondition)
            # self.kiwoom.OnReceiveConditionVer[int, str].connect(self.OnReceiveConditionVer)
            # self.kiwoom.OnReceiveRealCondition[str, str, str, str].connect(self.OnReceiveRealCondition)

        except Exception as e:
            print("CTrade : KiwoomConnect Error")
            print(e)
        # logger.info("%s : connected" % self.sName)

    def KiwoomDisConnect(self):
        """
        :return: Callback 연결해제
        """
        print("CTrade : KiwoomDisConnect")
        try:
            self.kiwoom.OnEventConnect[int].disconnect(self.OnEventConnect)
        except Exception:
            pass
        try:
            self.kiwoom.OnReceiveMsg[str, str, str, str].disconnect(self.OnReceiveMsg)
        except Exception:
            pass
        try:
            self.kiwoom.OnReceiveTrCondition[str, str, str, int, int].disconnect(self.OnReceiveTrCondition)
        except Exception:
            pass
        try:
            self.kiwoom.OnReceiveTrData[str, str, str, str, str, int, str, str, str].disconnect(self.OnReceiveTrData)
        except Exception:
            pass
        try:
            self.kiwoom.OnReceiveChejanData[str, int, str].disconnect(self.OnReceiveChejanData)
        except Exception:
            pass
        try:
            self.kiwoom.OnReceiveConditionVer[int, str].disconnect(self.OnReceiveConditionVer)
        except Exception:
            pass
        try:
            self.kiwoom.OnReceiveRealCondition[str, str, str, str].disconnect(self.OnReceiveRealCondition)
        except Exception:
            pass
        try:
            self.kiwoom.OnReceiveRealData[str, str, str].disconnect(self.OnReceiveRealData)
        except Exception:
            pass
            # logger.info("%s : disconnected" % self.sName)

    def KiwoomAccount(self):
        """
        :return: 계좌정보를 읽어옴
        """
        print("CTrade : KiwoomAccount")
        ACCOUNT_CNT = self.GetLoginInfo('ACCOUNT_CNT')
        ACC_NO = self.GetLoginInfo('ACCNO')
        self.account = ACC_NO.split(';')[0:-1]

        self.kiwoom.dynamicCall('SetInputValue(Qstring, Qstring)', "계좌번호", self.account[0])
        self.kiwoom.dynamicCall('CommRqData(QString, QString, int, QString)', "d+2예수금요청", "opw00001", 0, '{:04d}'.format(self.sScreenNo))
        self.depositLoop = QEventLoop() # self.d2_deposit를 로봇에서 바로 쓸 수 있도록하기 위해서 예수금을 받고나서 루프해제시킴
        self.depositLoop.exec_()
        # logger.debug("보유 계좌수: %s 계좌번호: %s [%s]" % (ACCOUNT_CNT, self.account[0], ACC_NO))

    def KiwoomSendOrder(self, sRQName, sScreenNo, sAccNo, nOrderType, sCode, nQty, nPrice, sHogaGb, sOrgOrderNo):
        """
        OpenAPI 메뉴얼 참조
        :param sRQName:
        :param sScreenNo:
        :param sAccNo:
        :param nOrderType:
        :param sCode:
        :param nQty:
        :param nPrice:
        :param sHogaGb:
        :param sOrgOrderNo:
        :return:
        """
        print("CTrade : KiwoomSendOrder")
        order = self.kiwoom.dynamicCall(
            'SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)',
            [sRQName, sScreenNo, sAccNo, nOrderType, sCode, nQty, nPrice, sHogaGb, sOrgOrderNo])
        return order

        # -거래구분값 확인(2자리)
        #
        # 00 : 지정가
        # 03 : 시장가
        # 05 : 조건부지정가
        # 06 : 최유리지정가
        # 07 : 최우선지정가
        # 10 : 지정가IOC
        # 13 : 시장가IOC
        # 16 : 최유리IOC
        # 20 : 지정가FOK
        # 23 : 시장가FOK
        # 26 : 최유리FOK
        # 61 : 장전 시간외단일가매매
        # 81 : 장후 시간외종가
        # 62 : 시간외단일가매매
        #
        # -매매구분값 (1 자리)
        # 1 : 신규매수
        # 2 : 신규매도
        # 3 : 매수취소
        # 4 : 매도취소
        # 5 : 매수정정
        # 6 : 매도정정

    def KiwoomSetRealReg(self, sScreenNo, sCode, sRealType='0'):
        """
        OpenAPI 메뉴얼 참조
        :param sScreenNo:
        :param sCode:
        :param sRealType:
        :return:
        """
        print("CTrade : KiwoomSetRealReg")
        ret = self.kiwoom.dynamicCall('SetRealReg(QString, QString, QString, QString)', sScreenNo, sCode, '9001;10',
                                      sRealType)
        return ret

    def KiwoomSetRealRemove(self, sScreenNo, sCode):
        """
        OpenAPI 메뉴얼 참조
        :param sScreenNo:
        :param sCode:
        :return:
        """
        print("CTrade : KiwoomSetRealRemove")
        ret = self.kiwoom.dynamicCall('SetRealRemove(QString, QString)', sScreenNo, sCode)
        return ret

    def OnEventConnect(self, nErrCode):
        """
        OpenAPI 메뉴얼 참조
        :param nErrCode:
        :return:
        """
        print("CTrade : OnEventConnect")
        logger.info('OnEventConnect', nErrCode)

    def OnReceiveMsg(self, sScrNo, sRQName, sTRCode, sMsg):
        """
        OpenAPI 메뉴얼 참조
        :param sScrNo:
        :param sRQName:
        :param sTRCode:
        :param sMsg:
        :return:
        """
        print("CTrade : OnReceiveMsg")
        logger.info('OnReceiveMsg [%s] [%s] [%s] [%s]' % (sScrNo, sRQName, sTRCode, sMsg))
        self.InquiryLoop.exit()

    def OnReceiveTrData(self, sScrNo, sRQName, sTRCode, sRecordName, sPreNext, nDataLength, sErrorCode, sMessage, sSPlmMsg):
        """
        OpenAPI 메뉴얼 참조
        :param sScrNo:
        :param sRQName:
        :param sTRCode:
        :param sRecordName:
        :param sPreNext:
        :param nDataLength:
        :param sErrorCode:
        :param sMessage:
        :param sSPlmMsg:
        :return:
        """
        logger.info('OnReceiveTrData [%s] [%s] [%s] [%s] [%s] [%s] [%s] [%s] [%s] ' % (sScrNo, sRQName, sTRCode, sRecordName, sPreNext, nDataLength, sErrorCode, sMessage, sSPlmMsg))
        print("CTrade : OnReceiveTrData")

        if self.sScreenNo != int(sScrNo[:4]):
            return

        # logger.info('OnReceiveTrData [%s] [%s] [%s] [%s] [%s] [%s] [%s] [%s] [%s] ' % (sScrNo, sRQName, sTRCode, sRecordName, sPreNext, nDataLength, sErrorCode, sMessage, sSPlmMsg))

        if 'B_' in sRQName or 'S_' in sRQName:
            주문번호 = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, 0, "주문번호")
            # logger.debug("화면번호: %s sRQName : %s 주문번호: %s" % (sScrNo, sRQName, 주문번호))

            self.주문등록(sRQName, 주문번호)

        if sRQName == "d+2예수금요청":
            data = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)',sTRCode, "", sRQName, 0, "d+2추정예수금")

            # 입력된 문자열에 대해 lstrip 메서드를 통해 문자열 왼쪽에 존재하는 '-' 또는 '0'을 제거. 그리고 format 함수를 통해 천의 자리마다 콤마를 추가한 문자열로 변경
            strip_data = data.lstrip('-0')
            if strip_data == '':
                strip_data = '0'

            format_data = format(int(strip_data), ',d')
            if data.startswith('-'):
                format_data = '-' + format_data

            self.sAsset = format_data
            self.depositLoop.exit() # self.d2_deposit를 로봇에서 바로 쓸 수 있도록하기 위해서 예수금을 받고나서 루프해제시킴

        if sRQName == "계좌평가잔고내역요청":
            print("계좌평가잔고내역요청_수신")

            cnt = self.kiwoom.dynamicCall('GetRepeatCnt(QString, QString)', sTRCode, sRQName)
            self.CList = []
            for i in range(0, cnt):
                S = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, i, '종목번호').strip().lstrip('0')
                # print(S)
                if len(S) > 0 and S[0] == '-':
                    S = '-' + S[1:].lstrip('0')

                S = self.종목코드변환(S) # 종목코드 맨 첫 'A'를 삭제하기 위함
                self.CList.append(S)

                # logger.debug("%s" % row)
            if sPreNext == '2':
                self.remained_data = True
                self.InquiryList(_repeat=2)
            else:
                self.remained_data = False

            print(self.CList)
            self.InquiryLoop.exit()

    def OnReceiveChejanData(self, sGubun, nItemCnt, sFidList):
        """
        OpenAPI 메뉴얼 참조
        :param sGubun:
        :param nItemCnt:
        :param sFidList:
        :return:
        """
        # logger.info('OnReceiveChejanData [%s] [%s] [%s]' % (sGubun, nItemCnt, sFidList))

        # 주문체결시 순서
        # 1 구분:0 GetChejanData(913) = '접수'
        # 2 구분:0 GetChejanData(913) = '체결'
        # 3 구분:1 잔고정보

        """
        # sFid별 주요데이터는 다음과 같습니다.
        # "9201" : "계좌번호"
        # "9203" : "주문번호"
        # "9001" : "종목코드"
        # "913" : "주문상태"
        # "302" : "종목명"
        # "900" : "주문수량"
        # "901" : "주문가격"
        # "902" : "미체결수량"
        # "903" : "체결누계금액"
        # "904" : "원주문번호"
        # "905" : "주문구분"
        # "906" : "매매구분"
        # "907" : "매도수구분"
        # "908" : "주문/체결시간"
        # "909" : "체결번호"
        # "910" : "체결가"
        # "911" : "체결량"
        # "10" : "현재가"
        # "27" : "(최우선)매도호가"
        # "28" : "(최우선)매수호가"
        # "914" : "단위체결가"
        # "915" : "단위체결량"
        # "919" : "거부사유"
        # "920" : "화면번호"
        # "917" : "신용구분"
        # "916" : "대출일"
        # "930" : "보유수량"
        # "931" : "매입단가"
        # "932" : "총매입가"
        # "933" : "주문가능수량"
        # "945" : "당일순매수수량"
        # "946" : "매도/매수구분"
        # "950" : "당일총매도손일"
        # "951" : "예수금"
        # "307" : "기준가"
        # "8019" : "손익율"
        # "957" : "신용금액"
        # "958" : "신용이자"
        # "918" : "만기일"
        # "990" : "당일실현손익(유가)"
        # "991" : "당일실현손익률(유가)"
        # "992" : "당일실현손익(신용)"
        # "993" : "당일실현손익률(신용)"
        # "397" : "파생상품거래단위"
        # "305" : "상한가"
        # "306" : "하한가"
        """
        print("CTrade : OnReceiveChejanData")
        # 접수
        if sGubun == "0":
            화면번호 = self.kiwoom.dynamicCall('GetChejanData(QString)', 920)
            if self.sScreenNo != int(화면번호[:4]):
                return

            param = dict()

            param['sGubun'] = sGubun
            param['계좌번호'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 9201)
            param['주문번호'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 9203)
            param['종목코드'] = self.종목코드변환(self.kiwoom.dynamicCall('GetChejanData(QString)', 9001))

            param['주문업무분류'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 912)

            # 접수 / 체결 확인
            # 주문상태(10:원주문, 11:정정주문, 12:취소주문, 20:주문확인, 21:정정확인, 22:취소확인, 90-92:주문거부)
            param['주문상태'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 913)  # 접수 or 체결 확인

            param['종목명'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 302).strip()
            param['주문수량'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 900)
            param['주문가격'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 901)
            param['미체결수량'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 902)
            param['체결누계금액'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 903)
            param['원주문번호'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 904)
            param['주문구분'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 905)
            param['매매구분'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 906)
            param['매도수구분'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 907)
            param['체결시간'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 908)
            param['체결번호'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 909)
            param['체결가'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 910)
            param['체결량'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 911)

            param['현재가'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 10)
            param['매도호가'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 27)
            param['매수호가'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 28)

            param['단위체결가'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 914).strip()
            param['단위체결량'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 915)
            param['화면번호'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 920)

            param['당일매매수수료'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 938)
            param['당일매매세금'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 939)

            param['체결수량'] = int(param['주문수량']) - int(param['미체결수량'])

            # logger.debug('계좌번호:{계좌번호} 체결시간:{체결시간} 주문번호:{주문번호} 체결번호:{체결번호} 종목코드:{종목코드} 종목명:{종목명} 체결량:{체결량} 체결가:{체결가} 단위체결가:{단위체결가} 주문수량:{주문수량} 체결수량:{체결수량} 미체결수량:{미체결수량}'.format(**param))

            if param["주문상태"] == "접수":
                self.접수처리(param)
            if param["주문상태"] == "체결":
                self.체결처리(param)

        # 잔고통보
        if sGubun == "1":
            # logger.debug('OnReceiveChejanData: 잔고통보 [%s] [%s] [%s]' % (sGubun, nItemCnt, sFidList))
            param = dict()

            param['sGubun'] = sGubun
            param['계좌번호'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 9201)
            param['종목코드'] = self.종목코드변환(self.kiwoom.dynamicCall('GetChejanData(QString)', 9001))

            param['신용구분'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 917)
            param['대출일'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 916)

            param['종목명'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 302).strip()
            param['현재가'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 10)

            param['보유수량'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 930)
            param['매입단가'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 931)
            param['총매입가'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 932)
            param['주문가능수량'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 933)
            param['당일순매수량'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 945)
            param['매도매수구분'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 946)
            param['당일총매도손익'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 950)
            param['예수금'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 951)

            param['매도호가'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 27)
            param['매수호가'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 28)

            param['기준가'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 307)
            param['손익율'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 8019)
            param['신용금액'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 957)
            param['신용이자'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 958)
            param['만기일'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 918)
            param['당일실현손익_유가'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 990)
            param['당일실현손익률_유가'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 991)
            param['당일실현손익_신용'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 992)
            param['당일실현손익률_신용'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 993)
            param['담보대출수량'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 959)

            # logger.debug('계좌번호:{계좌번호} 종목명:{종목명} 보유수량:{보유수량} 매입단가:{매입단가} 당일순매수량:{당일순매수량}'.format(**param))

            self.잔고처리(param)

        # 특이신호
        if sGubun == "3":
            # logger.debug('OnReceiveChejanData: 특이신호 [%s] [%s] [%s]' % (sGubun, nItemCnt, sFidList))
            pass

    def OnReceiveRealData(self, sRealKey, sRealType, sRealData):
        """
        OpenAPI 메뉴얼 참조
        :param sRealKey:
        :param sRealType:
        :param sRealData:
        :return:
        """
        # logger.info('OnReceiveRealData [%s] [%s] [%s]' % (sRealKey, sRealType, sRealData))
        _now = datetime.datetime.now()

        if _now.strftime('%H:%M:%S') < '09:00:00': # 9시 이전 데이터 버림(장 시작 전에 테이터 들어오는 것도 많으므로 버리기 위함)
            return

        if sRealKey not in self.실시간종목리스트: # 리스트에 없는 데이터 버림
            return

        if sRealType == "주식시세" or sRealType == "주식체결":
            param = dict()

            param['종목코드'] = self.종목코드변환(sRealKey)
            param['체결시간'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 20).strip()
            param['현재가'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 10).strip()
            param['전일대비'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 11).strip()
            param['등락률'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 12).strip()
            param['매도호가'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 27).strip()
            param['매수호가'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 28).strip()
            param['누적거래량'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 13).strip()
            param['시가'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 16).strip()
            param['고가'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 17).strip()
            param['저가'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 18).strip()
            param['거래회전율'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 31).strip()
            param['시가총액'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 311).strip()

            self.실시간데이타처리(param)

    """
    def OnReceiveTrCondition(self, sScrNo, strCodeList, strConditionName, nIndex, nNext):
        print("CTrade : OnReceiveTrCondition")
        try:
            if strCodeList == "":
                return

            self.codeList = strCodeList.split(';')
            del self.codeList[-1]

            print(self.codeList)
            self.초기조건(self.codeList)

        except Exception as e:
            print("OnReceiveTrCondition_Error")
            print(e)
        finally:
            # pass
            self.ConditionLoop.exit()
            # print("self.ConditionLoop : ", self.ConditionLoop.isRunning)
    
    def OnReceiveConditionVer(self, lRet, sMsg):
        print("CTrade : OnReceiveConditionVer")
        try:
            self.condition = self.getConditionNameList()

        except Exception as e:
            print("CTrade : OnReceiveConditionVer_Error")

        finally:
            self.ConditionLoop.exit()

    def OnReceiveRealCondition(self, sTrCode, strType, strConditionName, strConditionIndex):
        # OpenAPI 메뉴얼 참조
        # :param sTrCode:
        # :param strType:
        # :param strConditionName:
        # :param strConditionIndex:
        # :return:

        logger.info(
            'OnReceiveRealCondition [%s] [%s] [%s] [%s]' % (sTrCode, strType, strConditionName, strConditionIndex))
        print("CTrade : OnReceiveRealCondition")
    """

    def 종목코드변환(self, code): # TR 통해서 받은 종목 코드에 A가 붙을 경우 삭제
        return code.replace('A', '')

    def 정량매수(self, sRQName, 종목코드, 매수가, 수량):
        # sRQName = '정량매수%s' % self.sScreenNo
        sScreenNo = self.GenScreenNO() # 주문을 낼때 마다 스크린번호를 생성
        sAccNo = self.sAccount
        nOrderType = 1  # (1:신규매수, 2:신규매도 3:매수취소, 4:매도취소, 5:매수정정, 6:매도정정)
        sCode = 종목코드
        nQty = 수량
        nPrice = 매수가
        sHogaGb = self.매수방법  # 00:지정가, 03:시장가, 05:조건부지정가, 06:최유리지정가, 07:최우선지정가, 10:지정가IOC, 13:시장가IOC, 16:최유리IOC, 20:지정가FOK, 23:시장가FOK, 26:최유리FOK, 61:장개시전시간외, 62:시간외단일가매매, 81:시간외종가
        if sHogaGb in ['03', '07', '06']:
            nPrice = 0
        sOrgOrderNo = 0

        ret = self.parent.KiwoomSendOrder(sRQName, sScreenNo, sAccNo, nOrderType, sCode, nQty, nPrice, sHogaGb, sOrgOrderNo)

        return ret

    def 정액매수(self, sRQName, 종목코드, 매수가, 매수금액):
        # sRQName = '정액매수%s' % self.sScreenNo
        sScreenNo = self.GenScreenNO()
        sAccNo = self.sAccount
        nOrderType = 1  # (1:신규매수, 2:신규매도 3:매수취소, 4:매도취소, 5:매수정정, 6:매도정정)
        sCode = 종목코드
        nQty = 매수금액 // 매수가
        nPrice = 매수가
        sHogaGb = self.매수방법  # 00:지정가, 03:시장가, 05:조건부지정가, 06:최유리지정가, 07:최우선지정가, 10:지정가IOC, 13:시장가IOC, 16:최유리IOC, 20:지정가FOK, 23:시장가FOK, 26:최유리FOK, 61:장개시전시간외, 62:시간외단일가매매, 81:시간외종가
        if sHogaGb in ['03', '07', '06']:
            nPrice = 0
        sOrgOrderNo = 0

        # logger.debug('주문 - %s %s %s %s %s %s %s %s %s', sRQName, sScreenNo, sAccNo, nOrderType, sCode, nQty, nPrice, sHogaGb, sOrgOrderNo)
        ret = self.parent.KiwoomSendOrder(sRQName, sScreenNo, sAccNo, nOrderType, sCode, nQty, nPrice, sHogaGb,
                                          sOrgOrderNo)

        return ret

    def 정량매도(self, sRQName, 종목코드, 매도가, 수량):
        # sRQName = '정량매도%s' % self.sScreenNo
        sScreenNo = self.GenScreenNO()
        sAccNo = self.sAccount
        nOrderType = 2  # (1:신규매수, 2:신규매도 3:매수취소, 4:매도취소, 5:매수정정, 6:매도정정)
        sCode = 종목코드
        nQty = 수량
        nPrice = 매도가
        sHogaGb = self.매도방법  # 00:지정가, 03:시장가, 05:조건부지정가, 06:최유리지정가, 07:최우선지정가, 10:지정가IOC, 13:시장가IOC, 16:최유리IOC, 20:지정가FOK, 23:시장가FOK, 26:최유리FOK, 61:장개시전시간외, 62:시간외단일가매매, 81:시간외종가
        if sHogaGb in ['03', '07', '06']:
            nPrice = 0
        sOrgOrderNo = 0

        ret = self.parent.KiwoomSendOrder(sRQName, sScreenNo, sAccNo, nOrderType, sCode, nQty, nPrice, sHogaGb,
                                          sOrgOrderNo)

        return ret

    def 정액매도(self, sRQName, 종목코드, 매도가, 매도금액):
        # sRQName = '정액매도%s' % self.sScreenNo
        sScreenNo = self.GenScreenNO()
        sAccNo = self.sAccount
        nOrderType = 2  # (1:신규매수, 2:신규매도 3:매수취소, 4:매도취소, 5:매수정정, 6:매도정정)
        sCode = 종목코드
        nQty = 매도금액 // 매도가
        nPrice = 매도가
        sHogaGb = self.매도방법  # 00:지정가, 03:시장가, 05:조건부지정가, 06:최유리지정가, 07:최우선지정가, 10:지정가IOC, 13:시장가IOC, 16:최유리IOC, 20:지정가FOK, 23:시장가FOK, 26:최유리FOK, 61:장개시전시간외, 62:시간외단일가매매, 81:시간외종가
        if sHogaGb in ['03', '07', '06']:
            nPrice = 0
        sOrgOrderNo = 0

        ret = self.parent.KiwoomSendOrder(sRQName, sScreenNo, sAccNo, nOrderType, sCode, nQty, nPrice, sHogaGb,
                                          sOrgOrderNo)

        return ret

    def 주문등록(self, sRQName, 주문번호):
        self.주문번호_주문_매핑[주문번호] = sRQName



Ui_계좌정보조회, QtBaseClass_계좌정보조회 = uic.loadUiType("./UI/계좌정보조회.ui")
class 화면_계좌정보(QDialog, Ui_계좌정보조회):
    def __init__(self, sScreenNo, kiwoom=None, parent=None):
        super(화면_계좌정보, self).__init__(parent) # Initialize하는 형식
        self.setAttribute(Qt.WA_DeleteOnClose)
        self.setupUi(self)

        self.sScreenNo = sScreenNo
        self.kiwoom = kiwoom
        self.parent = parent

        self.model = PandasModel()
        self.tableView.setModel(self.model)

        self.columns = ['종목번호', '종목명', '현재가', '보유수량', '매입가', '매입금액', '평가금액', '수익률(%)', '평가손익', '매매가능수량']
        self.보이는컬럼 = ['종목번호', '종목명', '현재가', '보유수량', '매입가', '매입금액', '평가금액', '수익률(%)', '평가손익', '매매가능수량'] # 주당 손익 -> 수익률(%)

        self.result = []

        self.KiwoomAccount()

    def KiwoomConnect(self):
        self.kiwoom.OnReceiveMsg[str, str, str, str].connect(self.OnReceiveMsg)
        self.kiwoom.OnReceiveTrData[str, str, str, str, str, int, str, str, str].connect(self.OnReceiveTrData)

    def KiwoomDisConnect(self):
        self.kiwoom.OnReceiveMsg[str, str, str, str].disconnect(self.OnReceiveMsg)
        self.kiwoom.OnReceiveTrData[str, str, str, str, str, int, str, str, str].disconnect(self.OnReceiveTrData)

    def KiwoomAccount(self):
        ACCOUNT_CNT = self.kiwoom.dynamicCall('GetLoginInfo("ACCOUNT_CNT")')
        ACC_NO = self.kiwoom.dynamicCall('GetLoginInfo("ACCNO")')

        self.account = ACC_NO.split(';')[0:-1] # 계좌번호가 ;가 붙어서 나옴(에로 계좌가 3개면 111;222;333)

        self.comboBox.clear()
        self.comboBox.addItems(self.account)

        logger.debug("보유 계좌수: %s 계좌번호: %s [%s]" % (ACCOUNT_CNT, self.account[0], ACC_NO))

    def OnReceiveMsg(self, sScrNo, sRQName, sTrCode, sMsg):
        logger.debug('OnReceiveMsg [%s] [%s] [%s] [%s]' % (sScrNo, sRQName, sTrCode, sMsg))

    def OnReceiveTrData(self, sScrNo, sRQName, sTRCode, sRecordName, sPreNext, nDataLength, sErrorCode, sMessage, sSPlmMsg):
        # logger.debug('OnReceiveTrData [%s] [%s] [%s] [%s] [%s] [%s] [%s] [%s] [%s] ' % (sScrNo, sRQName, sTRCode, sRecordName, sPreNext, nDataLength, sErrorCode, sMessage, sSPlmMsg))
        if self.sScreenNo != int(sScrNo):
            return

        logger.debug('OnReceiveTrData [%s] [%s] [%s] [%s] [%s] [%s] [%s] [%s] [%s] ' % (
        sScrNo, sRQName, sTRCode, sRecordName, sPreNext, nDataLength, sErrorCode, sMessage, sSPlmMsg))

        if sRQName == "계좌평가잔고내역요청":
            cnt = self.kiwoom.dynamicCall('GetRepeatCnt(QString, QString)', sTRCode, sRQName)

            for i in range(0, cnt):
                row = []
                for j in self.columns:
                    # print(j)
                    S = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, i, j).strip().lstrip('0')
                    # print(S)
                    if len(S) > 0 and S[0] == '-':
                        S = '-' + S[1:].lstrip('0')
                    row.append(S)
                self.result.append(row)
                # logger.debug("%s" % row)
            if sPreNext == '2':
                self.Request(_repeat=2)
            else:
                self.model.update(DataFrame(data=self.result, columns=self.보이는컬럼))
                print(self.result)
                for i in range(len(self.columns)):
                    self.tableView.resizeColumnToContents(i)

    def Request(self, _repeat=0):
        계좌번호 = self.comboBox.currentText().strip()
        logger.debug("계좌번호 %s" % 계좌번호)
        # KOA StudioSA에서 opw00018 확인
        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, Qstring)', "계좌번호", 계좌번호) # 8132495511
        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, Qstring)', "비밀번호입력매체구분", '00')
        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, Qstring)', "조회구분", '1')
        ret = self.kiwoom.dynamicCall('CommRqData(QString, QString, int, QString)', "계좌평가잔고내역요청", "opw00018", _repeat,'{:04d}'.format(self.sScreenNo))

    # 조회 버튼(QtDesigner에서 조회버튼 누르고 오른쪽 하단에 시그널/슬롯편집기를 보면 조회버튼 시그널(clicked), 슬롯(Inquiry())로 확인가능함
    def inquiry(self):
        self.result = []
        self.Request(_repeat=0)

    def robot_account(self):
        global 로봇거래계좌번호

        로봇거래계좌번호 = self.comboBox.currentText().strip()

        # sqlite3 사용
        try:
            with sqlite3.connect(DATABASE) as conn:
                cursor = conn.cursor()

                robot_account = pickle.dumps(로봇거래계좌번호, protocol=pickle.HIGHEST_PROTOCOL, fix_imports=True)
                _robot_account = base64.encodebytes(robot_account)

                cursor.execute("REPLACE into Setting(keyword, value) values (?, ?)",
                               ['robotaccount', _robot_account])
                conn.commit()
                print("로봇 계좌 등록 완료")
        except Exception as e:
            print('robot_account', e)

Ui_업종정보, QtBaseClass_업종정보 = uic.loadUiType("./UI/업종정보조회.ui")
class 화면_업종정보(QDialog, Ui_업종정보):
    def __init__(self, sScreenNo, kiwoom=None, parent=None):
        super(화면_업종정보, self).__init__(parent)
        self.setAttribute(Qt.WA_DeleteOnClose)
        self.setupUi(self)

        self.setWindowTitle('업종정보 조회')

        self.sScreenNo = sScreenNo
        self.kiwoom = kiwoom
        self.parent = parent

        self.model = PandasModel()
        self.tableView.setModel(self.model)

        self.columns = ['종목코드', '종목명', '현재가', '대비기호', '전일대비', '등락률', '거래량', '비중', '거래대금', '상한', '상승', '보합', '하락', '하한',
                        '상장종목수']

        self.result = []

        d = today
        self.lineEdit_date.setText(str(d))

    def KiwoomConnect(self):
        self.kiwoom.OnReceiveMsg[str, str, str, str].connect(self.OnReceiveMsg)
        self.kiwoom.OnReceiveTrData[str, str, str, str, str, int, str, str, str].connect(self.OnReceiveTrData)

    def KiwoomDisConnect(self):
        self.kiwoom.OnReceiveMsg[str, str, str, str].disconnect(self.OnReceiveMsg)
        self.kiwoom.OnReceiveTrData[str, str, str, str, str, int, str, str, str].disconnect(self.OnReceiveTrData)

    def OnReceiveMsg(self, sScrNo, sRQName, sTrCode, sMsg):
        logger.debug('OnReceiveMsg [%s] [%s] [%s] [%s]' % (sScrNo, sRQName, sTrCode, sMsg))

    def OnReceiveTrData(self, sScrNo, sRQName, sTRCode, sRecordName, sPreNext, nDataLength, sErrorCode, sMessage,
                        sSPlmMsg):
        # logger.debug('OnReceiveTrData [%s] [%s] [%s] [%s] [%s] [%s] [%s] [%s] [%s] ' % (sScrNo, sRQName, sTRCode, sRecordName, sPreNext, nDataLength, sErrorCode, sMessage, sSPlmMsg))
        if self.sScreenNo != int(sScrNo):
            return

        if sRQName == "업종정보조회":
            종목코드 = ''
            cnt = self.kiwoom.dynamicCall('GetRepeatCnt(QString, QString)', sTRCode, sRQName)
            for i in range(0, cnt):
                row = []
                for j in self.columns:
                    S = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "",
                                                sRQName, i, j).strip().lstrip('0')
                    if len(S) > 0 and S[0] == '-':
                        S = '-' + S[1:].lstrip('0')
                    row.append(S)
                self.result.append(row)
            if sPreNext == '2':
                QTimer.singleShot(주문지연, lambda: self.Request(_repeat=2))
            else:
                df = DataFrame(data=self.result, columns=self.columns)
                df['업종코드'] = self.업종코드
                df.to_csv("업종정보.csv")
                self.model.update(df[['업종코드'] + self.columns])
                for i in range(len(self.columns)):
                    self.tableView.resizeColumnToContents(i)

    def Request(self, _repeat=0):
        self.업종코드 = self.lineEdit_code.text().strip()
        기준일자 = self.lineEdit_date.text().strip().replace('-', '')

        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, Qstring)', "업종코드", self.업종코드)
        ret = self.kiwoom.dynamicCall('CommRqData(QString, QString, int, QString)', "업종정보조회", "OPT20003", _repeat,
                                      '{:04d}'.format(self.sScreenNo))

    def inquiry(self):
        self.result = []
        self.Request(_repeat=0)


Ui_업종별주가조회, QtBaseClass_업종별주가조회 = uic.loadUiType("./UI/업종별주가조회.ui")
class 화면_업종별주가(QDialog, Ui_업종별주가조회):
    def __init__(self, sScreenNo, kiwoom=None, parent=None):
        super(화면_업종별주가, self).__init__(parent)
        self.setAttribute(Qt.WA_DeleteOnClose)
        self.setupUi(self)

        self.setWindowTitle('업종별 주가 조회')

        self.sScreenNo = sScreenNo
        self.kiwoom = kiwoom
        self.parent = parent

        self.model = PandasModel()
        self.tableView.setModel(self.model)

        self.columns = ['현재가', '거래량', '일자', '시가', '고가', '저가', '거래대금', '대업종구분', '소업종구분', '종목정보', '수정주가이벤트', '전일종가']

        self.result = []

        d = today
        self.lineEdit_date.setText(str(d))

    def KiwoomConnect(self):
        self.kiwoom.OnReceiveMsg[str, str, str, str].connect(self.OnReceiveMsg)
        self.kiwoom.OnReceiveTrData[str, str, str, str, str, int, str, str, str].connect(self.OnReceiveTrData)

    def KiwoomDisConnect(self):
        self.kiwoom.OnReceiveMsg[str, str, str, str].disconnect(self.OnReceiveMsg)
        self.kiwoom.OnReceiveTrData[str, str, str, str, str, int, str, str, str].disconnect(self.OnReceiveTrData)

    def OnReceiveMsg(self, sScrNo, sRQName, sTrCode, sMsg):
        logger.debug('OnReceiveMsg [%s] [%s] [%s] [%s]' % (sScrNo, sRQName, sTrCode, sMsg))

    def OnReceiveTrData(self, sScrNo, sRQName, sTRCode, sRecordName, sPreNext, nDataLength, sErrorCode, sMessage,
                        sSPlmMsg):
        # logger.debug('OnReceiveTrData [%s] [%s] [%s] [%s] [%s] [%s] [%s] [%s] [%s] ' % (sScrNo, sRQName, sTRCode, sRecordName, sPreNext, nDataLength, sErrorCode, sMessage, sSPlmMsg))
        if self.sScreenNo != int(sScrNo):
            return

        if sRQName == "업종일봉조회":
            종목코드 = ''
            cnt = self.kiwoom.dynamicCall('GetRepeatCnt(QString, QString)', sTRCode, sRQName)
            for i in range(0, cnt):
                row = []
                for j in self.columns:
                    S = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "",
                                                sRQName, i, j).strip().lstrip('0')
                    if len(S) > 0 and S[0] == '-':
                        S = '-' + S[1:].lstrip('0')
                    row.append(S)
                self.result.append(row)
            if sPreNext == '2':
                QTimer.singleShot(주문지연, lambda: self.Request(_repeat=2))
            else:
                df = DataFrame(data=self.result, columns=self.columns)
                df['업종코드'] = self.업종코드
                self.model.update(df[['업종코드'] + self.columns])
                for i in range(len(self.columns)):
                    self.tableView.resizeColumnToContents(i)

    def Request(self, _repeat=0):
        self.업종코드 = self.lineEdit_code.text().strip()
        기준일자 = self.lineEdit_date.text().strip().replace('-', '')

        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, Qstring)', "업종코드", self.업종코드)
        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, Qstring)', "기준일자", 기준일자)
        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, Qstring)', "수정주가구분", '1')
        ret = self.kiwoom.dynamicCall('CommRqData(QString, QString, int, QString)', "업종일봉조회", "OPT20006", _repeat,
                                      '{:04d}'.format(self.sScreenNo))

    def inquiry(self):
        self.result = []
        self.Request(_repeat=0)


Ui_일자별주가조회, QtBaseClass_일자별주가조회 = uic.loadUiType("./UI/일자별주가조회.ui")
class 화면_일별주가(QDialog, Ui_일자별주가조회):
    def __init__(self, sScreenNo, kiwoom=None, parent=None):
        super(화면_일별주가, self).__init__(parent)
        self.setAttribute(Qt.WA_DeleteOnClose)
        self.setupUi(self)

        self.setWindowTitle('일자별 주가 조회')

        self.sScreenNo = sScreenNo
        self.kiwoom = kiwoom
        self.parent = parent

        self.model = PandasModel()
        self.tableView.setModel(self.model)

        self.columns = ['일자', '현재가', '거래량', '시가', '고가', '저가', '거래대금']

        self.result = []

        d = today
        self.lineEdit_date.setText(str(d))

    def KiwoomConnect(self):
        self.kiwoom.OnReceiveMsg[str, str, str, str].connect(self.OnReceiveMsg)
        self.kiwoom.OnReceiveTrData[str, str, str, str, str, int, str, str, str].connect(self.OnReceiveTrData)

    def KiwoomDisConnect(self):
        self.kiwoom.OnReceiveMsg[str, str, str, str].disconnect(self.OnReceiveMsg)
        self.kiwoom.OnReceiveTrData[str, str, str, str, str, int, str, str, str].disconnect(self.OnReceiveTrData)

    def OnReceiveMsg(self, sScrNo, sRQName, sTrCode, sMsg):
        logger.debug('OnReceiveMsg [%s] [%s] [%s] [%s]' % (sScrNo, sRQName, sTrCode, sMsg))

    def OnReceiveTrData(self, sScrNo, sRQName, sTRCode, sRecordName, sPreNext, nDataLength, sErrorCode, sMessage, sSPlmMsg):
        # logger.debug('OnReceiveTrData [%s] [%s] [%s] [%s] [%s] [%s] [%s] [%s] [%s] ' % (sScrNo, sRQName, sTRCode, sRecordName, sPreNext, nDataLength, sErrorCode, sMessage, sSPlmMsg))
        if self.sScreenNo != int(sScrNo):
            return

        if sRQName == "주식일봉차트조회":
            종목코드 = ''
            cnt = self.kiwoom.dynamicCall('GetRepeatCnt(QString, QString)', sTRCode, sRQName)
            for i in range(0, cnt):
                row = []
                for j in self.columns:
                    S = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "",
                                                sRQName, i, j).strip().lstrip('0')
                    if len(S) > 0 and S[0] == '-':
                        S = '-' + S[1:].lstrip('0')
                    row.append(S)
                self.result.append(row)
            if sPreNext == '2':
                QTimer.singleShot(주문지연, lambda: self.Request(_repeat=2))
            else:
                df = DataFrame(data=self.result, columns=self.columns)
                df['종목코드'] = self.종목코드
                self.model.update(df[['종목코드'] + self.columns])
                for i in range(len(self.columns)):
                    self.tableView.resizeColumnToContents(i)

    def Request(self, _repeat=0):
        self.종목코드 = self.lineEdit_code.text().strip()
        기준일자 = self.lineEdit_date.text().strip().replace('-', '')

        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, Qstring)', "종목코드", self.종목코드)
        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, Qstring)', "기준일자", 기준일자)
        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, Qstring)', "수정주가구분", '1')
        ret = self.kiwoom.dynamicCall('CommRqData(QString, QString, int, QString)', "주식일봉차트조회", "OPT10081", _repeat,
                                      '{:04d}'.format(self.sScreenNo))

    def inquiry(self):
        self.result = []
        self.Request(_repeat=0)

class 화면_종목별투자자(QDialog, Ui_일자별주가조회):
    def __init__(self, sScreenNo, kiwoom=None, parent=None):
        super(화면_종목별투자자, self).__init__(parent)
        self.setAttribute(Qt.WA_DeleteOnClose)
        self.setupUi(self)

        self.setWindowTitle('종목별 투자자 조회')

        self.sScreenNo = sScreenNo
        self.kiwoom = kiwoom
        self.parent = parent

        self.model = PandasModel()
        self.tableView.setModel(self.model)

        self.columns = ['일자', '현재가', '전일대비', '누적거래대금', '개인투자자', '외국인투자자', '기관계', '금융투자', '보험', '투신', '기타금융', '은행',
                        '연기금등', '국가', '내외국인', '사모펀드', '기타법인']

        self.result = []

        d = today
        self.lineEdit_date.setText(str(d))

    def KiwoomConnect(self):
        self.kiwoom.OnReceiveMsg[str, str, str, str].connect(self.OnReceiveMsg)
        self.kiwoom.OnReceiveTrData[str, str, str, str, str, int, str, str, str].connect(self.OnReceiveTrData)

    def KiwoomDisConnect(self):
        self.kiwoom.OnReceiveMsg[str, str, str, str].disconnect(self.OnReceiveMsg)
        self.kiwoom.OnReceiveTrData[str, str, str, str, str, int, str, str, str].disconnect(self.OnReceiveTrData)

    def OnReceiveMsg(self, sScrNo, sRQName, sTrCode, sMsg):
        logger.debug('OnReceiveMsg [%s] [%s] [%s] [%s]' % (sScrNo, sRQName, sTrCode, sMsg))

    def OnReceiveTrData(self, sScrNo, sRQName, sTRCode, sRecordName, sPreNext, nDataLength, sErrorCode, sMessage, sSPlmMsg):
        # logger.debug('OnReceiveTrData [%s] [%s] [%s] [%s] [%s] [%s] [%s] [%s] [%s] ' % (sScrNo, sRQName, sTRCode, sRecordName, sPreNext, nDataLength, sErrorCode, sMessage, sSPlmMsg))
        if self.sScreenNo != int(sScrNo):
            return

        if sRQName == "종목별투자자조회":
            cnt = self.kiwoom.dynamicCall('GetRepeatCnt(QString, QString)', sTRCode, sRQName)
            for i in range(0, cnt):
                row = []
                for j in self.columns:
                    S = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "",
                                                sRQName, i, j).strip().lstrip('0')
                    row.append(S)
                self.result.append(row)
            if sPreNext == '2':
                QTimer.singleShot(주문지연, lambda: self.Request(_repeat=2))
            else:
                df = DataFrame(data=self.result, columns=self.columns)
                df['종목코드'] = self.lineEdit_code.text().strip()
                dfnew = df[['종목코드'] + self.columns]
                self.model.update(dfnew)
                for i in range(len(self.columns)):
                    self.tableView.resizeColumnToContents(i)

    def Request(self, _repeat=0):
        종목코드 = self.lineEdit_code.text().strip()
        기준일자 = self.lineEdit_date.text().strip().replace('-', '')

        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, Qstring)', "일자", 기준일자)
        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, Qstring)', "종목코드", 종목코드)
        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, int)', "금액수량구분", 2)  # 1:금액, 2:수량
        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, int)', "매매구분", 0)  # 0:순매수, 1:매수, 2:매도
        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, int)', "단위구분", 1)  # 1000:천주, 1:단주
        ret = self.kiwoom.dynamicCall('CommRqData(QString, QString, int, QString)', "종목별투자자조회", "OPT10060", _repeat,
                                      '{:04d}'.format(self.sScreenNo))

    def inquiry(self):
        self.result = []
        self.Request(_repeat=0)


Ui_분별주가조회, QtBaseClass_분별주가조회 = uic.loadUiType("./UI/분별주가조회.ui")
class 화면_분별주가(QDialog, Ui_분별주가조회):
    def __init__(self, sScreenNo, kiwoom=None, parent=None):
        super(화면_분별주가, self).__init__(parent)
        self.setAttribute(Qt.WA_DeleteOnClose)
        self.setupUi(self)

        self.setWindowTitle('분별 주가 조회')

        self.sScreenNo = sScreenNo
        self.kiwoom = kiwoom
        self.parent = parent

        self.model = PandasModel()
        self.tableView.setModel(self.model)

        self.columns = ['체결시간', '현재가', '시가', '고가', '저가', '거래량']

        self.result = []

    def KiwoomConnect(self):
        self.kiwoom.OnReceiveMsg[str, str, str, str].connect(self.OnReceiveMsg)
        self.kiwoom.OnReceiveTrData[str, str, str, str, str, int, str, str, str].connect(self.OnReceiveTrData)

    def KiwoomDisConnect(self):
        self.kiwoom.OnReceiveMsg[str, str, str, str].disconnect(self.OnReceiveMsg)
        self.kiwoom.OnReceiveTrData[str, str, str, str, str, int, str, str, str].disconnect(self.OnReceiveTrData)

    def OnReceiveMsg(self, sScrNo, sRQName, sTrCode, sMsg):
        logger.debug('OnReceiveMsg [%s] [%s] [%s] [%s]' % (sScrNo, sRQName, sTrCode, sMsg))

    def OnReceiveTrData(self, sScrNo, sRQName, sTRCode, sRecordName, sPreNext, nDataLength, sErrorCode, sMessage,
                        sSPlmMsg):
        # logger.debug('OnReceiveTrData [%s] [%s] [%s] [%s] [%s] [%s] [%s] [%s] [%s] ' % (sScrNo, sRQName, sTRCode, sRecordName, sPreNext, nDataLength, sErrorCode, sMessage, sSPlmMsg))
        if self.sScreenNo != int(sScrNo):
            return

        if sRQName == "주식분봉차트조회":
            종목코드 = ''
            cnt = self.kiwoom.dynamicCall('GetRepeatCnt(QString, QString)', sTRCode, sRQName)
            for i in range(0, cnt):
                row = []
                for j in self.columns:
                    S = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "",
                                                sRQName, i, j).strip().lstrip('0')
                    if len(S) > 0 and (S[0] == '-' or S[0] == '+'):
                        S = S[1:].lstrip('0')
                    row.append(S)
                self.result.append(row)
            if sPreNext == '2':
                QTimer.singleShot(주문지연, lambda: self.Request(_repeat=2))
            else:
                df = DataFrame(data=self.result, columns=self.columns)
                df['종목코드'] = self.종목코드
                self.model.update(df[['종목코드'] + self.columns])
                for i in range(len(self.columns)):
                    self.tableView.resizeColumnToContents(i)

    def Request(self, _repeat=0):
        self.종목코드 = self.lineEdit_code.text().strip()
        틱범위 = self.comboBox_min.currentText()[0:2].strip()
        if 틱범위[0] == '0':
            틱범위 = 틱범위[1:]
        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, Qstring)', "종목코드", self.종목코드)
        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, Qstring)', "틱범위", 틱범위)
        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, Qstring)', "수정주가구분", '1')
        ret = self.kiwoom.dynamicCall('CommRqData(QString, QString, int, QString)', "주식분봉차트조회", "OPT10080", _repeat,
                                      '{:04d}'.format(self.sScreenNo))

    def inquiry(self):
        self.result = []
        self.Request(_repeat=0)

class RealDataTableModel(QAbstractTableModel):
    def __init__(self, parent=None):
        QtCore.QAbstractTableModel.__init__(self, parent)
        self.realdata = {}
        self.headers = ['종목코드', '현재가', '전일대비', '등락률', '매도호가', '매수호가', '누적거래량', '시가', '고가', '저가', '거래회전율', '시가총액']

    def rowCount(self, index=QModelIndex()):
        return len(self.realdata)

    def columnCount(self, index=QModelIndex()):
        return len(self.headers)

    def data(self, index, role=Qt.DisplayRole):
        if (not index.isValid() or not (0 <= index.row() < len(self.realdata))):
            return None

        if role == Qt.DisplayRole:
            rows = []
            for k in self.realdata.keys():
                rows.append(k)
            one_row = rows[index.row()]
            selected_row = self.realdata[one_row]

            return selected_row[index.column()]

        return None

    def headerData(self, column, orientation, role=Qt.DisplayRole):
        if role != Qt.DisplayRole:
            return None
        if orientation == Qt.Horizontal:
            return self.headers[column]
        return int(column + 1)

    def flags(self, index):
        return QtCore.Qt.ItemIsEnabled

    def reset(self):
        self.beginResetModel()
        self.endResetModel()


Ui_실시간정보, QtBaseClass_실시간정보 = uic.loadUiType("./UI/실시간정보.ui")
class 화면_실시간정보(QDialog, Ui_실시간정보):
    def __init__(self, sScreenNo, kiwoom=None, parent=None):
        super(화면_실시간정보, self).__init__(parent)
        self.setAttribute(Qt.WA_DeleteOnClose)
        self.setupUi(self)

        self.sScreenNo = sScreenNo
        self.kiwoom = kiwoom
        self.parent = parent

        self.model = RealDataTableModel()
        self.tableView.setModel(self.model)

    def KiwoomConnect(self):
        self.kiwoom.OnEventConnect[int].connect(self.OnEventConnect)
        self.kiwoom.OnReceiveMsg[str, str, str, str].connect(self.OnReceiveMsg)
        self.kiwoom.OnReceiveTrCondition[str, str, str, int, int].connect(self.OnReceiveTrCondition)
        self.kiwoom.OnReceiveTrData[str, str, str, str, str, int, str, str, str].connect(self.OnReceiveTrData)
        self.kiwoom.OnReceiveChejanData[str, int, str].connect(self.OnReceiveChejanData)
        self.kiwoom.OnReceiveConditionVer[int, str].connect(self.OnReceiveConditionVer)
        self.kiwoom.OnReceiveRealCondition[str, str, str, str].connect(self.OnReceiveRealCondition)
        self.kiwoom.OnReceiveRealData[str, str, str].connect(self.OnReceiveRealData)

    def KiwoomDisConnect(self):
        self.kiwoom.OnEventConnect[int].disconnect(self.OnEventConnect)
        self.kiwoom.OnReceiveMsg[str, str, str, str].disconnect(self.OnReceiveMsg)
        self.kiwoom.OnReceiveTrCondition[str, str, str, int, int].disconnect(self.OnReceiveTrCondition)
        self.kiwoom.OnReceiveTrData[str, str, str, str, str, int, str, str, str].disconnect(self.OnReceiveTrData)
        self.kiwoom.OnReceiveChejanData[str, int, str].disconnect(self.OnReceiveChejanData)
        self.kiwoom.OnReceiveConditionVer[int, str].disconnect(self.OnReceiveConditionVer)
        self.kiwoom.OnReceiveRealCondition[str, str, str, str].disconnect(self.OnReceiveRealCondition)
        self.kiwoom.OnReceiveRealData[str, str, str].disconnect(self.OnReceiveRealData)

    def KiwoomAccount(self):
        ACCOUNT_CNT = self.kiwoom.dynamicCall('GetLoginInfo("ACCOUNT_CNT")')
        ACC_NO = self.kiwoom.dynamicCall('GetLoginInfo("ACCNO")')

        self.account = ACC_NO.split(';')[0:-1]
        self.plainTextEdit.insertPlainText("보유 계좌수: %s 계좌번호: %s [%s]" % (ACCOUNT_CNT, self.account[0], ACC_NO))
        self.plainTextEdit.insertPlainText("\r\n")

    def KiwoomSendOrder(self, sRQName, sScreenNo, sAccNo, nOrderType, sCode, nQty, nPrice, sHogaGb, sOrgOrderNo):
        order = self.kiwoom.dynamicCall(
            'SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)',
            [sRQName, sScreenNo, sAccNo, nOrderType, sCode, nQty, nPrice, sHogaGb, sOrgOrderNo])
        return order

    def KiwoomSetRealReg(self, sScreenNo, sCode, sRealType='0'):
        ret = self.kiwoom.dynamicCall('SetRealReg(QString, QString, QString, QString)', sScreenNo, sCode, '9001;10',
                                      sRealType)
        return ret

    def KiwoomSetRealRemove(self, sScreenNo, sCode):
        ret = self.kiwoom.dynamicCall('SetRealRemove(QString, QString)', sScreenNo, sCode)
        return ret

    def OnEventConnect(self, nErrCode):
        self.plainTextEdit.insertPlainText('OnEventConnect %s' % nErrCode)
        self.plainTextEdit.insertPlainText("\r\n")

    def OnReceiveMsg(self, sScrNo, sRQName, sTrCode, sMsg):
        self.plainTextEdit.insertPlainText('OnReceiveMsg [%s] [%s] [%s] [%s]' % (sScrNo, sRQName, sTrCode, sMsg))
        self.plainTextEdit.insertPlainText("\r\n")

    def OnReceiveTrCondition(self, sScrNo, strCodeList, strConditionName, nIndex, nNext):
        self.plainTextEdit.insertPlainText(
            'OnReceiveTrCondition [%s] [%s] [%s] [%s] [%s]' % (sScrNo, strCodeList, strConditionName, nIndex, nNext))
        self.plainTextEdit.insertPlainText("\r\n")
        self.plainTextEdit.insertPlainText("[화면번호] : %s" % sScrNo)
        self.plainTextEdit.insertPlainText("\r\n")
        self.plainTextEdit.insertPlainText("[종목리스트] : %s" % strCodeList)
        self.plainTextEdit.insertPlainText("\r\n")
        self.plainTextEdit.insertPlainText("[조건명] : %s" % strConditionName)
        self.plainTextEdit.insertPlainText("\r\n")
        self.plainTextEdit.insertPlainText("[조건명 인덱스 ] : %s" % nIndex)
        self.plainTextEdit.insertPlainText("\r\n")
        self.plainTextEdit.insertPlainText("[연속조회] : %s" % nNext)
        self.plainTextEdit.insertPlainText("\r\n")

    def OnReceiveTrData(self, sScrNo, sRQName, sTRCode, sRecordName, sPreNext, nDataLength, sErrorCode, sMessage,
                        sSPlmMsg):
        self.plainTextEdit.insertPlainText('OnReceiveTrData [%s] [%s] [%s] [%s] [%s] [%s] [%s] [%s] [%s] ' % (
        sScrNo, sRQName, sTRCode, sRecordName, sPreNext, nDataLength, sErrorCode, sMessage, sSPlmMsg))
        self.plainTextEdit.insertPlainText("\r\n")

    def OnReceiveChejanData(self, sGubun, nItemCnt, sFidList):
        self.plainTextEdit.insertPlainText('OnReceiveChejanData [%s] [%s] [%s]' % (sGubun, nItemCnt, sFidList))
        self.plainTextEdit.insertPlainText("\r\n")

        if sGubun == "0":
            param = dict()

            param['계좌번호'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 9201)
            param['주문번호'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 9203)
            param['종목코드'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 9001)
            param['종목명'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 302)
            param['주문수량'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 900)
            param['주문가격'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 901)
            param['원주문번호'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 904)
            param['체결량'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 911)
            param['미체결수량'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 902)
            param['매도수구분'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 907)
            param['단위체결가'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 914)
            param['화면번호'] = self.kiwoom.dynamicCall('GetChejanData(QString)', 920)

            self.plainTextEdit.insertPlainText(str(param))
            self.plainTextEdit.insertPlainText("\r\n")

        if sGubun == "1":
            self.plainTextEdit.insertPlainText(
                'OnReceiveChejanData: 잔고통보 [%s] [%s] [%s]' % (sGubun, nItemCnt, sFidList))
            self.plainTextEdit.insertPlainText("\r\n")
        if sGubun == "3":
            self.plainTextEdit.insertPlainText(
                'OnReceiveChejanData: 특이신호 [%s] [%s] [%s]' % (sGubun, nItemCnt, sFidList))
            self.plainTextEdit.insertPlainText("\r\n")

    def OnReceiveConditionVer(self, lRet, sMsg):
        self.plainTextEdit.insertPlainText('OnReceiveConditionVer : [이벤트] 조건식 저장 %s %s' % (lRet, sMsg))

    def OnReceiveRealCondition(self, sTrCode, strType, strConditionName, strConditionIndex):
        self.plainTextEdit.insertPlainText(
            'OnReceiveRealCondition [%s] [%s] [%s] [%s]' % (sTrCode, strType, strConditionName, strConditionIndex))
        self.plainTextEdit.insertPlainText("\r\n")
        self.plainTextEdit.insertPlainText("========= 조건조회 실시간 편입/이탈 ==========")
        self.plainTextEdit.insertPlainText("\r\n")
        self.plainTextEdit.insertPlainText("[종목코드] : %s" % sTrCode)
        self.plainTextEdit.insertPlainText("\r\n")
        self.plainTextEdit.insertPlainText("[실시간타입] : %s" % strType)
        self.plainTextEdit.insertPlainText("\r\n")
        self.plainTextEdit.insertPlainText("[조건명] : %s" % strConditionName)
        self.plainTextEdit.insertPlainText("\r\n")
        self.plainTextEdit.insertPlainText("[조건명 인덱스] : %s" % strConditionIndex)
        self.plainTextEdit.insertPlainText("\r\n")

    def OnReceiveRealData(self, sRealKey, sRealType, sRealData):
        self.plainTextEdit.insertPlainText("[%s] [%s] %s\n" % (sRealKey, sRealType, sRealData))

        if sRealType == "주식시세" or sRealType == "주식체결":
            param = dict()

            param['종목코드'] = sRealKey.strip()
            param['현재가'] = abs(int(self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 10).strip()))
            param['전일대비'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 11).strip()
            param['등락률'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 12).strip()
            param['매도호가'] = abs(int(self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 27).strip()))
            param['매수호가'] = abs(int(self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 28).strip()))
            param['누적거래량'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 13).strip()
            param['시가'] = abs(int(self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 16).strip()))
            param['고가'] = abs(int(self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 17).strip()))
            param['저가'] = abs(int(self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 18).strip()))
            param['거래회전율'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 31).strip()
            param['시가총액'] = self.kiwoom.dynamicCall("GetCommRealData(QString, int)", sRealType, 311).strip()

            self.model.realdata[sRealKey] = [param['종목코드'], param['현재가'], param['전일대비'], param['등락률'], param['매도호가'],
                                             param['매수호가'], param['누적거래량'], param['시가'], param['고가'], param['저가'],
                                             param['거래회전율'], param['시가총액']]
            self.model.reset()

            for i in range(len(self.model.realdata[sRealKey])):
                self.tableView.resizeColumnToContents(i)


## TickLogger
Ui_TickLogger, QtBaseClass_TickLogger = uic.loadUiType("./UI/TickLogger.ui")
class 화면_TickLogger(QDialog, Ui_TickLogger):
    def __init__(self, parent):
        super(화면_TickLogger, self).__init__(parent)
        self.setupUi(self)

class CTickLogger(CTrade):
    def __init__(self, sName, UUID, kiwoom=None, parent=None):
        self.sName = sName
        self.UUID = UUID

        self.sAccount = None
        self.kiwoom = kiwoom
        self.parent = parent

        self.running = False

        self.portfolio = dict()
        self.실시간종목리스트 = []

        self.SmallScreenNumber = 9999

        self.buffer = []

        self.d = today

    def Setting(self, sScreenNo, 종목유니버스):
        self.sScreenNo = sScreenNo
        self.종목유니버스 = 종목유니버스

        self.실시간종목리스트 = 종목유니버스

    def 실시간데이타처리(self, param):

        if self.running == True:
            _체결시간 = '%s %s:%s:%s' % (str(self.d), param['체결시간'][0:2], param['체결시간'][2:4], param['체결시간'][4:])
            if len(self.buffer) < 100:
                체결시간 = '%s %s:%s:%s' % (str(self.d), param['체결시간'][0:2], param['체결시간'][2:4], param['체결시간'][4:])
                종목코드 = param['종목코드']
                현재가 = abs(int(float(param['현재가'])))
                전일대비 = int(float(param['전일대비']))
                등락률 = float(param['등락률'])
                매도호가 = abs(int(float(param['매도호가'])))
                매수호가 = abs(int(float(param['매수호가'])))
                누적거래량 = abs(int(float(param['누적거래량'])))
                시가 = abs(int(float(param['시가'])))
                고가 = abs(int(float(param['고가'])))
                저가 = abs(int(float(param['저가'])))
                거래회전율 = abs(float(param['거래회전율']))
                시가총액 = abs(int(float(param['시가총액'])))

                lst = [체결시간, 종목코드, 현재가, 전일대비, 등락률, 매도호가, 매수호가, 누적거래량, 시가, 고가, 저가, 거래회전율, 시가총액]
                self.buffer.append(lst)
                self.parent.statusbar.showMessage(
                    "[%s]%s %s %s %s" % (_체결시간, 종목코드, self.parent.CODE_POOL[종목코드][1], 현재가, 전일대비))
            else:
                df = DataFrame(data=self.buffer,
                               columns=['체결시간', '종목코드', '현재가', '전일대비', '등락률', '매도호가', '매수호가', '누적거래량', '시가', '고가', '저가',
                                        '거래회전율', '시가총액'])
                df.to_csv('TickLogger.csv', mode='a', header=False)
                self.buffer = []
                self.parent.statusbar.showMessage("CTickLogger 기록함")

    def 접수처리(self, param):
        pass

    def 체결처리(self, param):
        pass

    def 잔고처리(self, param):
        pass

    def Run(self, flag=True, sAccount=None):
        self.running = flag

        ret = 0
        if flag == True:
            self.KiwoomConnect()
            ret = self.KiwoomSetRealReg(self.sScreenNo, ';'.join(self.종목유니버스) + ';')

        else:
            ret = self.KiwoomSetRealRemove(self.sScreenNo, 'ALL')
            self.KiwoomDisConnect()

            df = DataFrame(data=self.buffer,
                           columns=['체결시간', '종목코드', '현재가', '전일대비', '등락률', '매도호가', '매수호가', '누적거래량', '시가', '고가', '저가',
                                    '거래회전율', '시가총액'])
            df.to_csv('TickLogger.csv', mode='a', header=False)
            self.buffer = []
            self.parent.statusbar.showMessage("CTickLogger 기록함")


## TickMonitor
class CTickMonitor(CTrade):
    def __init__(self, sName, UUID, kiwoom=None, parent=None):
        self.sName = sName
        self.UUID = UUID

        self.sAccount = None
        self.kiwoom = kiwoom
        self.parent = parent

        self.running = False

        self.portfolio = dict()
        self.실시간종목리스트 = []

        self.SmallScreenNumber = 9999

        self.buffer = []

        self.d = today

        self.모니터링종목 = dict()
        self.누적거래량 = dict()

    def Setting(self, sScreenNo, 종목유니버스):
        self.sScreenNo = sScreenNo
        self.종목유니버스 = 종목유니버스

        self.실시간종목리스트 = 종목유니버스

    def 실시간데이타처리(self, param):

        if self.running == True:

            체결시간 = '%s %s:%s:%s' % (str(self.d), param['체결시간'][0:2], param['체결시간'][2:4], param['체결시간'][4:])
            종목코드 = param['종목코드']
            현재가 = abs(int(float(param['현재가'])))
            # 체결량 = abs(int(float(param['체결량'])))
            전일대비 = int(float(param['전일대비']))
            등락률 = float(param['등락률'])
            매도호가 = abs(int(float(param['매도호가'])))
            매수호가 = abs(int(float(param['매수호가'])))
            누적거래량 = abs(int(float(param['누적거래량'])))
            시가 = abs(int(float(param['시가'])))
            고가 = abs(int(float(param['고가'])))
            저가 = abs(int(float(param['저가'])))
            거래회전율 = abs(float(param['거래회전율']))
            시가총액 = abs(int(float(param['시가총액'])))

            체결량 = 0
            if self.누적거래량.get(종목코드) == None:
                self.누적거래량[종목코드] = 누적거래량
            else:
                체결량 = 누적거래량 - self.누적거래량[종목코드]
                if 체결량 < 2:                                  # 1틱씩 거래가 되는 종목만 포착
                # if 체결량 in [1, 12, 18]:                     $ 틱 조건 예시
                    if self.모니터링종목.get(종목코드) == None:
                        self.모니터링종목[종목코드] = 1
                    else:
                        self.모니터링종목[종목코드] += 1

                    temp = []
                    for k, v in self.모니터링종목.items():
                        if v >= 10:                            # 1틱씩 거래가 되는게 10번 넘으면 temp에 저장하고 화면에 보여줌
                            temp.append(k)
                    if len(temp) > 0:
                        logger.info("%s %s" % (체결시간, temp))
                else:
                    pass

                self.누적거래량[종목코드] = 누적거래량

                self.parent.statusbar.showMessage(
                    "[%s]%s %s %s %s" % (체결시간, 종목코드, self.parent.CODE_POOL[종목코드][1], 현재가, 전일대비))

    def 초기조건(self):
        pass

    def 접수처리(self, param):
        pass

    def 체결처리(self, param):
        pass

    def 잔고처리(self, param):
        pass

    def Run(self, flag=True, sAccount=None):
        self.running = flag

        ret = 0
        if flag == True:
            ret = self.KiwoomSetRealReg(self.sScreenNo, ';'.join(self.종목유니버스) + ';')  # 실시간 데이터 요청
            self.KiwoomConnect()
        else:
            ret = self.KiwoomSetRealRemove(self.sScreenNo, 'ALL')
            self.KiwoomDisConnect()

## TickTradeRSI
Ui_TickTradeRSI, QtBaseClass_TickTradeRSI = uic.loadUiType("./UI/TickTradeRSI.ui")
class 화면_TickTradeRSI(QDialog, Ui_TickTradeRSI):
    def __init__(self, parent):
        super(화면_TickTradeRSI, self).__init__(parent)
        self.setupUi(self)

class CTickTradeRSI(CTrade): # 로봇 추가 시 __init__ : 복사, Setting / 초기조건:전략에 맞게, 데이터처리 / Run:복사
    # 동작 순서
    # 1. Robot Add에서 화면에서 주요 파라미터 받고 호출되면서 __init__ 실행
    # 2. Setting 실행 후 로봇 추가 셋팅 완료
    # 3. Robot_Run이 되면 초기 조건 실행하여 매수/매도 종목을 리스트로 저장하고 Run 실행

    def __init__(self, sName, UUID, kiwoom=None, parent=None):
        self.sName = sName
        self.UUID = UUID

        self.sAccount = None
        self.kiwoom = kiwoom
        self.parent = parent

        self.running = False

        self.주문결과 = dict()
        self.주문번호_주문_매핑 = dict()
        self.주문실행중_Lock = dict()

        self.portfolio = dict()

        self.실시간종목리스트 = []

        self.SmallScreenNumber = 9999

        self.d = today

    def Setting(self, sScreenNo, 포트폴리오수=10, 단위투자금=300 * 10000, 시총상한=4000, 시총하한=500, 매수방법='00', 매도방법='00'):
        self.sScreenNo = sScreenNo
        self.실시간종목리스트 = []
        self.단위투자금 = 단위투자금
        self.매수방법 = 매수방법
        self.매도방법 = 매도방법
        self.포트폴리오수 = 포트폴리오수
        self.시총상한 = 시총상한
        self.시총하한 = 시총하한

    def get_price(self, code, 시작일자=None, 종료일자=None):

        if 시작일자 == None and 종료일자 == None:
            query = """
            SELECT 일자, 종가, 시가, 고가, 저가, 거래량
            FROM 일별주가
            WHERE 종목코드='%s'
            ORDER BY 일자 ASC
            """ % (code)
        if 시작일자 != None and 종료일자 == None:
            query = """
            SELECT 일자, 종가, 시가, 고가, 저가, 거래량
            FROM 일별주가
            WHERE 종목코드='%s' AND A.일자 >= '%s'
            ORDER BY 일자 ASC
            """ % (code, 시작일자)
        if 시작일자 == None and 종료일자 != None:
            query = """
            SELECT 일자, 종가, 시가, 고가, 저가, 거래량
            FROM 일별주가
            WHERE 종목코드='%s' AND 일자 <= '%s'
            ORDER BY 일자 ASC
            """ % (code, 종료일자)
        if 시작일자 != None and 종료일자 != None:
            query = """
            SELECT 일자, 종가, 시가, 고가, 저가, 거래량
            FROM 일별주가
            WHERE 종목코드='%s' AND 일자 BETWEEN '%s' AND '%s'
            ORDER BY 일자 ASC
            """ % (code, 시작일자, 종료일자)

        conn = sqliteconn()
        df = pdsql.read_sql_query(query, con=conn)
        conn.close()

        df.fillna(0, inplace=True)
        df.set_index('일자', inplace=True)

        df['RSI'] = ta.RSI(np.array(df['종가'].astype(float)))

        df['macdhigh'], df['macdsignal'], df['macdhist'] = ta.MACD(np.array(df['고가'].astype(float)), fastperiod=12, slowperiod=26, signalperiod=9)
        df['macdlow'], df['macdsignal'], df['macdhist'] = ta.MACD(np.array(df['저가'].astype(float)), fastperiod=12, slowperiod=26, signalperiod=9)
        df['macdclose'], df['macdsignal'], df['macdhist'] = ta.MACD(np.array(df['종가'].astype(float)), fastperiod=12, slowperiod=26, signalperiod=9)

        try:
            df['slowk'], df['slowd'] = ta.STOCH(np.array(df['macdhigh'].astype(float)),
                                                np.array(df['macdlow'].astype(float)),
                                                np.array(df['macdclose'].astype(float)),
                                                fastk_period=15, slowk_period=15, slowd_period=5)
        except Exception as e:
            logger.info("데이타부족 %s" % code)
            return None

        df.dropna(inplace=True)

        return df

    # Robot_Run이 되면 실행됨 - 매수/매도 종목을 리스트로 저장
    def 초기조건(self): # 종목 선정
        self.parent.statusbar.showMessage("[%s] 초기조건준비" % (self.sName))

        # 설정한 시총 상하한 사이 종목을 뽑아서 CODES에 저장 - 항목 : 시장구분, 종목코드, 종목명, 주식수, 감리구분, 상징일, 정일종가, 시가총액, 종목상태
        query = """
                SELECT
                    시장구분, 종목코드, 종목명, 주식수, 감리구분, 상장일, 전일종가,
                    CAST(((주식수 * 전일종가) / 100000000) AS UNSIGNED) AS 시가총액,
                    종목상태
                FROM
                    종목코드
                WHERE
                    ((시장구분 IN ('KOSPI' , 'KOSDAQ'))
                        AND (SUBSTR(종목코드, -1) = '0')
                        AND ((주식수 * 전일종가) between %s * (10000 * 10000) and %s * (10000 * 10000))
                        AND (NOT ((종목명 LIKE '%%스팩')))
                        AND (NOT ((종목명 LIKE '%%SPAC')))
                        AND (NOT ((종목상태 LIKE '%%관리종목%%')))
                        AND (NOT ((종목상태 LIKE '%%거래정지%%')))
                        AND (NOT ((감리구분 LIKE '%%투자경고%%')))
                        AND (NOT ((감리구분 LIKE '%%투자주의%%')))
                        AND (NOT ((감리구분 LIKE '%%환기종목%%')))
                        AND (NOT ((종목명 LIKE '%%ETN%%')))
                        AND (NOT ((종목명 LIKE '%%0호')))
                        AND (NOT ((종목명 LIKE '%%1호')))
                        AND (NOT ((종목명 LIKE '%%2호')))
                        AND (NOT ((종목명 LIKE '%%3호')))
                        AND (NOT ((종목명 LIKE '%%4호')))
                        AND (NOT ((종목명 LIKE '%%5호')))
                        AND (NOT ((종목명 LIKE '%%6호')))
                        AND (NOT ((종목명 LIKE '%%7호')))
                        AND (NOT ((종목명 LIKE '%%8호')))
                        AND (NOT ((종목명 LIKE '%%9호')))
                        AND (NOT ((종목명 LIKE '%%0')))
                        AND (NOT ((종목명 LIKE '%%1')))
                        AND (NOT ((종목명 LIKE '%%2')))
                        AND (NOT ((종목명 LIKE '%%3')))
                        AND (NOT ((종목명 LIKE '%%4')))
                        AND (NOT ((종목명 LIKE '%%5')))
                        AND (NOT ((종목명 LIKE '%%6')))
                        AND (NOT ((종목명 LIKE '%%7')))
                        AND (NOT ((종목명 LIKE '%%8')))
                        AND (NOT ((종목명 LIKE '%%9'))))
                ORDER BY 종목코드 ASC
            """ % (self.시총하한, self.시총상한)

        conn = sqliteconn()
        CODES = pdsql.read_sql_query(query, con=conn)
        conn.close()

        NOW = datetime.datetime.now()
        시작일자 = (NOW + datetime.timedelta(days=-366)).strftime('%Y-%m-%d')
        종료일자 = NOW.strftime('%Y-%m-%d')

        # 1년치 일봉데이터로 일자, 종가, 시가, 고가, 저가, 거래량저장해서 RSI등 계산해서 df로 반환
        pool = dict()
        for market, code, name, 주식수, 시가총액 in CODES[['시장구분', '종목코드', '종목명', '주식수', '시가총액']].values.tolist():
            df = self.get_price(code, 시작일자, 종료일자)
            if df is not None and len(df) > 0:
                pool[code] = df # Dictionary로 코드에 해당하는 일봉데이터를 재정리

        self.금일매도 = []
        self.매도할종목 = []
        self.매수할종목 = []

        for code, df in pool.items():

            try:
                종가D1, RSID1, macdD1, slowkD1, slowdD1 = df[['종가', 'RSI', 'macdclose', 'slowk', 'slowd']].values[-2] # 어제
                종가D0, RSID0, macdD0, slowkD0, slowdD0 = df[['종가', 'RSI', 'macdclose', 'slowk', 'slowd']].values[-1] # 오늘

                stock = self.portfolio.get(code) # 초기 로봇 실행 시 포트폴리오는 비어있음
                if stock != None: # 포트폴리오에 있으면

                    if RSID1 > 70.0 and RSID0 < 70.0:
                        self.매도할종목.append(code) # 포트폴리오에 있고, 매도 신호 포착시 매도종목리스트에 저장

                if stock == None: #포트폴리오에 없으면

                    if RSID1 < 30.0 and RSID0 > 30.0:
                        self.매수할종목.append(code) # 포트폴리오에 없고, 매수 신호 포착시 매수종목리스트에 저장

            except Exception as e:
                logger.info("데이타부족 %s" % code)
                print(df)

        pool = None # 초기화

    # 주문처리
    def 실시간데이타처리(self, param):
        if self.running == True:

            체결시간 = '%s %s:%s:%s' % (str(self.d), param['체결시간'][0:2], param['체결시간'][2:4], param['체결시간'][4:])
            종목코드 = param['종목코드']
            현재가 = abs(int(float(param['현재가'])))
            전일대비 = int(float(param['전일대비']))
            등락률 = float(param['등락률'])
            매도호가 = abs(int(float(param['매도호가'])))
            매수호가 = abs(int(float(param['매수호가'])))
            누적거래량 = abs(int(float(param['누적거래량'])))
            시가 = abs(int(float(param['시가'])))
            고가 = abs(int(float(param['고가'])))
            저가 = abs(int(float(param['저가'])))
            거래회전율 = abs(float(param['거래회전율']))
            시가총액 = abs(int(float(param['시가총액'])))

            # MainWindow의 __init__에서 CODE_POOL 변수 선언(self.CODE_POOL = self.get_code_pool()), pool[종목코드] = [시장구분, 종목명, 주식수, 시가총액]
            종목명 = self.parent.CODE_POOL[종목코드][1]

            self.parent.statusbar.showMessage("[%s] %s %s %s %s" % (체결시간, 종목코드, 종목명, 현재가, 전일대비))

            if 종목코드 in self.매도할종목:
                if self.portfolio.get(종목코드) is not None and self.주문실행중_Lock.get('S_%s' % 종목코드) is None:
                    (result, order) = self.정량매도(sRQName='S_%s' % 종목코드, 종목코드=종목코드, 매도가=현재가, 수량=self.portfolio[종목코드].수량)
                    if result == True:
                        self.주문실행중_Lock['S_%s' % 종목코드] = True
                        logger.debug('정량매도 : sRQName=%s, 종목코드=%s, 매도가=%s, 수량=%s' % (
                        'S_%s' % 종목코드, 종목코드, 현재가, self.portfolio[종목코드].수량))
                    else:
                        logger.debug('정량매도실패 : sRQName=%s, 종목코드=%s, 매도가=%s, 수량=%s' % (
                        'S_%s' % 종목코드, 종목코드, 현재가, self.portfolio[종목코드].수량))

            # 매수할 종목에 대해서 정액매수 주문하고 포트폴리오 저장
            if 종목코드 in self.매수할종목 and 종목코드 not in self.금일매도:
                if len(self.portfolio) < self.포트폴리오수 and self.portfolio.get(종목코드) is None and self.주문실행중_Lock.get('B_%s' % 종목코드) is None:
                    if 현재가 < (현재가 - 전일대비):
                        (result, order) = self.정액매수(sRQName='B_%s' % 종목코드, 종목코드=종목코드, 매수가=현재가, 매수금액=self.단위투자금)
                        if result == True:
                            self.portfolio[종목코드] = CPortStock(종목코드=종목코드, 종목명=종목명, 매수가=현재가, 매도가1차=0, 매도가2차=0, 손절가=0,
                                                              수량=0, 매수일=datetime.datetime.now())
                            self.주문실행중_Lock['B_%s' % 종목코드] = True
                            logger.debug(
                                '정액매수 : sRQName=%s, 종목코드=%s, 매수가=%s, 단위투자금=%s' % ('B_%s' % 종목코드, 종목코드, 현재가, self.단위투자금))
                        else:
                            logger.debug('정액매수실패 : sRQName=%s, 종목코드=%s, 매수가=%s, 단위투자금=%s' % (
                            'B_%s' % 종목코드, 종목코드, 현재가, self.단위투자금))

    def 접수처리(self, param):
        pass

    # OnReceiveChejanData에서 체결처리가 되면 체결처리 호출
    def 체결처리(self, param):
        종목코드 = param['종목코드']
        주문번호 = param['주문번호']
        self.주문결과[주문번호] = param

        # 매수
        if param['매도수구분'] == '2':
            주문수량 = int(param['주문수량'])
            미체결수량 = int(param['미체결수량'])
            if self.주문번호_주문_매핑.get(주문번호) is not None:
                주문 = self.주문번호_주문_매핑[주문번호]
                매수가 = int(주문[2:])
                단위체결가 = int(0 if (param['단위체결가'] is None or param['단위체결가'] == '') else param['단위체결가'])

                # logger.debug('매수-------> %s %s %s %s %s' % (param['종목코드'], param['종목명'], 매수가, 주문수량 - 미체결수량, 미체결수량))

                P = self.portfolio.get(종목코드) # 실시간데이터 처리에서 저장함
                if P is not None:
                    P.종목명 = param['종목명']
                    P.매수가 = 단위체결가
                    P.수량 = 주문수량 - 미체결수량
                else:
                    logger.debug('ERROR : 포트에 종목이 없음 !!!!')

                if 미체결수량 == 0:
                    try:
                        self.주문실행중_Lock.pop(주문)
                        # logger.info('POP성공 %s ' % 주문)
                    except Exception as e:
                        # logger.info('POP에러 %s ' % 주문)
                        pass

        if param['매도수구분'] == '1':  # 매도
            주문수량 = int(param['주문수량'])
            미체결수량 = int(param['미체결수량'])
            if self.주문번호_주문_매핑.get(주문번호) is not None:
                주문 = self.주문번호_주문_매핑[주문번호]
                매수가 = int(주문[2:])

                if 미체결수량 == 0:
                    try:
                        self.portfolio.pop(종목코드) # 매도가 완료되면 포트폴리오에서 삭제
                        # logger.info('포트폴리오POP성공 %s ' % 종목코드)
                        self.금일매도.append(종목코드)
                    except Exception as e:
                        # logger.info('포트폴리오POP에러 %s ' % 종목코드)
                        pass

                    try:
                        self.주문실행중_Lock.pop(주문)
                        # logger.info('POP성공 %s ' % 주문)
                    except Exception as e:
                        # logger.info('POP에러 %s ' % 주문)
                        pass
                else:
                    # logger.debug('매도-------> %s %s %s %s %s' % (param['종목코드'], param['종목명'], 매수가, 주문수량 - 미체결수량, 미체결수량))
                    P = self.portfolio.get(종목코드)
                    if P is not None:
                        P.종목명 = param['종목명']
                        P.수량 = 미체결수량

        # 메인 화면에 반영
        self.parent.RobotView()

    def 잔고처리(self, param):
        pass

    def Run(self, flag=True, sAccount=None):
        self.running = flag

        ret = 0
        if flag == True:
            self.sAccount = sAccount
            if self.sAccount is None:
                self.KiwoomAccount()
                self.sAccount = self.account[0]

            self.주문결과 = dict()
            self.주문번호_주문_매핑 = dict()
            self.주문실행중_Lock = dict()

            self.초기조건()

            self.실시간종목리스트 = self.매도할종목 + self.매수할종목 + list(self.portfolio.keys())

            logger.debug("오늘 거래 종목 : %s %s" % (self.sName, ';'.join(self.실시간종목리스트) + ';'))
            self.KiwoomConnect()

            # 실시간종목리스트에 저장된 종목에 대해서 실시간으로 종목코드, 체결시간, 현재가, 전일대비, 등략률, 매도호가, 매수호가, 누적거래량, 시가, 고가, 저가, 거래회전율, 시가총액받음
            # param 변수로 받아서 실시간데이터처리 함수 실행
            if len(self.실시간종목리스트) > 0:
                ret = self.KiwoomSetRealReg(self.sScreenNo, ';'.join(self.실시간종목리스트) + ';')
                logger.debug("실시간데이타요청 등록결과 %s" % ret)
        else:
            ret = self.KiwoomSetRealRemove(self.sScreenNo, 'ALL')
            self.KiwoomDisConnect()


## TradeSuperValue
Ui_TradeSuperValue, QtBaseClass_TradeSuperValue = uic.loadUiType("./UI/TradeSuperValue.ui")
class 화면_TradeSuperValue(QDialog, Ui_TradeSuperValue):
    def __init__(self, parent):
        super(화면_TradeSuperValue, self).__init__(parent)
        self.setupUi(self)


        self.parent = parent

        self.model = PandasModel()
        self.tableView.setModel(self.model)

        # self.columns = ['종목코드', '종목명']
        # self.보이는컬럼 = ['종목코드', '종목명']

        self.result = []

    # "종목 import"버튼 클릭 시 실행됨(시그널/슬롯 추가)
    def inquiry(self):
        # Google spreadsheet 사용
        try:
            self.data = import_googlesheet()
            print(self.data)
            if '_' in self.lineEdit_name.text():
                strategy = self.lineEdit_name.text().split('_')[0]
                self.data = self.data[self.data['매수전략'] == strategy]

            self.model.update(self.data)

            for i in range(len(self.data)):
                self.tableView.resizeColumnToContents(i)

        except Exception as e:
            print('화면_TradeSuperValue : inquiry Error ', e)
            logger.info('화면_TradeSuperValue : inquiry Error ', e)

class CTradeSuperValue(CTrade):  # 로봇 추가 시 __init__ : 복사, Setting, 초기조건:전략에 맞게, 데이터처리~Run:복사
    def __init__(self, sName, UUID, kiwoom=None, parent=None):
        self.sName = sName
        self.UUID = UUID

        self.sAccount = None
        self.kiwoom = kiwoom
        self.parent = parent

        self.running = False

        self.주문결과 = dict()
        self.주문번호_주문_매핑 = dict()
        self.주문실행중_Lock = dict()

        self.portfolio = dict()

        self.실시간종목리스트 = []

        self.SmallScreenNumber = 9999

        self.d = today

    # google spreadsheet 포트폴리오에 추가하기 위한 dataframe 생성
    def save_port(self):
        codes = list(self.portfolio.keys())

        data_dict = {
            '종목코드' : [],
            '종목명' : [],
            '매수가' : [],
            '목표가' : [],
            '수량' : [],
            '매수일': []
        }

        for code in codes:
            data_dict['종목코드'].append(code)
            data_dict['종목명'].append(self.portfolio[code].종목명)
            data_dict['매수가'].append(self.portfolio[code].매수가)
            data_dict['목표가'].append(self.portfolio[code].매도가1차)
            data_dict['수량'].append(self.portfolio[code].수량)
            data_dict['매수일'].append(self.portfolio[code].매수일)

        df = pd.DataFrame(data_dict)
        df['거래로봇'] = self.sName

        return df

    # google spreadsheet 종목 선정에서 체결 완료 시 해당 매수/매도가에 글씨체 진하게+배경 노란색 적용
    def gspread_update(self, code, price):
        pos = alpha_list[Stocklist['컬럼명'].index(price)-1] + str(Stocklist[code]['번호']+1)
        data_sheet.format(pos, {'textFormat': {'bold': True},
                                "backgroundColor": {
                                    "red": 1.0,
                                    "green": 1.0,
                                    "blue": 0.0
                                }})

    # 구글 스프레드시트에서 읽은 DataFrame에서 로봇별 종목리스트 셋팅
    def set_stocklist(self, data):
        Stocklist = dict()
        Stocklist['컬럼명'] = list(data.columns)
        for 종목코드 in data['종목코드'].unique():
            temp_list = data[data['종목코드'] == 종목코드].values[0]
            Stocklist[종목코드] = {
                '번호' : int(temp_list[0]),
                '종목명': temp_list[1],
                '종목코드': 종목코드,
                '매수전략': temp_list[3],
                '매도전략': temp_list[7],
                '매수가': list(int(temp_list[list(data.columns).index(col)]) for col in data.columns if
                                               '매수가' in col and temp_list[list(data.columns).index(col)] != ''),
                '매도가': list(int(temp_list[list(data.columns).index(col)]) for col in data.columns if
                                               '매도가' in col and temp_list[list(data.columns).index(col)] != '')
            }
        return Stocklist

    # 매수 전략별 매수 조건 확인
    def buy_strategy(self, code, price):
        result = False
        condition = 0
        strategy = self.Stocklist[code]['매수전략']
        현재가, 시가, 고가, 저가, 전일종가 = price # 시세 = [현재가, 시가, 고가, 저가, 전일종가]

        if strategy == '10':
            매수가 = self.Stocklist[code]['매수가'][0]
            시가위치하한 = 매수가 * (1 - self.Stocklist['전략']['시가위치'][0] / 100)
            시가위치상한 = 매수가 * (1 + self.Stocklist['전략']['시가위치'][1] / 100)

            if current_time < self.Stocklist['전략']['모니터링종료시간']:
                if 시가위치하한 <= 시가 and 시가 <= 시가위치상한 and 현재가 == 매수가:
                    print('조건 1')
                    result = True
                    condition = 1
                elif 매수가 * 1.05 <= 시가 and 전일종가 <= 시가 and 현재가 == 전일종가:
                    print('조건 2')
                    result = True
                    condition = 2
            else:
                pass

        elif strategy == '5':
            pass

        elif strategy == '3':
            pass

        return result, condition

    # 매도 전략별 매도 조건 확인
    def sell_strategy(self, code, price):
        strategy = self.Stocklist[code]['매수전략']
        현재가, 시가, 고가, 저가, 전일종가 = price  # 시세 = [현재가, 시가, 고가, 저가, 전일종가]
        # self.portfolio[종목코드].매도가
        매수가 = self.portfolio[code].매수가
        if len(self.portfolio[code].매도가) > 0:
            매도가 = self.portfolio[code].매도가
        else:
            매도가 = 0


        if strategy == '10':
            pass

        elif strategy == '5':
            pass

    # RobotAdd 함수에서 초기화 다음 셋팅 실행해서 설정값 넘김
    # def Setting(self, sScreenNo, 단위투자금=50 * 10000, 매수방법='00', 목표율=5, 손절율=3, 최대보유일=3, 매도방법='00', 종목리스트=pd.DataFrame()):
    def Setting(self, sScreenNo, 매수방법='00',매도방법='00', 종목리스트=pd.DataFrame()):
        try:
            self.sScreenNo = sScreenNo
            self.실시간종목리스트 = []
            self.매수방법 = 매수방법
            self.매도방법 = 매도방법
            self.종목리스트 = 종목리스트
            # self.단위투자비율 = 10

            self.Stocklist = self.set_stocklist(self.종목리스트)
            self.Stocklist['전략'] = {
                '단위투자금': '',
                '모니터링종료시간': '',
                '보유일': '',
                '시가위치': '',
                '구간별매도조건': []
            }

            row_data = strategy_sheet.get_all_values()

            for data in row_data:
                if data[0] == '단위투자금':
                    self.Stocklist['전략']['단위투자금'] = int(data[1])
                elif data[0] == '매수모니터링 종료시간':
                    self.Stocklist['전략']['모니터링종료시간'] = data[1] + ':00'
                elif data[0] == '보유일':
                    self.Stocklist['전략']['보유일'] = data[1]
                elif data[0] == '손절율':
                    self.Stocklist['전략']['구간별매도조건'].append(float(data[1][:-1]))
                elif data[0] == '시가 위치':
                    self.Stocklist['전략']['시가위치'] = list(map(int, data[1].split(',')))
                elif '구간' in data[0]:
                    self.Stocklist['전략']['구간별매도조건'].append(float(data[1][:-1]))

            self.Stocklist['전략']['구간별매도조건'].insert(1, 0.3)
            self.단위투자금 = self.Stocklist['전략']['단위투자금']

            print(self.Stocklist)
        except Exception as e:
            print('CTradeSuperValue_Setting Error :', e)
            Telegram('[XTrader]CTradeSuperValue_Setting Error : ', e)
            logger.info('CTradeSuperValue_Setting Error :', e)

    # Robot_Run이 되면 실행됨 - 매수/매도 종목을 리스트로 저장
    def 초기조건(self, codes):
        self.parent.statusbar.showMessage("[%s] 초기조건준비" % (self.sName))

        self.금일매도 = []
        self.매도할종목 = []
        self.매수할종목 = []

        for code in codes:
            stock = self.portfolio.get(code)
            if stock != None:  # 포트폴리오에 있으면
                # print(stock)
                self.매도할종목.append(code)
            else:  # 포트폴리오에 없으면
                # print(code)
                self.매수할종목.append(code)

        # for stock in df_keeplist['종목번호'].values: # 보유 종목 체크해서 매도 종목에 추가 → 로봇이 두개 이상일 경우 중복되므로 미적용
        #     self.매도할종목.append(stock)
            # 종목명 = df_keeplist[df_keeplist['종목번호']==stock]['종목명'].values[0]
            # 매입가 = df_keeplist[df_keeplist['종목번호']==stock]['매입가'].values[0]
            # 보유수량 = df_keeplist[df_keeplist['종목번호']==stock]['보유수량'].values[0]
            # print('종목코드 : %s, 종목명 : %s, 매입가 : %s, 보유수량 : %s' %(stock, 종목명, 매입가, 보유수량))
            # self.portfolio[stock] = CPortStock(종목코드=stock, 종목명=종목명, 매수가=매입가, 수량=보유수량, 매수일='')

    def 실시간데이타처리(self, param):
        try:
            if self.running == True:

                체결시간 = '%s %s:%s:%s' % (str(self.d), param['체결시간'][0:2], param['체결시간'][2:4], param['체결시간'][4:])
                종목코드 = param['종목코드']
                현재가 = abs(int(float(param['현재가'])))
                전일대비 = int(float(param['전일대비']))
                등락률 = float(param['등락률'])
                매도호가 = abs(int(float(param['매도호가'])))
                매수호가 = abs(int(float(param['매수호가'])))
                누적거래량 = abs(int(float(param['누적거래량'])))
                시가 = abs(int(float(param['시가'])))
                고가 = abs(int(float(param['고가'])))
                저가 = abs(int(float(param['저가'])))
                거래회전율 = abs(float(param['거래회전율']))
                시가총액 = abs(int(float(param['시가총액'])))

                종목명 = self.parent.CODE_POOL[종목코드][1] # pool[종목코드] = [시장구분, 종목명, 주식수, 전일종가, 시가총액]
                전일종가 = self.parent.CODE_POOL[종목코드][3]
                시세 = [현재가, 시가, 고가, 저가, 전일종가]

                self.parent.statusbar.showMessage("[%s] %s %s %s %s" % (체결시간, 종목코드, 종목명, 현재가, 전일대비))

                if 종목코드 in self.매수할종목 and 종목코드 not in self.금일매도:
                    if self.portfolio.get(종목코드) is None and self.주문실행중_Lock.get('B_%s' % 종목코드) is None:
                        result, condition = self.buy_strategy(종목코드, 시세)
                        if result == True:
                            Telegram('[XTrader]정액매수 : 종목코드=%s, 종목명=%s, 매수가=%s, 매수조건=%s' % (종목코드, 종목명, 현재가, condition))

                if 종목코드 in self.매도할종목:
                    if self.portfolio.get(종목코드) is not None and self.주문실행중_Lock.get('S_%s' % 종목코드) is None and self.주문실행중_Lock.get('B_%s' % 종목코드) is None:
                        self.sell_strategy(self.portfolio[종목코드], 시세)

                """
                # 매도 주문
                if 종목코드 in self.매도할종목:
                    if self.portfolio.get(종목코드) is not None and self.주문실행중_Lock.get('S_%s' % 종목코드) is None and self.주문실행중_Lock.get('B_%s' % 종목코드) is None:
                        if (현재가 >= self.portfolio[종목코드].매도가1차) or (현재가 <= self.portfolio[종목코드].손절가):
                            Telegram('[XTrader]정량매도주문 : 종목코드=%s, 종목명=%s, 매도가=%s, 수량=%s' % (종목코드, 종목명, 현재가, self.portfolio[종목코드].수량))
                            logger.debug('정량매도 : sRQName=%s, 종목코드=%s, 매도가=%s, 수량=%s' % (
                                'S_%s' % 종목코드, 종목코드, 현재가, self.portfolio[종목코드].수량))

                            (result, order) = self.정량매도(sRQName='S_%s' % 종목코드, 종목코드=종목코드, 매도가=현재가,
                                                        수량=self.portfolio[종목코드].수량)
                            if result == True:
                                self.주문실행중_Lock['S_%s' % 종목코드] = True
                                logger.debug('정량매도 : sRQName=%s, 종목코드=%s, 매도가=%s, 수량=%s' % (
                                    'S_%s' % 종목코드, 종목코드, 현재가, self.portfolio[종목코드].수량))
                            else:
                                logger.debug('정량매도실패 : sRQName=%s, 종목코드=%s, 매도가=%s, 수량=%s' % (
                                    'S_%s' % 종목코드, 종목코드, 현재가, self.portfolio[종목코드].수량))

                # if 종목코드 in self.매도할종목:
                #     if self.portfolio.get(종목코드) is not None and self.주문실행중_Lock.get('S_%s' % 종목코드) is None:
                #         Telegram('[XTrader]정량매도 : 종목코드=%s, 종목명=%s, 매도가=%s, 수량=%s' % (종목코드, 종목명, 현재가, self.portfolio[종목코드].수량))
                #         logger.debug('정량매도 : sRQName=%s, 종목코드=%s, 매도가=%s, 수량=%s' % (
                #             'S_%s' % 종목코드, 종목코드, 현재가, self.portfolio[종목코드].수량))

                        # (result, order) = self.정량매도(sRQName='S_%s' % 종목코드, 종목코드=종목코드, 매도가=현재가,
                        #                             수량=self.portfolio[종목코드].수량)
                        # if result == True:
                        #     self.주문실행중_Lock['S_%s' % 종목코드] = True
                        #     logger.debug('정량매도 : sRQName=%s, 종목코드=%s, 매도가=%s, 수량=%s' % (
                        #         'S_%s' % 종목코드, 종목코드, 현재가, self.portfolio[종목코드].수량))
                        # else:
                        #     logger.debug('정량매도실패 : sRQName=%s, 종목코드=%s, 매도가=%s, 수량=%s' % (
                        #         'S_%s' % 종목코드, 종목코드, 현재가, self.portfolio[종목코드].수량))

                # 매수 주문
                if 종목코드 in self.매수할종목 and 종목코드 not in self.금일매도:
                    if self.portfolio.get(종목코드) is None and self.주문실행중_Lock.get('B_%s' % 종목코드) is None:
                        if 현재가 <= self.종목리스트[self.종목리스트['종목코드']==종목코드]['매수가'].values[0]: # 현재가가 설정한 매수가 이하일 경우
                            # Telegram('[XTrader]정액매수 : 종목코드=%s, 종목명=%s, 매수가=%s, 단위투자금=%s' % (종목코드, 종목명, 현재가, self.단위투자금))
                            # logger.debug(
                            #     '정액매수 : sRQName=%s, 종목코드=%s, 매수가=%s, 단위투자금=%s' % (
                            #         'B_%s' % 종목코드, 종목코드, 현재가, self.단위투자금))
                            # self.손절가 = int(현재가 * (1- (self.손절율 / 100)))
                            # self.목표가 = self.종목리스트[self.종목리스트['종목코드']==종목코드]['목표가'].values[0]
                            # self.portfolio[종목코드] = CPortStock(종목코드=종목코드, 종목명=종목명, 매수가=현재가, 매도가1차=self.목표가, 매도가2차=self.목표가,
                            #                                   손절가=self.손절가,
                            #                                       수량=0, 매수일=datetime.datetime.now())
                            #
                            # print(self.portfolio[종목코드].종목명, self.portfolio[종목코드].매수가, self.portfolio[종목코드].매도가1차, self.portfolio[종목코드].손절가)
                            # self.df_portfolio = self.save_port()
                            # d2g.upload(self.df_portfolio, spreadsheet_key, portfolio_sheet, credentials=credentials, row_names=True)

                            (result, order) = self.정액매수(sRQName='B_%s' % 종목코드, 종목코드=종목코드, 매수가=현재가, 매수금액=self.단위투자금)
                            if result == True:
                                self.손절가 = int(현재가 * (1 - (self.손절율 / 100)))
                                self.목표가 = self.종목리스트[self.종목리스트['종목코드'] == 종목코드]['목표가'].values[0]
                                self.portfolio[종목코드] = CPortStock(종목코드=종목코드, 종목명=종목명, 매수가=현재가, 매도가1차=self.목표가, 매도가2차=self.목표가,
                                                                  손절가=self.손절가,
                                                                  수량=0, 매수일=datetime.datetime.now())

                                self.주문실행중_Lock['B_%s' % 종목코드] = True
                                Telegram('[XTrader]정액매수주문 : 종목코드=%s, 종목명=%s, 매수가=%s, 단위투자금=%s' % (종목코드, 종목명, 현재가, self.단위투자금))
                                logger.debug(
                                    '정액매수 : sRQName=%s, 종목코드=%s, 종목명=%s, 매수가=%s, 단위투자금=%s' % (
                                    'B_%s' % 종목코드, 종목코드, 종목명, 현재가, self.단위투자금))
                            else:
                                logger.debug('정액매수실패 : sRQName=%s, 종목코드=%s, 종목명=%s, 매수가=%s, 단위투자금=%s' % (
                                    'B_%s' % 종목코드, 종목코드, 종목명, 현재가, self.단위투자금))
            """
        except Exception as e:
            print('CTradeSuperValue_실시간데이타처리', e)

    def 접수처리(self, param):
        pass

    def 체결처리(self, param):
        종목코드 = param['종목코드']
        주문번호 = param['주문번호']
        self.주문결과[주문번호] = param

        # 매수
        if param['매도수구분'] == '2':
            주문수량 = int(param['주문수량'])
            미체결수량 = int(param['미체결수량'])
            if self.주문번호_주문_매핑.get(주문번호) is not None:
                주문 = self.주문번호_주문_매핑[주문번호]
                매수가 = int(주문[2:])
                단위체결가 = int(0 if (param['단위체결가'] is None or param['단위체결가'] == '') else param['단위체결가'])

                # logger.debug('매수-------> %s %s %s %s %s' % (param['종목코드'], param['종목명'], 매수가, 주문수량 - 미체결수량, 미체결수량))

                P = self.portfolio.get(종목코드)
                if P is not None:
                    P.종목명 = param['종목명']
                    P.매수가 = 단위체결가
                    P.수량 = 주문수량 - 미체결수량
                    P.매수일 = datetime.datetime.now()
                    Telegram('[XTrader]매수체결완료_종목명:%s, 매수가:%s, 수량:%s' %(P.종목명, P.매수가, P.수량))
                else:
                    logger.debug('ERROR 포트에 종목이 없음 !!!!')

                if 미체결수량 == 0:
                    try:
                        self.주문실행중_Lock.pop(주문)
                        # logger.info('POP성공 %s ' % 주문)
                    except Exception as e:
                        # logger.info('POP에러 %s ' % 주문)
                        pass

        # 매도
        if param['매도수구분'] == '1':
            주문수량 = int(param['주문수량'])
            미체결수량 = int(param['미체결수량'])
            if self.주문번호_주문_매핑.get(주문번호) is not None:
                주문 = self.주문번호_주문_매핑[주문번호]
                매도가 = int(주문[2:])

                if 미체결수량 == 0:
                    try:
                        Telegram('[XTrader]매도체결완료_종목명:%s, 매도가:%s, 수량:%s' % (param['종목명'], 매도가, 주문수량))
                        logger.info('포트폴리오POP성공 %s ' % 종목코드)
                        self.portfolio.pop(종목코드)
                        self.금일매도.append(종목코드)
                    except Exception as e:
                        # logger.info('포트폴리오POP에러 %s ' % 종목코드)
                        pass

                    try:
                        self.주문실행중_Lock.pop(주문)
                        # logger.info('POP성공 %s ' % 주문)
                    except Exception as e:
                        # logger.info('POP에러 %s ' % 주문)
                        pass
                else:
                    # logger.debug('매도-------> %s %s %s %s %s' % (param['종목코드'], param['종목명'], 매수가, 주문수량 - 미체결수량, 미체결수량))
                    P = self.portfolio.get(종목코드)
                    if P is not None:
                        P.종목명 = param['종목명']
                        P.수량 = 미체결수량

        # 메인 화면에 반영
        self.parent.RobotView()

    def 잔고처리(self, param):
        pass

    def Run(self, flag=True, sAccount=None):
        self.running = flag
        ret = 0

        if flag == True:
            self.KiwoomConnect()
            try:
                Telegram("[XTrader]%s ROBOT 실행" % (self.sName))

                self.단위투자금 = self.Stocklist['전략']['단위투자금'] #floor(int(d2deposit.replace(",", "")) * self.단위투자비율 / 100) # floor : 소수점 버림
                print('D+2 예수금 : ', int(d2deposit.replace(",", "")))
                print('단위투자금 : ', self.단위투자금)
                print('로봇 수 : ', len(self.parent.robots))
                self.주문결과 = dict()
                self.주문번호_주문_매핑 = dict()
                self.주문실행중_Lock = dict()

                codes = list(self.Stocklist.keys())[1:-1]
                print('종목리스트 : ', codes)
                self.초기조건(codes)

                print("매도 : ", self.매도할종목)
                print("매수 : ", self.매수할종목)

                self.실시간종목리스트 = self.매도할종목 + self.매수할종목 + list(self.portfolio.keys())

                logger.debug("오늘 거래 종목 : %s %s" % (self.sName, ';'.join(self.실시간종목리스트) + ';'))

                if len(self.실시간종목리스트) > 0:
                    ret = self.KiwoomSetRealReg(self.sScreenNo, ';'.join(self.실시간종목리스트) + ';')
                    logger.debug("실시간데이타요청 등록결과 %s" % ret)

            except Exception as e:
                print('CTradeSuperValue_Run Error :', e)
                Telegram('[XTrader]CTradeSuperValue_Run Error :', e)
                logger.info('CTradeSuperValue_Run Error :', e)

        else:
            Telegram("[XTrader]%s ROBOT 실행 중지" % (self.sName))
            ret = self.KiwoomSetRealRemove(self.sScreenNo, 'ALL')
            self.KiwoomDisConnect()

            if self.portfolio is not None:
                for code in list(self.portfolio.keys()):
                    if self.portfolio[code].수량 == 0:
                        self.portfolio.pop(code)

                self.save_port()


## TradeCondition
Ui_TradeCondition, QtBaseClass_TradeCondition = uic.loadUiType("./UI/TradeCondition.ui")
class 화면_TradeCondition(QDialog, Ui_TradeCondition):
    # def __init__(self, parent):
    def __init__(self, sScreenNo, kiwoom=None, parent=None): #
        super(화면_TradeCondition, self).__init__(parent)
        # self.setAttribute(Qt.WA_DeleteOnClose) # 위젯이 닫힐때 내용 삭제하는 것으로 창이 닫힐때 정보를 저장해야되는 로봇 세팅 시에는 쓰면 에러남!!
        self.setupUi(self)
        print("화면_TradeCondition : __init__")
        self.sScreenNo = sScreenNo
        self.kiwoom = kiwoom #
        self.parent = parent

        self.model = PandasModel()
        self.tableView.setModel(self.model)

        self.columns = ['종목코드', '종목명']
        self.보이는컬럼 = ['종목코드', '종목명']

        self.result = []

        self.KiwoomConnect()
        self.GetCondition()

    # 저장된 조건 검색식 목록 읽음
    def GetCondition(self):

        try:
            print("화면_TradeCondition : GetCondition1")
            self.getConditionLoad()
            print("화면_TradeCondition : GetCondition2")

            self.df_condition = DataFrame()
            self.idx = []
            self.conName = []

            for index in self.condition.keys(): # condition은 dictionary
                # print(self.condition)

                self.idx.append(str(index))
                self.conName.append(self.condition[index])

                # self.sendCondition("0156", self.condition[index], index, 1)

            self.df_condition['Index'] = self.idx
            self.df_condition['Name'] = self.conName
            self.df_condition['Table'] = ">> 조건식 " + self.df_condition['Index'] + " : " + self.df_condition['Name']
            self.df_condition = self.df_condition.sort_values(by='Index').reset_index(drop=True) # 추가
            print(self.df_condition) # 추가
            self.comboBox_condition.clear()
            self.comboBox_condition.addItems(self.df_condition['Table'].values)

        except Exception as e:
            print("GetCondition_Error")
            print(e)

    # 종목 조건검색 요청 메서드
    def sendCondition(self, screenNo, conditionName, conditionIndex, isRealTime):
        print("화면_TradeCondition : sendCondition")
        """
        종목 조건검색 요청 메서드

        이 메서드로 얻고자 하는 것은 해당 조건에 맞는 종목코드이다.
        해당 종목에 대한 상세정보는 setRealReg() 메서드로 요청할 수 있다.
        요청이 실패하는 경우는, 해당 조건식이 없거나, 조건명과 인덱스가 맞지 않거나, 조회 횟수를 초과하는 경우 발생한다.

        조건검색에 대한 결과는
        1회성 조회의 경우, receiveTrCondition() 이벤트로 결과값이 전달되며
        실시간 조회의 경우, receiveTrCondition()과 receiveRealCondition() 이벤트로 결과값이 전달된다.

        :param screenNo: string
        :param conditionName: string - 조건식 이름
        :param conditionIndex: int - 조건식 인덱스
        :param isRealTime: int - 조건검색 조회구분(0: 1회성 조회, 1: 실시간 조회)
        """
        isRequest = self.kiwoom.dynamicCall("SendCondition(QString, QString, int, int",
                                     screenNo, conditionName, conditionIndex, isRealTime)

        # OnReceiveTrCondition() 이벤트 메서드에서 루프 종료
        self.conditionLoop = QEventLoop()
        self.conditionLoop.exec_()

    # 조건 검색 관련 ActiveX와 On시리즈와 붙임(콜백)
    def KiwoomConnect(self):
        print("화면_TradeCondition : KiwoomConnect")
        self.kiwoom.OnReceiveTrCondition[str, str, str, int, int].connect(self.OnReceiveTrCondition)
        self.kiwoom.OnReceiveConditionVer[int, str].connect(self.OnReceiveConditionVer)
        self.kiwoom.OnReceiveRealCondition[str, str, str, str].connect(self.OnReceiveRealCondition)

    # 조건 검색 관련 ActiveX와 On시리즈 연결 해제
    def KiwoomDisConnect(self):
        print("화면_TradeCondition : KiwoomDisConnect")
        self.kiwoom.OnReceiveTrCondition[str, str, str, int, int].disconnect(self.OnReceiveTrCondition)
        self.kiwoom.OnReceiveConditionVer[int, str].disconnect(self.OnReceiveConditionVer)
        self.kiwoom.OnReceiveRealCondition[str, str, str, str].disconnect(self.OnReceiveRealCondition)

    # 조건식 목록 요청 메서드
    def getConditionLoad(self):
        """ 조건식 목록 요청 메서드 """
        print("화면_TradeCondition : getConditionLoad")
        self.kiwoom.dynamicCall("GetConditionLoad()")

        # receiveConditionVer() 이벤트 메서드에서 루프 종료
        self.conditionLoop = QEventLoop()
        self.conditionLoop.exec_()

    # 조건식 목록 획득 메서드(조건식 목록을 딕셔너리로 리턴)
    def getConditionNameList(self):
        """
        조건식 획득 메서드

        조건식을 딕셔너리 형태로 반환합니다.
        이 메서드는 반드시 receiveConditionVer() 이벤트 메서드안에서 사용해야 합니다.

        :return: dict - {인덱스:조건명, 인덱스:조건명, ...}
        """
        print("화면_TradeCondition : getConditionNameList")
        data = self.kiwoom.dynamicCall("GetConditionNameList()")


        conditionList = data.split(';')
        del conditionList[-1]

        conditionDictionary = {}

        for condition in conditionList:
            key, value = condition.split('^')
            conditionDictionary[int(key)] = value

        return conditionDictionary

    # 조건검색 세부 종목 조회 요청시 발생되는 이벤트
    def OnReceiveTrCondition(self, sScrNo, strCodeList, strConditionName, nIndex, nNext):
        logger.debug('main:OnReceiveTrCondition [%s] [%s] [%s] [%s] [%s]' % (sScrNo, strCodeList, strConditionName, nIndex, nNext))
        print("화면_TradeCondition : OnReceiveTrCondition")
        """
        (1회성, 실시간) 종목 조건검색 요청시 발생되는 이벤트
        :param screenNo: string
        :param codes: string - 종목코드 목록(각 종목은 세미콜론으로 구분됨)
        :param conditionName: string - 조건식 이름
        :param conditionIndex: int - 조건식 인덱스
        :param inquiry: int - 조회구분(0: 남은데이터 없음, 2: 남은데이터 있음)
        """

        try:
            if strCodeList == "":
                return

            self.codeList = strCodeList.split(';')
            del self.codeList[-1]

            # print("종목개수: ", len(self.codeList))
            # print(self.codeList)

            for code in self.codeList:
                row = []
                # code.append(c)
                row.append(code)
                n = self.kiwoom.dynamicCall("GetMasterCodeName(QString)", code)
                # now = abs(int(self.kiwoom.dynamicCall("GetCommRealData(QString, int)", code, 10)))
                # name.append(n)
                row.append(n)
                # row.append(now)
                self.result.append(row)
            # self.df_con['종목코드'] = code
            # self.df_con['종목명'] = name
            # print(self.df_con)

            self.data =DataFrame(data=self.result, columns=self.보이는컬럼)
            print(self.data)
            self.model.update(self.data )
            # self.model.update(self.df_con)
            for i in range(len(self.columns)):
                self.tableView.resizeColumnToContents(i)

        finally:
            self.conditionLoop.exit()

    # 조건식 목록 요청에 대한 응답 이벤트
    def OnReceiveConditionVer(self, lRet, sMsg):
        logger.debug('main:OnReceiveConditionVer : [이벤트] 조건식 저장 [%s] [%s]' % (lRet, sMsg))
        print("화면_TradeCondition : OnReceiveConditionVer")
        """
        getConditionLoad() 메서드의 조건식 목록 요청에 대한 응답 이벤트

        :param receive: int - 응답결과(1: 성공, 나머지 실패)
        :param msg: string - 메세지
        """

        try:
            self.condition = self.getConditionNameList() # condition이 리턴되서 오면 GetCondition에서 condition 변수 사용 가능
            # print("조건식 개수: ", len(self.condition))

            # for key in self.condition.keys():
                # print("조건식: ", key, ": ", self.condition[key])

        except Exception as e:
            print("OnReceiveConditionVer_Error")

        finally:
            self.conditionLoop.exit()

        # print(self.conditionName)
        # self.kiwoom.dynamicCall("SendCondition(QString,QString, int, int)", '0156', '갭상승', 0, 0)

    # 실시간 종목 조건검색 요청시 발생되는 이벤트
    def OnReceiveRealCondition(self, sTrCode, strType, strConditionName, strConditionIndex):
        logger.debug('main:OnReceiveRealCondition [%s] [%s] [%s] [%s]' % (sTrCode, strType, strConditionName, strConditionIndex))
        print("화면_TradeCondition : OnReceiveRealCondition")
        """
        실시간 종목 조건검색 요청시 발생되는 이벤트

        :param code: string - 종목코드
        :param event: string - 이벤트종류("I": 종목편입, "D": 종목이탈)
        :param conditionName: string - 조건식 이름
        :param conditionIndex: string - 조건식 인덱스(여기서만 인덱스가 string 타입으로 전달됨)
        """

        print("[receiveRealCondition]")

        print("종목코드: ", sTrCode)
        print("이벤트: ", "종목편입" if strType == "I" else "종목이탈")

    # 조건식 종목 검색 버튼 클릭 시 실행됨(시그널/슬롯 추가)
    def inquiry(self):
        print("화면_TradeCondition : inquiry")
        self.result = []
        index = self.comboBox_condition.currentIndex() # currentIndex() : 현재 콤보박스에서 선택된 index를 받음 int형
        print(index, self.condition[index])
        self.sendCondition("0156", self.condition[index], index, 1)

    # def accept(self):
    #     self.KiwoomDisConnect()
    #     self.r.exit()

class CTradeCondition(CTrade):  # 로봇 추가 시 __init__ : 복사, Setting / 초기조건:전략에 맞게, 데이터처리 / Run:복사
    # 동작 순서
    # 1. Robot Add에서 화면에서 주요 파라미터 받고 호출되면서 __init__ 실행
    # 2. Setting 실행 후 로봇 추가 셋팅 완료
    # 3. Robot_Run이 되면 초기 조건 실행하여 매수/매도 종목을 리스트로 저장하고 Run 실행

    def __init__(self, sName, UUID, kiwoom=None, parent=None):
        print("CTradeCondition : __init__")
        self.sName = sName
        self.UUID = UUID

        self.sAccount = None
        self.kiwoom = kiwoom
        self.parent = parent

        self.running = False

        self.remained_data = True
        self.초기설정상태 = False

        self.주문결과 = dict()
        self.주문번호_주문_매핑 = dict()
        self.주문실행중_Lock = dict()

        self.portfolio = dict()

        self.CList = []

        self.실시간종목리스트 = []

        self.SmallScreenNumber = 9999

        self.d = today

        self.최대보유일 = 1
        self.매수가비율 = -1.5 # percent 전일 종가 대비
        self.익절 = 3 # percent
        self.손절 = -10 # percent

    # 조건식 선택에 의해서 투자금, 매수/도 방법, 포트폴리오 수, 검색 종목 등이 저장됨
    def Setting(self, sScreenNo, 포트폴리오수, 조건식인덱스, 조건식명, 단위투자비율, 매수방법, 매도방법, 종목리스트):
        print("CTradeCondition : Setting")
        self.sScreenNo = sScreenNo
        self.단위투자비율 = 단위투자비율
        self.매수방법 = 매수방법
        self.매도방법 = 매도방법
        self.포트폴리오수 = 포트폴리오수
        self.조건식인덱스 = 조건식인덱스
        self.조건식명 = 조건식명
        self.종목리스트=종목리스트

        # ACCOUNT_CNT = self.GetLoginInfo('ACCOUNT_CNT')
        # ACC_NO = self.GetLoginInfo('ACCNO')
        # self.account = ACC_NO.split(';')[0:-1]
        # self.sAccount = self.account[0]
        # print("계좌 번호 : %s" % self.sAccount)
        #
        # print('스크린 번호 : ', self.sScreenNo)
        # self.InquiryList(_repeat=0)
        # self.InquiryLoop = QEventLoop()  # 로봇에서 바로 쓸 수 있도록하기 위해서 계좌 조회해서 종목을 받고나서 루프해제시킴
        # self.InquiryLoop.exec_()
        # print(("잔고 종목 : %s" % self.CList))

        # self.KiwoomAccount()
        # self.sAccount = self.account[0]
        # print('계좌 : ', self.sAccount)
        # print(int(self.sAsset.replace(",", "")))

        print("조검검색 로봇 셋팅 완료")

    # Robot_Run이 되면 실행됨 - 매수/매도 종목을 리스트로 저장
    def 초기조건(self, codes):  # 종목 선정
        print("CTradeCondition : 초기조건")
        self.parent.statusbar.showMessage("[%s] 초기조건준비" % (self.sName))

        # 매수할 종목은 해당 조건에서 검색된 종목
        # 매도할 종목은 이미 매수가 되어 포트폴리오에 저장되어 있는 종목 중 익절, 손절, 보유일 조건에 만족하는 종목

        NOW = datetime.datetime.now()

        self.금일매도 = []
        self.매도할종목 = []
        self.매수할종목 = []

        self.InquiryList(_repeat=0)
        # print(("잔고 종목 : %s" % self.CList))

        for code in codes:
            stock = self.portfolio.get(code)  # 초기 로봇 실행 시 포트폴리오는 비어있음
            if stock != None:  # 포트폴리오에 있고, 매도 신호 포착시 매도종목리스트에 저장(익절, 손절, 보유일만기)
                self.매도할종목.append(code)
            else:  # 포트폴리오에 없으면 매수종목리스트에 저장
                self.매수할종목.append(code)

        for c_stock in self.CList: # 보유 종목 체크
            if c_stock not in codes:
                self.매도할종목.append(c_stock)

        self.KiwoomAccount()
        self.sAccount = self.account[0]
        self.단위투자금 = floor(int(self.sAsset.replace(",","")) * self.단위투자비율 / 100) # floor : 소수점버림
        print('D+2 예수금 : ', int(self.sAsset.replace(",","")))
        print('단위투자금 : ', self.단위투자금)

        self.초기설정상태 = True

    # 주문처리
    def 실시간데이타처리(self, param):
        if self.running == True:

            체결시간 = '%s %s:%s:%s' % (str(self.d), param['체결시간'][0:2], param['체결시간'][2:4], param['체결시간'][4:])
            종목코드 = param['종목코드']
            현재가 = abs(int(float(param['현재가'])))
            전일대비 = int(float(param['전일대비']))
            등락률 = float(param['등락률'])
            매도호가 = abs(int(float(param['매도호가'])))
            매수호가 = abs(int(float(param['매수호가'])))
            누적거래량 = abs(int(float(param['누적거래량'])))
            시가 = abs(int(float(param['시가'])))
            고가 = abs(int(float(param['고가'])))
            저가 = abs(int(float(param['저가'])))
            거래회전율 = abs(float(param['거래회전율']))
            시가총액 = abs(int(float(param['시가총액'])))
            전일종가 = 현재가 - 전일대비

            # MainWindow의 __init__에서 CODE_POOL 변수 선언(self.CODE_POOL = self.get_code_pool()), pool[종목코드] = [시장구분, 종목명, 주식수, 시가총액]
            종목명 = self.parent.CODE_POOL[종목코드][1]

            self.parent.statusbar.showMessage("[%s] %s %s %s %s" % (체결시간, 종목코드, 종목명, 현재가, 전일대비))


            if 종목코드 in self.매도할종목:
                """ 보유기간에 따른 만기일 매도 조건이 있을 경우"""
                # stock = self.portfolio.get(종목코드)
                # if stock is not None and self.주문실행중_Lock.get('S_%s' % 종목코드) is None:
                #     보유기간 = periodcal(stock.매수일)
                    # if 보유기간 >= self.최대보유일: # 만기일 매도
                    #     전일종가 = 현재가 - 전일대비
                    #     logger.info('종목 : %s -> 만기일 매도' % (종목코드))
                    #     (result, order) = self.정량매도(sRQName='S_%s' % 종목코드, 종목코드=종목코드, 매도가=전일종가, 수량=self.portfolio[종목코드].수량)
                if self.portfolio.get(종목코드) is not None and self.주문실행중_Lock.get('S_%s' % 종목코드) is None:
                    (result, order) = self.정량매도(sRQName='S_%s' % 종목코드, 종목코드=종목코드, 매도가=전일종가, 수량=self.portfolio[종목코드].수량)
                    if result == True:
                        self.주문실행중_Lock['S_%s' % 종목코드] = True
                        logger.debug('정량매도 : sRQName=%s, 종목코드=%s, 매도가=%s, 수량=%s' % (
                            'S_%s' % 종목코드, 종목코드, 현재가, self.portfolio[종목코드].수량))
                    else:
                        logger.debug('정량매도실패 : sRQName=%s, 종목코드=%s, 매도가=%s, 수량=%s' % (
                            'S_%s' % 종목코드, 종목코드, 현재가, self.portfolio[종목코드].수량))

            # 매수할 종목에 대해서 정액매수 주문하고 포트폴리오 저장
            if 종목코드 in self.매수할종목 and 종목코드 not in self.금일매도:
                if len(self.portfolio) < self.포트폴리오수 and self.portfolio.get(종목코드) is None and self.주문실행중_Lock.get('B_%s' % 종목코드) is None:
                    매수가 = 전일종가 * (1-(abs(self.매수가비율)/100))
                    (result, order) = self.정액매수(sRQName='B_%s' % 종목코드, 종목코드=종목코드, 매수가=매수가, 매수금액=self.단위투자금)
                    if result == True:
                        self.portfolio[종목코드] = CPortStock(종목코드=종목코드, 종목명=종목명, 매수가=현재가, 매도가1차=0, 매도가2차=0, 손절가=0,
                                                          수량=0, 매수일=datetime.datetime.now())
                        self.주문실행중_Lock['B_%s' % 종목코드] = True
                        logger.debug(
                            '정액매수 : sRQName=%s, 종목코드=%s, 매수가=%s, 단위투자금=%s' % ('B_%s' % 종목코드, 종목코드, 현재가, self.단위투자금))
                    else:
                        logger.debug('정액매수실패 : sRQName=%s, 종목코드=%s, 매수가=%s, 단위투자금=%s' % (
                            'B_%s' % 종목코드, 종목코드, 현재가, self.단위투자금))

    def 접수처리(self, param):
        pass

    # OnReceiveChejanData에서 체결처리가 되면 체결처리 호출
    def 체결처리(self, param):
        종목코드 = param['종목코드']
        주문번호 = param['주문번호']
        self.주문결과[주문번호] = param

        # 매수
        if param['매도수구분'] == '2':
            주문수량 = int(param['주문수량'])
            미체결수량 = int(param['미체결수량'])
            if self.주문번호_주문_매핑.get(주문번호) is not None:
                주문 = self.주문번호_주문_매핑[주문번호]
                매수가 = int(주문[2:])
                단위체결가 = int(0 if (param['단위체결가'] is None or param['단위체결가'] == '') else param['단위체결가'])

                # logger.debug('매수-------> %s %s %s %s %s' % (param['종목코드'], param['종목명'], 매수가, 주문수량 - 미체결수량, 미체결수량))

                P = self.portfolio.get(종목코드)  # 실시간데이터 처리에서 매수 주문 후 저장함
                if P is not None:
                    P.종목명 = param['종목명']
                    P.매수가 = 단위체결가
                    P.수량 = 주문수량 - 미체결수량
                else:
                    logger.debug('ERROR : 포트에 종목이 없음 !!!!')

                if 미체결수량 == 0:
                    try:
                        self.주문실행중_Lock.pop(주문)
                        # logger.info('POP성공 %s ' % 주문)
                    except Exception as e:
                        # logger.info('POP에러 %s ' % 주문)
                        pass

        # 매도
        if param['매도수구분'] == '1':
            주문수량 = int(param['주문수량'])
            미체결수량 = int(param['미체결수량'])
            if self.주문번호_주문_매핑.get(주문번호) is not None:
                주문 = self.주문번호_주문_매핑[주문번호]
                매수가 = int(주문[2:])

                if 미체결수량 == 0:
                    try:
                        self.portfolio.pop(종목코드)  # 매도가 완료되면 포트폴리오에서 삭제
                        # logger.info('포트폴리오POP성공 %s ' % 종목코드)
                        self.금일매도.append(종목코드)
                    except Exception as e:
                        # logger.info('포트폴리오POP에러 %s ' % 종목코드)
                        pass

                    try:
                        self.주문실행중_Lock.pop(주문)
                        # logger.info('POP성공 %s ' % 주문)
                    except Exception as e:
                        # logger.info('POP에러 %s ' % 주문)
                        pass
                else:
                    # logger.debug('매도-------> %s %s %s %s %s' % (param['종목코드'], param['종목명'], 매수가, 주문수량 - 미체결수량, 미체결수량))
                    P = self.portfolio.get(종목코드)
                    if P is not None:
                        P.종목명 = param['종목명']
                        P.수량 = 미체결수량

        # 메인 화면에 반영
        self.parent.RobotView()

    def 잔고처리(self, param):
        pass

    def Run(self, flag=True, sAccount=None):
        print("CTradeCondition : Run")
        try:
            self.running = flag

            ret = 0
            if flag == True:
                self.KiwoomConnect()
                logger.debug("조건식 거래 로봇 실행")
                self.sAccount = sAccount

                # if self.sAccount is None:
                #     self.KiwoomAccount()
                #     self.sAccount = self.account[0]
                #     self.단위투자금 = floor(int(self.sAsset.replace(",","")) * self.단위투자비율 / 100) # floor : 소수점버림
                #     print(self.단위투자금)

                self.주문결과 = dict()
                self.주문번호_주문_매핑 = dict()
                self.주문실행중_Lock = dict()

                print("조건식 인덱스 : ", self.조건식인덱스)
                print("조건식명 : ", self.조건식명)
                # self.GetCodes(self.조건식인덱스, self.조건식명)
                print("종목 : ", self.종목리스트)

                codes = self.종목리스트['종목코드'].values
                self.초기조건(codes)

                print("매도 : ", self.매도할종목)
                print("매수 : ", self.매수할종목)

                # logger.debug("로봇 검색식 : %s %s" % (self.조건식인덱스, self.조건식명))
                #
                # self.GetCodes(self.조건식인덱스, self.조건식명)
                #
                #
                self.실시간종목리스트 = self.매도할종목 + self.매수할종목 + list(self.portfolio.keys())
                #
                logger.debug("오늘 거래 종목 : %s %s" % (self.sName, ';'.join(self.실시간종목리스트) + ';'))
                # self.KiwoomConnect()
                #
                # # 실시간종목리스트에 저장된 종목에 대해서 실시간으로 종목코드, 체결시간, 현재가, 전일대비, 등략률, 매도호가, 매수호가, 누적거래량, 시가, 고가, 저가, 거래회전율, 시가총액받음
                # # param 변수로 받아서 실시간데이터처리 함수 실행
                # if len(self.실시간종목리스트) > 0:
                #     ret = self.KiwoomSetRealReg(self.sScreenNo, ';'.join(self.실시간종목리스트) + ';')
                #     logger.debug("실시간데이타요청 등록결과 %s" % ret)

            else:
                ret = self.KiwoomSetRealRemove(self.sScreenNo, 'ALL')
                self.KiwoomDisConnect()

        except Exception as e:
            print('CTradeCondition_run', e)


##################################################################################
# 메인
##################################################################################
Ui_MainWindow, QtBaseClass_MainWindow = uic.loadUiType("./UI/MainWindow.ui")
class MainWindow(QMainWindow, Ui_MainWindow):
    def __init__(self):
        # 화면을 보여주기 위한 코드
        print("MainWindow : __init__1")
        super().__init__()
        QMainWindow.__init__(self)
        Ui_MainWindow.__init__(self)
        self.setupUi(self)
        self.setWindowTitle("XTrader")

        # 현재 시간 받음
        self.시작시각 = datetime.datetime.now()

        # 메인윈도우가 뜨고 키움증권과 붙이기 위한 작업
        self.KiwoomAPI()          # 키움 ActiveX를 메모리에 올림
        self.KiwoomConnect()      # 메모리에 올라온 ActiveX와 내가 만든 함수 On시리즈와 연결(콜백 : 이벤트가 오면 나를 불러줘)
        self.ScreenNumber = 5000

        self.robots = []

        self.dialog = dict()
        # self.dialog['리얼데이타'] = None
        # self.dialog['계좌정보조회'] = None

        self.model = PandasModel()
        self.tableView_robot.setModel(self.model)
        self.tableView_robot.setSelectionBehavior(QTableView.SelectRows)
        self.tableView_robot.setSelectionMode(QTableView.SingleSelection)
        self.tableView_robot.pressed.connect(self.RobotCurrentIndex)
        # self.connect(self.tableView_robot.selectionModel(), SIGNAL("currentRowChanged(QModelIndex,QModelIndex)"), self.RobotCurrentIndex)
        self.tableView_robot_current_index = None

        self.portfolio_model = PandasModel()
        self.tableView_portfolio.setModel(self.portfolio_model)
        self.tableView_portfolio.setSelectionBehavior(QTableView.SelectRows)
        self.tableView_portfolio.setSelectionMode(QTableView.SingleSelection)
        self.portfolio_model.update((DataFrame(columns=['종목코드', '종목명', '매수가', '수량', '매수일'])))

        self.robot_columns = ['Robot타입', 'Robot명', 'RobotID', '스크린번호', '실행상태', '포트수', '포트폴리오']

        # TODO: 주문제한 설정
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.limit_per_second) # 초당 4번
        # QtCore.QObject.connect(self.timer, QtCore.SIGNAL("timeout()"), self.limit_per_second)
        self.timer.start(1000) # 1초마다 리셋

        self.주문제한 = 0
        self.조회제한 = 0
        self.금일백업작업중 = False
        self.종목선정작업중 = False

        self._login = False

        self.KiwoomLogin()  # 프로그램 실행 시 자동로그인
        self.CODE_POOL = self.get_code_pool() # DB 종목데이블에서 시장구분, 코드, 종목명, 주식수, 전일종가 읽어옴

    # DB에 저장된 상장 종목 코드 읽음
    def get_code_pool(self):
        print("MainWindow : get_code_pool")
        query = """
            select 시장구분, 종목코드, 종목명, 주식수, 전일종가, 전일종가*주식수 as 시가총액
            from 종목코드
            order by 시장구분, 종목코드
        """
        conn = sqliteconn()
        df = pd.read_sql(query, con=conn)
        conn.close()

        pool = dict()
        for idx, row in df.iterrows():
            시장구분, 종목코드, 종목명, 주식수, 전일종가, 시가총액 = row
            pool[종목코드] = [시장구분, 종목명, 주식수, 전일종가, 시가총액]
        return pool

    # 구글스프레드시트 종목 Import
    def update_googledata(self, check):
        data = import_googlesheet()

        if check == False:
            # 매수 전략 확인
            strategy_list = list(data['매수전략'].unique())

            # 로딩된 로봇을 robot_list에 저장
            robot_list = []
            for robot in self.robots:
                robot_list.append(robot.sName.split('_')[0])

            # 매수 전략별 로봇 자동 편집/추가
            for strategy in strategy_list:
                df_stock = data[data['매수전략'] == strategy]

                if strategy in robot_list:
                    print('로봇 편집')
                    Telegram('[XTrader]로봇 편집')
                    for robot in self.robots:
                        if robot.sName.split('_')[0] == strategy:
                            self.RobotAutoEdit_TradeSuperValue(robot, df_stock)
                            self.RobotView()
                            break
                else:
                    print('로봇 추가')
                    Telegram('[XTrader]로봇 추가')
                    self.RobotAutoAdd_TradeSuperValue(df_stock, strategy)
                    self.RobotView()

            Telegram('[XTrader]로봇 준비 완료')
            logger.info("로봇 준비 완료")

    """
    # 조건 검색식 읽어서 해당 종목 저장
    def GetCondition(self):
        # logger.info("조건 검색식 종목 읽기")
        print("MainWindow : GetCondition")
        self.kiwoom.OnReceiveTrCondition[str, str, str, int, int].connect(self.OnReceiveTrCondition)
        self.kiwoom.OnReceiveTrData[str, str, str, str, str, int, str, str, str].connect(self.OnReceiveTrData)
        self.kiwoom.OnReceiveConditionVer[int, str].connect(self.OnReceiveConditionVer)
        self.kiwoom.OnReceiveRealCondition[str, str, str, str].connect(self.OnReceiveRealCondition)

        try:
            self.getConditionLoad()

            self.conditionid = []
            self.conditionname = []

            for index in self.condition.keys(): # condition은 dictionary
                # print(self.condition)
                self.conditionid.append(str(index))
                self.conditionname.append(self.condition[index])
                # print('조건별 종목 검색 시작')
                self.sendCondition("0156", self.condition[index], index, 0)


        except Exception as e:
            print("GetCondition_Error")
            print(e)

        finally:
            # print(self.df_condition)
            logger.info("조건 검색식 종목 저장완료")
            self.kiwoom.OnReceiveTrCondition[str, str, str, int, int].disconnect(self.OnReceiveTrCondition)
            self.kiwoom.OnReceiveTrData[str, str, str, str, str, int, str, str, str].disconnect(self.OnReceiveTrData)
            self.kiwoom.OnReceiveConditionVer[int, str].disconnect(self.OnReceiveConditionVer)
            self.kiwoom.OnReceiveRealCondition[str, str, str, str].disconnect(self.OnReceiveRealCondition)
    
    # 조건식 목록 요청 메서드
    def getConditionLoad(self):
        print("MainWindow : getConditionLoad")
        self.kiwoom.dynamicCall("GetConditionLoad()")

        # receiveConditionVer() 이벤트 메서드에서 루프 종료
        self.conditionLoop = QEventLoop()
        self.conditionLoop.exec_()
    
    # 조건식 획득 메서드
    def getConditionNameList(self):
        # 조건식을 딕셔너리 형태로 반환합니다.
        # 이 메서드는 반드시 receiveConditionVer() 이벤트 메서드안에서 사용해야 합니다.
        # 
        # :return: dict - {인덱스:조건명, 인덱스:조건명, ...}
        
        print("MainWindow : getConditionNameList")
        data = self.kiwoom.dynamicCall("GetConditionNameList()")

        conditionList = data.split(';')
        del conditionList[-1]

        conditionDictionary = {}

        for condition in conditionList:
            key, value = condition.split('^')
            conditionDictionary[int(key)] = value

        return conditionDictionary
    
    # 종목 조건검색 요청 메서드
    def sendCondition(self, screenNo, conditionName, conditionIndex, isRealTime):
        # 이 메서드로 얻고자 하는 것은 해당 조건에 맞는 종목코드이다.
        # 해당 종목에 대한 상세정보는 setRealReg() 메서드로 요청할 수 있다.
        # 요청이 실패하는 경우는, 해당 조건식이 없거나, 조건명과 인덱스가 맞지 않거나, 조회 횟수를 초과하는 경우 발생한다.
        # 
        # 조건검색에 대한 결과는
        # 1회성 조회의 경우, receiveTrCondition() 이벤트로 결과값이 전달되며
        # 실시간 조회의 경우, receiveTrCondition()과 receiveRealCondition() 이벤트로 결과값이 전달된다.
        # 
        # :param screenNo: string
        # :param conditionName: string - 조건식 이름
        # :param conditionIndex: int - 조건식 인덱스
        # :param isRealTime: int - 조건검색 조회구분(0: 1회성 조회, 1: 실시간 조회)
        
        print("MainWindow : sendCondition")
        isRequest = self.kiwoom.dynamicCall("SendCondition(QString, QString, int, int",
                                     screenNo, conditionName, conditionIndex, isRealTime)

        # receiveTrCondition() 이벤트 메서드에서 루프 종료
        self.conditionLoop = QEventLoop()
        self.conditionLoop.exec_()
    """

    # 프로그램 실행 3초 후 저장된 로봇 정보받아옴
    def OnQApplicationStarted(self):
        print("MainWindow : OnQApplicationStarted")        
        global 로봇거래계좌번호

        try:
            with sqlite3.connect(DATABASE) as conn:
                cursor = conn.cursor()

                cursor.execute("select value from Setting where keyword='robotaccount'")

                for row in cursor.fetchall():
                    # _temp = base64.decodestring(row[0])  # base64에 text화해서 암호화 : DB에 잘 넣기 위함
                    _temp = base64.decodebytes(row[0])
                    로봇거래계좌번호 = pickle.loads(_temp)
                    # print(로봇거래계좌번호)
                cursor.execute('select uuid, strategy, name, robot from Robots')

                self.robots = []
                for row in cursor.fetchall():
                    uuid, strategy, name, robot_encoded = row
                    robot = base64.decodebytes(robot_encoded)
                    # r = base64.decodebytes(robot_encoded)
                    r = pickle.loads(robot)

                    r.kiwoom = self.kiwoom
                    r.parent = self
                    r.d = today

                    r.running = False
                    # logger.debug(r.sName, r.UUID, len(r.portfolio))
                    self.robots.append(r)

        except Exception as e:
            print('OnQApplicationStarted', e)

        self.RobotView()

    # 프로그램 실행 후 1초 마다 실행 : 조건에 맞는 시간이 되면 백업 시작
    def OnClockTick(self):
        current = datetime.datetime.now()
        global current_time
        current_time = current.strftime('%H:%M:%S')
        savetime_list = ['08:50:00']  # 지정된 시간에 종목 저장하기 위함
        """savetime_list = ['09:00:00', '09:05:00', '09:10:00', '09:30:00', '10:00:00', '11:00:00', '12:00:00',
                         '13:00:00', '14:00:00', '15:00:00', '15:10:00', '15:15:00'] # 지정된 시간에 종목 저장하기 위함"""
        workday_list = [0, 1, 2, 3, 4] # 평일만 저장
        # print(current.strftime('%H:%M:%S'))

        # if '08:30:00' <= current_time and current_time < '08:30:30':
        #     if len(self.robots) > 0:
        #         for r in self.robots:
        #             if r.running == False:  # 로봇이 실행중이 아니면
        #                 self.RobotRun()
        #                 self.RobotView()
        #
        # elif '15:30:00' < current_time and current_time < '15:30:30':
        #     if len(self.robots) > 0:
        #         for r in self.robots:
        #             if r.running == True:  # 로봇이 실행중이면
        #                 self.RobotStop()
        #                 self.RobotView()

        if '15:36:00' < current_time and current_time < '15:36:59' and self.금일백업작업중 == False and self._login == True:# and current.weekday() == 4:
        # 수능일이면 아래 시간 조건으로 수정
        # if '17:00:00' < current.strftime('%H:%M:%S') and current.strftime('%H:%M:%S') < '17:00:59' and self.금일백업작업중 == False and self._login == True:
        #     self.금일백업작업중 = True
        #     self.Backup(작업=None)
            pass

        # 8시 32분 : 종목 데이블 생성
        if current_time == '08:32:00':
            Telegram('[XTrader]종목테이블 생성')
            self.StockCodeBuild(to_db=True)
            self.CODE_POOL = self.get_code_pool()  # DB 종목데이블에서 시장구분, 코드, 종목명, 주식수, 전일종가 읽어옴

        # 8시 35분 : 구글 시트 오류 체크 시작
        if current_time == '08:35:00':
            Telegram('[XTrader]구글 시트 오류 체크 시작')
            self.checkclock = QTimer(self)
            self.checkclock.timeout.connect(self.OnGoogleCheck)  # 5분마다 구글 시트 읽음 : MainWindow.OnGoogleCheck 실행
            self.checkclock.start(300000)  # 300000초마다 타이머 작동

        # 8시 59분 : 구글 시트 종목 Import
        if current_time == '08:59:00':
            Telegram('[XTrader]구글 시트 오류 체크 중지')
            self.checkclock.stop()

            Telegram('[XTrader]구글시트 Import')
            self.update_googledata(check=False)

        # 8시 59분 30초 : 로봇 실행
        if '08:59:30' <= current_time and current_time < '08:59:50':
            if len(self.robots) > 0:
                for r in self.robots:
                    if r.running == False:  # 로봇이 실행중이 아니면
                        self.RobotRun()
                        self.RobotView()

        # 로봇을 저장
        # if self.시작시각.strftime('%H:%M:%S') > '08:00:00' and self.시작시각.strftime('%H:%M:%S') < '15:30:00' and current.strftime('%H:%M:%S') > '01:00:00':
        #     if len(self.robots) > 0:
        #         self.RobotSave()

        #     for k in self.dialog:
        #         self.dialog[k].KiwoomDisConnect()
        #         try:
        #             self.dialog[k].close()
        #         except Exception as e:
        #             pass

        #     self.close()

        # 지정 시간에 로봇을 중지한다던가 원하는 실행을 아래 pass에 작성
        """elif current_time < '15:30:00': # 장 마감 이전
            if current.weekday() in workday_list: # 주중인지 확인
                if current_time in savetime_list: # 지정된 시간인지 확인
                    logger.info("조건검색식 타이머 작동")
                    Telegram(str(current)[:-7] + " : " + "조건검색식 종목 검색")
                    self.GetCondition()  # 조건검색식을 모두 읽어서 해당하는 종목 저장"""
            # if current.second == 0:  # 매 0초
            #     # if current.minute % 10 == 0:  # 매 10 분
            #     if current.minute == 1 or current.strftime('%H:%M:%S') == '09:30:00' or current.strftime('%H:%M:%S') == '15:15:00':  # 매시 1분
            #         logger.info("조건검색식 타이머 작동")
            #         Telegram(str(current)[:-7] + " : " + "조건검색식 종목 검색")
            #         # print(current.minute, current.second)
            #         self.GetCondition() # 조건검색식을 모두 읽어서 해당하는 종목 저장
                    # for r in self.robots:
                    #     if r.running == True:  # 로봇이 실행중이면
                    #         # print(r.sName, r.running)
                    #         pass

    # 주문 제한 초기화
    def limit_per_second(self):
        self.주문제한 = 0
        self.조회제한 = 0
        #logger.info("초당제한 주문 클리어")

    # 프로그램 실행 후 5분 마다 실행 : 구글 스프레드 시트 오류 확인
    def OnGoogleCheck(self):
        print('OnGoogleCheck')
        self.update_googledata(check=True)

    # 메인 윈도우에서의 모든 액션에 대한 처리
    def MENU_Action(self, qaction):
        logger.debug("Action Slot %s %s " % (qaction.objectName(), qaction.text()))
        _action = qaction.objectName()
        if _action == "actionExit":
            if len(self.robots) > 0:
                self.RobotSave()

            for k in self.dialog:
                self.dialog[k].KiwoomDisConnect()
                try:
                    self.dialog[k].close()
                except Exception as e:
                    pass

            self.close()
        elif _action == "actionLogin":
            self.KiwoomLogin()
        elif _action == "actionLogout":
            self.KiwoomLogout()
        elif _action == "actionDailyPrice":
            # self.F_dailyprice()
            if self.dialog.get('일자별주가') is not None:
                try:
                    self.dialog['일자별주가'].show()
                except Exception as e:
                    self.dialog['일자별주가'] = 화면_일별주가(sScreenNo=9902, kiwoom=self.kiwoom, parent=self)
                    self.dialog['일자별주가'].KiwoomConnect()
                    self.dialog['일자별주가'].show()
            else:
                self.dialog['일자별주가'] = 화면_일별주가(sScreenNo=9902, kiwoom=self.kiwoom, parent=self)
                self.dialog['일자별주가'].KiwoomConnect()
                self.dialog['일자별주가'].show()
        elif _action == "actionMinuitePrice":
            # self.F_minprice()
            if self.dialog.get('분별주가') is not None:
                try:
                    self.dialog['분별주가'].show()
                except Exception as e:
                    self.dialog['분별주가'] = 화면_분별주가(sScreenNo=9903, kiwoom=self.kiwoom, parent=self)
                    self.dialog['분별주가'].KiwoomConnect()
                    self.dialog['분별주가'].show()
            else:
                self.dialog['분별주가'] = 화면_분별주가(sScreenNo=9903, kiwoom=self.kiwoom, parent=self)
                self.dialog['분별주가'].KiwoomConnect()
                self.dialog['분별주가'].show()
        elif _action == "actionInvestors":
            # self.F_investor()
            if self.dialog.get('종목별투자자') is not None:
                try:
                    self.dialog['종목별투자자'].show()
                except Exception as e:
                    self.dialog['종목별투자자'] = 화면_종목별투자자(sScreenNo=9904, kiwoom=self.kiwoom, parent=self)
                    self.dialog['종목별투자자'].KiwoomConnect()
                    self.dialog['종목별투자자'].show()
            else:
                self.dialog['종목별투자자'] = 화면_종목별투자자(sScreenNo=9904, kiwoom=self.kiwoom, parent=self)
                self.dialog['종목별투자자'].KiwoomConnect()
                self.dialog['종목별투자자'].show()
        elif _action == "actionRealDataDialog":
            _code = '122630;114800'
            if self.dialog.get('리얼데이타') is not None:
                try:
                    self.dialog['리얼데이타'].show()
                except Exception as e:
                    self.dialog['리얼데이타'] = 화면_실시간정보(sScreenNo=9901, kiwoom=self.kiwoom, parent=self)
                    self.dialog['리얼데이타'].KiwoomConnect()
                    _screenno = self.dialog['리얼데이타'].sScreenNo
                    self.dialog['리얼데이타'].KiwoomSetRealRemove(_screenno, _code)
                    self.dialog['리얼데이타'].KiwoomSetRealReg(_screenno, _code, sRealType='0')
                    self.dialog['리얼데이타'].show()
            else:
                self.dialog['리얼데이타'] = 화면_실시간정보(sScreenNo=9901, kiwoom=self.kiwoom, parent=self)
                self.dialog['리얼데이타'].KiwoomConnect()
                _screenno = self.dialog['리얼데이타'].sScreenNo
                self.dialog['리얼데이타'].KiwoomSetRealRemove(_screenno, _code)
                self.dialog['리얼데이타'].KiwoomSetRealReg(_screenno, _code, sRealType='0')
                self.dialog['리얼데이타'].show()
        elif _action == "actionAccountDialog": # 계좌정보조회
            if self.dialog.get('계좌정보조회') is not None: # dialog : __init__()에 dict로 정의됨
                try:
                    self.dialog['계좌정보조회'].show()
                except Exception as e:
                    self.dialog['계좌정보조회'] = 화면_계좌정보(sScreenNo=7000, kiwoom=self.kiwoom, parent=self) # self는 메인윈도우, 계좌정보윈도우는 자식윈도우/부모는 메인윈도우
                    self.dialog['계좌정보조회'].KiwoomConnect()
                    self.dialog['계좌정보조회'].show()
            else:
                self.dialog['계좌정보조회'] = 화면_계좌정보(sScreenNo=7000, kiwoom=self.kiwoom, parent=self)
                self.dialog['계좌정보조회'].KiwoomConnect()
                self.dialog['계좌정보조회'].show()
        elif _action == "actionSectorView":
            # self.F_sectorview()
            if self.dialog.get('업종정보조회') is not None:
                try:
                    self.dialog['업종정보조회'].show()
                except Exception as e:
                    self.dialog['업종정보조회'] = 화면_업종정보(sScreenNo=9900, kiwoom=self.kiwoom, parent=self)
                    self.dialog['업종정보조회'].KiwoomConnect()
                    self.dialog['업종정보조회'].show()
            else:
                self.dialog['업종정보조회'] = 화면_업종정보(sScreenNo=9900, kiwoom=self.kiwoom, parent=self)
                self.dialog['업종정보조회'].KiwoomConnect()
                self.dialog['업종정보조회'].show()
        elif _action == "actionSectorPriceView":
            # self.F_sectorpriceview()
            if self.dialog.get('업종별주가조회') is not None:
                try:
                    self.dialog['업종별주가조회'].show()
                except Exception as e:
                    self.dialog['업종별주가조회'] = 화면_업종별주가(sScreenNo=9900, kiwoom=self.kiwoom, parent=self)
                    self.dialog['업종별주가조회'].KiwoomConnect()
                    self.dialog['업종별주가조회'].show()
            else:
                self.dialog['업종별주가조회'] = 화면_업종별주가(sScreenNo=9900, kiwoom=self.kiwoom, parent=self)
                self.dialog['업종별주가조회'].KiwoomConnect()
                self.dialog['업종별주가조회'].show()
        elif _action == "actionTickLogger":
            self.RobotAdd_TickLogger()
            self.RobotView()
        elif _action == "actionTickMonitor":
            self.RobotAdd_TickMonitor()
            self.RobotView()
        elif _action == "actionTickTradeRSI":
            self.RobotAdd_TickTradeRSI()
            self.RobotView()
        elif _action == "actionTradeSuperValue":
            self.RobotAdd_TradeSuperValue()
            self.RobotView()
        elif _action == "actionTradeCondition": # 키움 조건검색식을 이용한 트레이딩
            print("MainWindow : MENU_Action_actionTradeCondition")
            self.RobotAdd_TradeCondition()
            self.RobotView()
        elif _action == "actionRobotLoad":
            self.RobotLoad()
            self.RobotView()
        elif _action == "actionRobotSave":
            self.RobotSave()
        elif _action == "actionRobotOneRun":
            self.RobotOneRun()
            self.RobotView()
        elif _action == "actionRobotOneStop":
            self.RobotOneStop()
            self.RobotView()
        elif _action == "actionRobotMonitoringStop":
            self.RobotMonitoringStop()
            self.RobotView()
        elif _action == "actionRobotRun":
            self.RobotRun()
            self.RobotView()
        elif _action == "actionRobotStop":
            self.RobotStop()
            self.RobotView()
        elif _action == "actionRobotRemove":
            self.RobotRemove()
            self.RobotView()
        elif _action == "actionRobotClear":
            self.RobotClear()
            self.RobotView()
        elif _action == "actionRobotView":
            self.RobotView()
            for r in self.robots:
                logger.debug('%s %s %s %s' % (r.sName, r.UUID, len(r.portfolio), r.GetStatus()))
        elif _action == "actionCodeBuild":
            self.종목코드 = self.StockCodeBuild(to_db=True)
            QMessageBox.about(self, "종목코드 생성", " %s 항목의 종목코드를 생성하였습니다." % (len(self.종목코드.index)))
        elif _action == "actionOpenAPI_document":
            self.kiwoom_doc()
        elif _action == "actionTEST":
            futurecodelist = self.kiwoom.dynamicCall('GetFutureList')
            codes = futurecodelist.split(';')
            print(futurecodelist)

    # -------------------------------------------
    # 키움증권 OpenAPI
    # -------------------------------------------

    # 키움API ActiveX를 메모리에 올림
    def KiwoomAPI(self):
        self.kiwoom = QAxWidget("KHOPENAPI.KHOpenAPICtrl.1")

    # 메모리에 올라온 ActiveX와 On시리즈와 붙임(콜백 : 이벤트가 오면 나를 불러줘)
    def KiwoomConnect(self):
        print("MainWindow : KiwoomConnect")
        self.kiwoom.OnEventConnect[int].connect(self.OnEventConnect) # 키움의 OnEventConnect와 이 프로그램의 OnEventConnect 함수와 연결시킴
        self.kiwoom.OnReceiveMsg[str, str, str, str].connect(self.OnReceiveMsg)
        # self.kiwoom.OnReceiveTrCondition[str, str, str, int, int].connect(self.OnReceiveTrCondition)
        self.kiwoom.OnReceiveTrData[str, str, str, str, str, int, str, str, str].connect(self.OnReceiveTrData)
        self.kiwoom.OnReceiveChejanData[str, int, str].connect(self.OnReceiveChejanData)
        # self.kiwoom.OnReceiveConditionVer[int, str].connect(self.OnReceiveConditionVer)
        # self.kiwoom.OnReceiveRealCondition[str, str, str, str].connect(self.OnReceiveRealCondition)
        self.kiwoom.OnReceiveRealData[str, str, str].connect(self.OnReceiveRealData)

    # ActiveX와 On시리즈 연결 해제
    def KiwoomDisConnect(self):
        self.kiwoom.OnEventConnect[int].disconnect(self.OnEventConnect)
        self.kiwoom.OnReceiveMsg[str, str, str, str].disconnect(self.OnReceiveMsg)
        # self.kiwoom.OnReceiveTrCondition[str, str, str, int, int].disconnect(self.OnReceiveTrCondition)
        # self.kiwoom.OnReceiveTrData[str, str, str, str, str, int, str, str, str].disconnect(self.OnReceiveTrData)
        self.kiwoom.OnReceiveChejanData[str, int, str].disconnect(self.OnReceiveChejanData)
        # self.kiwoom.OnReceiveConditionVer[int, str].disconnect(self.OnReceiveConditionVer)
        # self.kiwoom.OnReceiveRealCondition[str, str, str, str].disconnect(self.OnReceiveRealCondition)
        self.kiwoom.OnReceiveRealData[str, str, str].disconnect(self.OnReceiveRealData)

    # 키움 로그인
    def KiwoomLogin(self):
        print("MainWindow : KiwoomLogin")
        self.kiwoom.dynamicCall("CommConnect()")
        self._login = True
        self.statusbar.showMessage("로그인...")

    # 키움 로그아웃
    def KiwoomLogout(self):
        if self.kiwoom is not None:
            self.kiwoom.dynamicCall("CommTerminate()")

        self.statusbar.showMessage("연결해제됨...")

    # 계좌 보유 종목 받음
    def InquiryList(self, _repeat=0):
        print("MainWindow : InquiryList")
        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, Qstring)', "계좌번호", self.sAccount)
        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, Qstring)', "비밀번호입력매체구분", '00')
        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, Qstring)', "조회구분", '1')
        ret = self.kiwoom.dynamicCall('CommRqData(QString, QString, int, QString)', "계좌평가잔고내역요청", "opw00018",
                                      _repeat, '{:04d}'.format(self.ScreenNumber))

        self.InquiryLoop = QEventLoop()  # 로봇에서 바로 쓸 수 있도록하기 위해서 계좌 조회해서 종목을 받고나서 루프해제시킴
        self.InquiryLoop.exec_()

    # 계좌 번호 / D+2 예수금 받음
    def KiwoomAccount(self):
        print("MainWindow : KiwoomAccount")
        ACCOUNT_CNT = self.kiwoom.dynamicCall('GetLoginInfo("ACCOUNT_CNT")')
        ACC_NO = self.kiwoom.dynamicCall('GetLoginInfo("ACCNO")')
        self.account = ACC_NO.split(';')[0:-1]
        self.sAccount = self.account[0]
        print('계좌 : ', self.sAccount)
        self.kiwoom.dynamicCall('SetInputValue(Qstring, Qstring)', "계좌번호", self.sAccount)
        self.kiwoom.dynamicCall('CommRqData(QString, QString, int, QString)', "d+2예수금요청", "opw00001", 0,
                                '{:04d}'.format(self.ScreenNumber))
        self.depositLoop = QEventLoop()  # self.d2_deposit를 로봇에서 바로 쓸 수 있도록하기 위해서 예수금을 받고나서 루프해제시킴
        self.depositLoop.exec_()

        # return (ACCOUNT_CNT, ACC_NO)

    def KiwoomSendOrder(self, sRQName, sScreenNo, sAccNo, nOrderType, sCode, nQty, nPrice, sHogaGb, sOrgOrderNo):
        if self.주문제한 < 초당횟수제한:
            Order = self.kiwoom.dynamicCall('SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)',
                [sRQName, sScreenNo, sAccNo, nOrderType, sCode, nQty, nPrice, sHogaGb, sOrgOrderNo])
            self.주문제한 += 1
            return (True, Order)
        else:
            return (False, 0)

            # -거래구분값 확인(2자리)
            #
            # 00 : 지정가
            # 03 : 시장가
            # 05 : 조건부지정가
            # 06 : 최유리지정가
            # 07 : 최우선지정가
            # 10 : 지정가IOC
            # 13 : 시장가IOC
            # 16 : 최유리IOC
            # 20 : 지정가FOK
            # 23 : 시장가FOK
            # 26 : 최유리FOK
            # 61 : 장전 시간외단일가매매
            # 81 : 장후 시간외종가
            # 62 : 시간외단일가매매
            #
            # -매매구분값 (1 자리)
            # 1 : 신규매수
            # 2 : 신규매도
            # 3 : 매수취소
            # 4 : 매도취소
            # 5 : 매수정정
            # 6 : 매도정정

    def KiwoomSetRealReg(self, sScreenNo, sCode, sRealType='0'):
        ret = self.kiwoom.dynamicCall('SetRealReg(QString, QString, QString, QString)', sScreenNo, sCode, '9001;10', sRealType) # 10은 실시간FID로 메뉴얼에 나옴(현재가,체결가, 실시간종가)
        return ret
        # pass

    def KiwoomSetRealRemove(self, sScreenNo, sCode):
        ret = self.kiwoom.dynamicCall('SetRealRemove(QString, QString)', sScreenNo, sCode)
        return ret

    def KiwoomScreenNumber(self):
        self.screen_number += 1
        if self.screen_number > 8999:
            self.screen_number = 5000
        return self.screen_number

    def OnEventConnect(self, nErrCode):
        # logger.debug('main:OnEventConnect', nErrCode)

        if nErrCode == 0:
            # self.kiwoom.dynamicCall("KOA_Functions(QString, QString)", ["ShowAccountWindow", ""]) # 계좌 비밀번호 등록 창 실행(자동화를 위해서 AUTO 설정 후 등록 창 미실행
            self.statusbar.showMessage("로그인 성공")
            Telegram("[XTrader]키움API 로그인 성공")
            로그인상태 = True
            # 로그인 성공하고 바로 계좌 및 보유 주식 목록 저장
            self.KiwoomAccount()
            self.InquiryList()
            # self.GetCondition() # 조건검색식을 모두 읽어서 해당하는 종목 저장
        else:
            self.statusbar.showMessage("연결실패... %s" % nErrCode)
            로그인상태 = False

    def OnReceiveMsg(self, sScrNo, sRQName, sTrCode, sMsg):
        # logger.debug('main:OnReceiveMsg [%s] [%s] [%s] [%s]' % (sScrNo, sRQName, sTrCode, sMsg))
        pass

    def OnReceiveTrData(self, sScrNo, sRQName, sTRCode, sRecordName, sPreNext, nDataLength, sErrorCode, sMessage,sSPlmMsg):
        # logger.debug('main:OnReceiveTrData [%s] [%s] [%s] [%s] [%s] [%s] [%s] [%s] [%s] ' % (sScrNo, sRQName, sTRCode, sRecordName, sPreNext, nDataLength, sErrorCode, sMessage, sSPlmMsg))
        print("MainWindow : OnReceiveTrData")
        if self.ScreenNumber != int(sScrNo):
            return

        if sRQName == "주식일봉차트조회":

            self.주식일봉컬럼 = ['일자', '현재가', '거래량', '시가', '고가', '저가', '거래대금']

            cnt = self.kiwoom.dynamicCall('GetRepeatCnt(QString, QString)', sTRCode, sRQName)
            for i in range(0, cnt):
                row = []
                for j in self.주식일봉컬럼:
                    S = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "",
                                                sRQName, i, j).strip().lstrip('0')
                    if len(S) > 0 and S[0] == '-':
                        S = '-' + S[1:].lstrip('0')
                    row.append(S)
                self.종목일봉.append(row)
            if sPreNext == '2' and False:  # 과거 모든데이타 백업시 True로 변경할것
                QTimer.singleShot(주문지연, lambda: self.ReguestPriceDaily(_repeat=2))
            else:
                df = DataFrame(data=self.종목일봉, columns=self.주식일봉컬럼)
                df['일자'] = df['일자'].apply(lambda x: x[0:4] + '-' + x[4:6] + '-' + x[6:])
                df['종목코드'] = self.종목코드[0]
                df = df[['종목코드', '일자', '현재가', '시가', '고가', '저가', '거래량', '거래대금']]
                values = list(df.values)

                try:
                    df.ix[df.현재가 == '', ['현재가']] = 0
                except Exception as e:
                    pass
                try:
                    df.ix[df.시가 == '', ['시가']] = 0
                except Exception as e:
                    pass
                try:
                    df.ix[df.고가 == '', ['고가']] = 0
                except Exception as e:
                    pass
                try:
                    df.ix[df.저가 == '', ['저가']] = 0
                except Exception as e:
                    pass
                try:
                    df.ix[df.거래량 == '', ['거래량']] = 0
                except Exception as e:
                    pass
                try:
                    df.ix[df.거래대금 == '', ['거래대금']] = 0
                except Exception as e:
                    pass
                try:
                    df.ix[df.거래대금 == '-', ['거래대금']] = 0
                except Exception as e:
                    pass

                conn = mysqlconn()

                cursor = conn.cursor()
                cursor.executemany(
                    "replace into 일별주가(종목코드,일자,종가,시가,고가,저가,거래량,거래대금) values( %s, %s, %s, %s, %s, %s, %s, %s )",
                    df.values.tolist())

                conn.commit()
                conn.close()

                self.백업한종목수 += 1
                if len(self.백업할종목코드) > 0:
                    self.종목코드 = self.백업할종목코드.pop(0)
                    self.종목일봉 = []

                    QTimer.singleShot(주문지연, lambda: self.ReguestPriceDaily(_repeat=0))
                else:
                    QTimer.singleShot(주문지연, lambda: self.Backup(작업="주식일봉백업"))

        if sRQName == "종목별투자자조회":
            self.종목별투자자컬럼 = ['일자', '현재가', '전일대비', '누적거래대금', '개인투자자', '외국인투자자', '기관계', '금융투자', '보험', '투신', '기타금융', '은행',
                             '연기금등', '국가', '내외국인', '사모펀드', '기타법인']

            cnt = self.kiwoom.dynamicCall('GetRepeatCnt(QString, QString)', sTRCode, sRQName)
            for i in range(0, cnt):
                row = []
                for j in self.종목별투자자컬럼:
                    S = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "",
                                                sRQName, i, j).strip().lstrip('0').replace('--', '-')
                    row.append(S)
                self.종목별투자자.append(row)
            if sPreNext == '2' and False:
                QTimer.singleShot(주문지연, lambda: self.RequestInvestorDaily(_repeat=2))
            else:
                if len(self.종목별투자자) > 0:
                    df = DataFrame(data=self.종목별투자자, columns=self.종목별투자자컬럼)
                    # df['일자'] = pd.to_datetime(df['일자'], format='%Y%m%d')
                    df['일자'] = df['일자'].apply(lambda x: x[0:4] + '-' + x[4:6] + '-' + x[6:])
                    # df['현재가'] = np.abs(df['현재가'].convert_objects(convert_numeric=True))
                    df['현재가'] = np.abs(pd.to_numeric(df['현재가'], errors='coerce'))
                    df['종목코드'] = self.종목코드[0]
                    df = df[['종목코드'] + self.종목별투자자컬럼]
                    # values = list(df.values)

                    try:
                        df.ix[df.현재가 == '', ['현재가']] = 0
                    except Exception as e:
                        pass
                    try:
                        df.ix[df.전일대비 == '', ['전일대비']] = 0
                    except Exception as e:
                        pass
                    try:
                        df.ix[df.누적거래대금 == '', ['누적거래대금']] = 0
                    except Exception as e:
                        pass
                    try:
                        df.ix[df.개인투자자 == '', ['개인투자자']] = 0
                    except Exception as e:
                        pass
                    try:
                        df.ix[df.외국인투자자 == '', ['외국인투자자']] = 0
                    except Exception as e:
                        pass
                    try:
                        df.ix[df.기관계 == '', ['기관계']] = 0
                    except Exception as e:
                        pass
                    try:
                        df.ix[df.금융투자 == '', ['금융투자']] = 0
                    except Exception as e:
                        pass
                    try:
                        df.ix[df.금융투자 == '', ['금융투자']] = 0
                    except Exception as e:
                        pass
                    try:
                        df.ix[df.보험 == '', ['보험']] = 0
                    except Exception as e:
                        pass
                    try:
                        df.ix[df.투신 == '', ['투신']] = 0
                    except Exception as e:
                        pass
                    try:
                        df.ix[df.기타금융 == '', ['기타금융']] = 0
                    except Exception as e:
                        pass
                    try:
                        df.ix[df.은행 == '', ['은행']] = 0
                    except Exception as e:
                        pass
                    try:
                        df.ix[df.연기금등 == '', ['연기금등']] = 0
                    except Exception as e:
                        pass
                    try:
                        df.ix[df.국가 == '', ['국가']] = 0
                    except Exception as e:
                        pass
                    try:
                        df.ix[df.내외국인 == '', ['내외국인']] = 0
                    except Exception as e:
                        pass
                    try:
                        df.ix[df.사모펀드 == '', ['사모펀드']] = 0
                    except Exception as e:
                        pass
                    try:
                        df.ix[df.기타법인 == '', ['기타법인']] = 0
                    except Exception as e:
                        pass

                    df.dropna(inplace=True)

                    conn = mysqlconn()

                    cursor = conn.cursor()
                    cursor.executemany(
                        "replace into 종목별투자자(종목코드,일자,종가,전일대비,누적거래대금,개인투자자,외국인투자자,기관계,금융투자,보험,투신,기타금융,은행,연기금등,국가,내외국인,사모펀드,기타법인) values(%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)",
                        df.values.tolist())
                    conn.commit()
                    conn.close()

                else:
                    logger.info("%s 데이타없음", self.종목코드)

                self.백업한종목수 += 1
                if len(self.백업할종목코드) > 0:
                    self.종목코드 = self.백업할종목코드.pop(0)
                    self.종목별투자자 = []

                    QTimer.singleShot(주문지연, lambda: self.RequestInvestorDaily(_repeat=0))
                else:
                    QTimer.singleShot(주문지연, lambda: self.Backup(작업="종목별투자자백업"))

        if sRQName == "주식분봉차트조회":
            self.주식분봉컬럼 = ['체결시간', '현재가', '시가', '고가', '저가', '거래량']

            cnt = self.kiwoom.dynamicCall('GetRepeatCnt(QString, QString)', sTRCode, sRQName)
            for i in range(0, cnt):
                row = []
                for j in self.주식분봉컬럼:
                    S = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "",
                                                sRQName, i, j).strip().lstrip('0')
                    if len(S) > 0 and (S[0] == '-' or S[0] == '+'):
                        S = S[1:].lstrip('0')
                    row.append(S)
                self.종목분봉.append(row)
            if sPreNext == '2' and False:
                QTimer.singleShot(주문지연, lambda: self.ReguestPriceMin(_repeat=2))
            else:
                df = DataFrame(data=self.종목분봉, columns=self.주식분봉컬럼)
                df['체결시간'] = df['체결시간'].apply(
                    lambda x: x[0:4] + '-' + x[4:6] + '-' + x[6:8] + ' ' + x[8:10] + ':' + x[10:12] + ':' + x[12:])
                df['종목코드'] = self.종목코드[0]
                df['틱범위'] = self.틱범위
                df = df[['종목코드', '틱범위', '체결시간', '현재가', '시가', '고가', '저가', '거래량']]
                values = list(df.values)

                try:
                    df.ix[df.현재가 == '', ['현재가']] = 0
                except Exception as e:
                    pass
                try:
                    df.ix[df.시가 == '', ['시가']] = 0
                except Exception as e:
                    pass
                try:
                    df.ix[df.고가 == '', ['고가']] = 0
                except Exception as e:
                    pass
                try:
                    df.ix[df.저가 == '', ['저가']] = 0
                except Exception as e:
                    pass
                try:
                    df.ix[df.거래량 == '', ['거래량']] = 0
                except Exception as e:
                    pass

                conn = mysqlconn()

                cursor = conn.cursor()
                cursor.executemany(
                    "replace into 분별주가(종목코드,틱범위,체결시간,종가,시가,고가,저가,거래량) values( %s, %s, %s, %s, %s, %s, %s, %s )",
                    df.values.tolist())

                conn.commit()
                conn.close()

                self.백업한종목수 += 1
                if len(self.백업할종목코드) > 0:
                    self.종목코드 = self.백업할종목코드.pop(0)
                    self.종목분봉 = []

                    QTimer.singleShot(주문지연, lambda: self.ReguestPriceMin(_repeat=0))
                else:
                    QTimer.singleShot(주문지연, lambda: self.Backup(작업="주식분봉백업"))

        if sRQName == "d+2예수금요청":
            data = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)',sTRCode, "", sRQName, 0, "d+2추정예수금")

            # 입력된 문자열에 대해 lstrip 메서드를 통해 문자열 왼쪽에 존재하는 '-' 또는 '0'을 제거. 그리고 format 함수를 통해 천의 자리마다 콤마를 추가한 문자열로 변경
            strip_data = data.lstrip('-0')
            if strip_data == '':
                strip_data = '0'

            format_data = format(int(strip_data), ',d')
            if data.startswith('-'):
                format_data = '-' + format_data

            global d2deposit # D+2 예수금 → 매수 가능 금액 계산을 위함
            d2deposit = format_data
            print("예수금 %s 저장 완료" % (d2deposit))
            self.depositLoop.exit() # self.d2_deposit를 로봇에서 바로 쓸 수 있도록하기 위해서 예수금을 받고나서 루프해제시킴

        if sRQName == "계좌평가잔고내역요청":
            try:
                cnt = self.kiwoom.dynamicCall('GetRepeatCnt(QString, QString)', sTRCode, sRQName)

                global df_keeplist # 계좌 보유 종목 리스트

                result = []

                cols = ['종목번호', '종목명', '보유수량', '매입가', '매입금액'] #, '평가금액', '수익률(%)', '평가손익', '매매가능수량']
                for i in range(0, cnt):
                    row = []
                    for j in cols:
                      # S = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, i, '종목번호').strip().lstrip('0')
                        S = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, i, j).strip().lstrip('0')

                        if len(S) > 0 and S[0] == '-':
                            S = '-' + S[1:].lstrip('0')

                        if j == '종목번호':
                            S = S.replace('A', '')  # 종목코드 맨 첫 'A'를 삭제하기 위함

                        row.append(S)

                    result.append(row)

                    # logger.debug("%s" % row)

                if sPreNext == '2':
                    self.remained_data = True
                    self.InquiryList(_repeat=2)
                else:
                    self.remained_data = False
                    df_keeplist = DataFrame(data=result, columns=cols)

                print('계좌평가잔고내역', df_keeplist)
                self.InquiryLoop.exit()

                # if sRQName == "계좌평가잔고내역요청":
                #     cnt = self.kiwoom.dynamicCall('GetRepeatCnt(QString, QString)', sTRCode, sRQName)
                #
                #     for i in range(0, cnt):
                #         row = []
                #         for j in self.columns:
                #             # print(j)
                #             S = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode,
                #                                         "", sRQName, i, j).strip().lstrip('0')
                #             # print(S)
                #             if len(S) > 0 and S[0] == '-':
                #                 S = '-' + S[1:].lstrip('0')
                #             row.append(S)
                #         self.result.append(row)
                #         # logger.debug("%s" % row)
                #     if sPreNext == '2':
                #         self.Request(_repeat=2)
                #     else:
                #         self.model.update(DataFrame(data=self.result, columns=self.보이는컬럼))
                #         print(self.result)
                #         for i in range(len(self.columns)):
                #             self.tableView.resizeColumnToContents(i)
            except Exception as e:
                print(e)

    def OnReceiveChejanData(self, sGubun, nItemCnt, sFidList):
        # logger.debug('main:OnReceiveChejanData [%s] [%s] [%s]' % (sGubun, nItemCnt, sFidList))
        pass

        # sFid별 주요데이터는 다음과 같습니다.
        # "9201" : "계좌번호"
        # "9203" : "주문번호"
        # "9001" : "종목코드"
        # "913" : "주문상태"
        # "302" : "종목명"
        # "900" : "주문수량"
        # "901" : "주문가격"
        # "902" : "미체결수량"
        # "903" : "체결누계금액"
        # "904" : "원주문번호"
        # "905" : "주문구분"
        # "906" : "매매구분"
        # "907" : "매도수구분"
        # "908" : "주문/체결시간"
        # "909" : "체결번호"
        # "910" : "체결가"
        # "911" : "체결량"
        # "10" : "현재가"
        # "27" : "(최우선)매도호가"
        # "28" : "(최우선)매수호가"
        # "914" : "단위체결가"
        # "915" : "단위체결량"
        # "919" : "거부사유"
        # "920" : "화면번호"
        # "917" : "신용구분"
        # "916" : "대출일"
        # "930" : "보유수량"
        # "931" : "매입단가"
        # "932" : "총매입가"
        # "933" : "주문가능수량"
        # "945" : "당일순매수수량"
        # "946" : "매도/매수구분"
        # "950" : "당일총매도손일"
        # "951" : "예수금"
        # "307" : "기준가"
        # "8019" : "손익율"
        # "957" : "신용금액"
        # "958" : "신용이자"
        # "918" : "만기일"
        # "990" : "당일실현손익(유가)"
        # "991" : "당일실현손익률(유가)"
        # "992" : "당일실현손익(신용)"
        # "993" : "당일실현손익률(신용)"
        # "397" : "파생상품거래단위"
        # "305" : "상한가"
        # "306" : "하한가"

    def OnReceiveRealData(self, sRealKey, sRealType, sRealData):
        # logger.debug('main:OnReceiveRealData [%s] [%s] [%s]' % (sRealKey, sRealType, sRealData))
        pass

    """
    def OnReceiveTrCondition(self, sScrNo, strCodeList, strConditionName, nIndex, nNext):
        logger.debug('main:OnReceiveTrCondition [%s] [%s] [%s] [%s] [%s]' % (
        sScrNo, strCodeList, strConditionName, nIndex, nNext))
        print("MainWindow : OnReceiveTrCondition")

        # (1회성, 실시간) 종목 조건검색 요청시 발생되는 이벤트
        # :param screenNo: string
        # :param codes: string - 종목코드 목록(각 종목은 세미콜론으로 구분됨)
        # :param conditionName: string - 조건식 이름
        # :param conditionIndex: int - 조건식 인덱스
        # :param inquiry: int - 조회구분(0: 남은데이터 없음, 2: 남은데이터 있음)


        current = datetime.datetime.now()
        time = str(current)[:-7] #str(current.hour) +str(current.minute) + str(current.second)

        cindexs = []  # 조건식 컨디션 인덱스
        cnames = []  # 조건식 컨디션 이름
        codes = []
        codenames = []
        price = []
        times = []
        self.df_condition = DataFrame()
        try:
            if strCodeList == "":
                return
            # print(nIndex, strConditionName)
            self.codeList = strCodeList.split(';')
            del self.codeList[-1]

            # print("종목개수: ", len(self.codeList))
            # print(self.codeList)

            for code in self.codeList:
                times.append(time)
                cindexs.append(nIndex)
                cnames.append(strConditionName)
                codes.append(code)
                codenames.append(self.kiwoom.dynamicCall("GetMasterCodeName(QString)", code))
                # try:
                #     # ret = self.KiwoomSetRealReg(self.ScreenNumber, code)
                #     price.append(self.realprice)
                #
                # except Exception as e:
                #     price.append(1)
                price.append(1)

            self.df_condition['시간'] = times
            self.df_condition['인덱스'] = cindexs
            self.df_condition['조건명'] = cnames
            self.df_condition['종목코드'] = codes
            self.df_condition['종목명'] = codenames
            self.df_condition['현재가'] = price

            with sqlite3.connect(DATABASE) as conn:
                cursor = conn.cursor()
                query = "insert or replace into 조건검색식분석( 시간, 인덱스, 조건명, 종목코드, 종목명, 현재가) values(?, ?, ?, ?, ?, ?)"
                cursor.executemany(query, self.df_condition.values.tolist())
                conn.commit()

        except Exception as e:
            print("OnReceiveTrCondition_Error")
            print(e)

        finally:
            with sqlite3.connect(DATABASE) as conn:
                query = "select * from 조건검색식분석"
                df = pdsql.read_sql_query(query, con=conn)
                df.to_csv("조건검색분석.csv")
            # Telegram(time + " : " + "조건검색식 종목 저장완료")
            self.conditionLoop.exit()

    def OnReceiveConditionVer(self, lRet, sMsg):
        # logger.debug('main:OnReceiveConditionVer : [이벤트] 조건식 저장',lRet, sMsg)  머니봇의 오류 코드를 아래와 같이 수정하여 logger Error 방지함(18.06.13)
        print("MainWindow : OnReceiveConditionVer")
        logger.debug('main:OnReceiveConditionVer : [이벤트] 조건식 저장 [%s] [%s]' % (lRet, sMsg))

        # getConditionLoad() 메서드의 조건식 목록 요청에 대한 응답 이벤트
        # :param receive: int - 응답결과(1: 성공, 나머지 실패)
        # :param msg: string - 메세지

        try:
            self.condition = self.getConditionNameList() 
            # print("조건식 개수: ", len(self.condition))

            # for key in self.condition.keys():
            # print("조건식: ", key, ": ", self.condition[key])

        except Exception as e:
            print("OnReceiveConditionVer_Error")

        finally:
            self.conditionLoop.exit()

    def OnReceiveRealCondition(self, sTrCode, strType, strConditionName, strConditionIndex):
        logger.debug(
            'main:OnReceiveRealCondition [%s] [%s] [%s] [%s]' % (sTrCode, strType, strConditionName, strConditionIndex))
        print("MainWindow : OnReceiveRealCondition")
    """

    # ------------------------------------------------------------
    # robot 함수
    # ------------------------------------------------------------
    def GetUnAssignedScreenNumber(self):
        스크린번호 = 0
        사용중인스크린번호 = []
        for r in self.robots:
            사용중인스크린번호.append(r.sScreenNo)

        for i in range(로봇스크린번호시작, 로봇스크린번호종료 + 1):
            if i not in 사용중인스크린번호:
                스크린번호 = i
                break
        return 스크린번호

    def RobotRemove(self):
        RobotUUID = \
        self.model._data[self.tableView_robot_current_index.row():self.tableView_robot_current_index.row() + 1][
            'RobotID'].values[0]

        robot_found = None
        for r in self.robots:
            if r.UUID == RobotUUID:
                robot_found = r
                break

        if robot_found == None:
            return

        reply = QMessageBox.question(self,
                                     "로봇 삭제", "로봇을 삭제할까요?\n%s" % robot_found.GetStatus()[0:4],
                                     QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel)
        if reply == QMessageBox.Cancel:
            pass
        elif reply == QMessageBox.No:
            pass
        elif reply == QMessageBox.Yes:
            self.robots.remove(robot_found)

    def RobotClear(self):
        reply = QMessageBox.question(self,
                                     "로봇 전체 삭제", "로봇 전체를 삭제할까요?",
                                     QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel)
        if reply == QMessageBox.Cancel:
            pass
        elif reply == QMessageBox.No:
            pass
        elif reply == QMessageBox.Yes:
            self.robots = []

    def RobotRun(self):
        for r in self.robots:
            r.초기조건()
            # logger.debug('%s %s %s %s' % (r.sName, r.UUID, len(r.portfolio), r.GetStatus()))
            r.Run(flag=True, sAccount=로봇거래계좌번호)

        self.statusbar.showMessage("RUN !!!")

    def RobotStop(self):
        # reply = QMessageBox.question(self,
        #                              "전체 로봇 실행 중지", "전체 로봇 실행을 중지할까요?",
        #                              QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel)
        # if reply == QMessageBox.Cancel:
        #     pass
        # elif reply == QMessageBox.No:
        #     pass
        # elif reply == QMessageBox.Yes:
        #     for r in self.robots:
        #         r.Run(flag=False)
        #
        #     self.RobotSaveSilently()

        for r in self.robots:
            r.Run(flag=False)
            print('RobotStop')
            for code in list(r.portfolio.keys()):
                print('RobotStop_code : ', code)
                if r.portfolio[code].수량 == 0:
                    print('RobotStop_pop : ', code)
                    r.portfolio.pop(code)
            print('RobotStop_result : ', r.portfolio)
        self.RobotView()
        self.RobotSaveSilently()
        Telegram("[XTrader]전체 ROBOT 실행 중지시킵니다.")

        self.statusbar.showMessage("STOP !!!")

    def RobotOneRun(self):
        try:
            RobotUUID = \
            self.model._data[self.tableView_robot_current_index.row():self.tableView_robot_current_index.row() + 1][
                'RobotID'].values[0]
        except Exception as e:
            RobotUUID = ''

        robot_found = None
        for r in self.robots:
            if r.UUID == RobotUUID:
                robot_found = r
                break

        if robot_found == None:
            return

        robot_found.Run(flag=True, sAccount=로봇거래계좌번호)

    def RobotOneStop(self):
        try:
            RobotUUID = \
            self.model._data[self.tableView_robot_current_index.row():self.tableView_robot_current_index.row() + 1][
                'RobotID'].values[0]
        except Exception as e:
            RobotUUID = ''

        robot_found = None
        for r in self.robots:
            if r.UUID == RobotUUID:
                robot_found = r
                break

        if robot_found == None:
            return

        reply = QMessageBox.question(self,
                                     "로봇 실행 중지", "로봇 실행을 중지할까요?\n%s" % robot_found.GetStatus(),
                                     QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel)
        if reply == QMessageBox.Cancel:
            pass
        elif reply == QMessageBox.No:
            pass
        elif reply == QMessageBox.Yes:
            robot_found.Run(flag=False)
            for code in list(robot_found.portfolio.keys()):
                    if robot_found.portfolio[code].수량 == 0:
                        robot_found.portfolio.pop(code)
            self.RobotView()
            self.RobotSaveSilently()

    def RobotMonitoringStop(self):
        print('RobotMonitoringStop')
        for r in self.robots:
            print(r.매수할종목)
            r.매수할종목 = []

    def RobotLoad(self):
        reply = QMessageBox.question(self,
                                     "로봇 탑제", "저장된 로봇을 읽어올까요?",
                                     QMessageBox.Yes | QMessageBox.Cancel)
        if reply == QMessageBox.Cancel:
            pass

        elif reply == QMessageBox.Yes:
            with sqlite3.connect(DATABASE) as conn:
                cursor = conn.cursor()

                cursor.execute('select uuid, strategy, name, robot from Robots')

                self.robots = []
                for row in cursor.fetchall():
                    uuid, strategy, name, robot_encoded = row
                    robot = base64.decodebytes(robot_encoded)
                    r = pickle.loads(robot)

                    r.kiwoom = self.kiwoom
                    r.parent = self
                    r.d = today

                    r.running = False
                    print(r.portfolio)
                    # logger.debug(r.sName, r.UUID, len(r.portfolio))
                    self.robots.append(r)

            self.RobotView()

    def RobotSave(self):
        print("MainWindow : RobotSave")

        if len(self.robots)>0:
            reply = QMessageBox.question(self,
                                         "로봇 저장", "현재 로봇을 저장할까요?",
                                         QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel)
            if reply == QMessageBox.Cancel:
                pass
            elif reply == QMessageBox.No:
                pass
            elif reply == QMessageBox.Yes:
                self.RobotSaveSilently()
        else:
            QMessageBox.about(self, "Robot Save Error", "현재 설정된 로봇이 없습니다.")

    def RobotSaveSilently(self):
        # sqlite3 사용
        try:
            with sqlite3.connect(DATABASE) as conn:
                cursor = conn.cursor()
                cursor.execute('delete from Robots')
                conn.commit()

                for r in self.robots:
                    r.kiwoom = None
                    r.parent = None

                    for code in list(r.portfolio.keys()):
                        print('RobotSaveSilently_code', code)
                        if r.portfolio[code].수량 == 0:
                            r.portfolio.pop(code)
                    print('포트폴리오 : ', r.portfolio)
                    uuid = r.UUID
                    strategy = r.__class__.__name__
                    name = r.sName

                    robot = pickle.dumps(r, protocol=pickle.HIGHEST_PROTOCOL, fix_imports=True)

                    # robot_encoded = base64.encodestring(robot)
                    robot_encoded = base64.encodebytes(robot)

                    cursor.execute("insert or replace into Robots(uuid, strategy, name, robot) values (?, ?, ?, ?)",
                                   [uuid, strategy, name, robot_encoded])
                    conn.commit()
        except Exception as e:
            print('RobotSaveSilently', e)
        finally:
            r.kiwoom = self.kiwoom
            r.parent = self
            self.statusbar.showMessage("로봇 저장 완료")

    def RobotView(self):
        print("MainWindow : RobotView")
        result = []
        for r in self.robots:
            logger.debug('%s %s %s %s' % (r.sName, r.UUID, len(r.portfolio), r.GetStatus()))
            result.append(r.GetStatus())

        self.model.update(DataFrame(data=result, columns=self.robot_columns))

        # RobotID 숨김
        self.tableView_robot.setColumnHidden(2, True)

        for i in range(len(self.robot_columns)):
            self.tableView_robot.resizeColumnToContents(i)

            # self.tableView_robot.horizontalHeader().setStretchLastSection(True)

    def RobotEdit(self, QModelIndex):
        try:
            # print(self.model._data[QModelIndex.row()])
            Robot타입 = self.model._data[QModelIndex.row():QModelIndex.row() + 1]['Robot타입'].values[0]
            RobotUUID = self.model._data[QModelIndex.row():QModelIndex.row() + 1]['RobotID'].values[0]
            # print(Robot타입, RobotUUID)

            robot_found = None
            for r in self.robots:
                if r.UUID == RobotUUID:
                    robot_found = r
                    break

            if robot_found == None:
                return

            if Robot타입 == 'CTickLogger':
                self.RobotEdit_TickLogger(robot_found)
            elif Robot타입 == 'CTickMonitor':
                self.RobotEdit_TickMonitor(robot_found)
            elif Robot타입 == 'CTickTradeRSI':
                self.RobotEdit_TickTradeRSI(robot_found)
            elif Robot타입 == 'CTickFuturesLogger':
                self.RobotEdit_TickFuturesLogger(robot_found)
            elif Robot타입 == 'CTradeSuperValue':
                self.RobotEdit_TradeSuperValue(robot_found)
            elif Robot타입 == 'CTradeCondition':
                self.RobotEdit_TradeCondition(robot_found)
        except Exception as e:
            print('RobotEdit', e)

    def robot_selected(self, QModelIndex):
        # print(self.model._data[QModelIndex.row()])
        try:
            Robot타입 = self.model._data[QModelIndex.row():QModelIndex.row() + 1]['Robot타입'].values[0]

            uuid = self.model._data[QModelIndex.row():QModelIndex.row() + 1]['RobotID'].values[0]
            portfolio = None
            for r in self.robots:
                if r.UUID == uuid:
                    portfolio = r.portfolio
                    model = PandasModel()
                    result = []
                    for p, v in portfolio.items():
                        result.append((v.종목코드, v.종목명.strip(), v.매수가, v.수량, v.매수일))
                    self.portfolio_model.update((DataFrame(data=result, columns=['종목코드', '종목명', '매수가', '수량', '매수일'])))

                    break
        except Exception as e:
            print('robot_selected', e)

    def robot_double_clicked(self, QModelIndex):
        self.RobotEdit(QModelIndex)
        self.RobotView()

    def RobotCurrentIndex(self, index):
        self.tableView_robot_current_index = index

    def RobotAdd_TickLogger(self):
        스크린번호 = self.GetUnAssignedScreenNumber()
        R = 화면_TickLogger(parent=self)
        R.lineEdit_screen_number.setText('{:04d}'.format(스크린번호))
        if R.exec_():
            이름 = R.lineEdit_name.text()
            스크린번호 = int(R.lineEdit_screen_number.text())
            종목유니버스 = R.plainTextEdit_base_price.toPlainText()
            종목유니버스리스트 = [x.strip() for x in 종목유니버스.split(',')]

            self.KiwoomAccount()
            r = CTickLogger(sName=이름, UUID=uuid.uuid4().hex, kiwoom=self.kiwoom, parent=self)
            r.Setting(sScreenNo=스크린번호, 종목유니버스=종목유니버스리스트)

            self.robots.append(r)

    def RobotEdit_TickLogger(self, robot):
        R = 화면_TickLogger(parent=self)
        R.lineEdit_name.setText(robot.sName)
        R.lineEdit_screen_number.setText('{:04d}'.format(robot.sScreenNo))
        R.plainTextEdit_base_price.setPlainText(','.join([str(x) for x in robot.종목유니버스]))

        if R.exec_():
            이름 = R.lineEdit_name.text()
            스크린번호 = int(R.lineEdit_screen_number.text())
            종목유니버스 = R.plainTextEdit_base_price.toPlainText()
            종목유니버스리스트 = [x.strip() for x in 종목유니버스.split(',')]

            robot.sName = 이름
            robot.Setting(sScreenNo=스크린번호, 종목유니버스=종목유니버스리스트)

    def RobotAdd_TickMonitor(self):
        스크린번호 = self.GetUnAssignedScreenNumber()
        R = 화면_TickLogger(parent=self)
        R.lineEdit_screen_number.setText('{:04d}'.format(스크린번호))
        if R.exec_():
            이름 = R.lineEdit_name.text()
            스크린번호 = int(R.lineEdit_screen_number.text())
            종목유니버스 = R.plainTextEdit_base_price.toPlainText()
            종목유니버스리스트 = [x.strip() for x in 종목유니버스.split(',')]

            self.KiwoomAccount()
            r = CTickMonitor(sName=이름, UUID=uuid.uuid4().hex, kiwoom=self.kiwoom, parent=self)
            r.Setting(sScreenNo=스크린번호, 종목유니버스=종목유니버스리스트)

            self.robots.append(r)

    def RobotEdit_TickMonitor(self, robot):
        R = 화면_TickLogger(parent=self)
        R.lineEdit_name.setText(robot.sName)
        R.lineEdit_screen_number.setText('{:04d}'.format(robot.sScreenNo))
        R.plainTextEdit_base_price.setPlainText(','.join([str(x) for x in robot.종목유니버스]))

        if R.exec_():
            이름 = R.lineEdit_name.text()
            스크린번호 = int(R.lineEdit_screen_number.text())
            종목유니버스 = R.plainTextEdit_base_price.toPlainText()
            종목유니버스리스트 = [x.strip() for x in 종목유니버스.split(',')]

            robot.sName = 이름
            robot.Setting(sScreenNo=스크린번호, 종목유니버스=종목유니버스리스트)

    def RobotAdd_TickTradeRSI(self):
        스크린번호 = self.GetUnAssignedScreenNumber()
        R = 화면_TickTradeRSI(parent=self)
        R.lineEdit_screen_number.setText('{:04d}'.format(스크린번호))

        if R.exec_():

            이름 = R.lineEdit_name.text()
            스크린번호 = int(R.lineEdit_screen_number.text())
            단위투자금 = int(R.lineEdit_unit.text()) * 10000
            매수방법 = R.comboBox_buy_sHogaGb.currentText().strip()[0:2]
            매도방법 = R.comboBox_sell_sHogaGb.currentText().strip()[0:2]
            시총상한 = int(R.lineEdit_max.text().strip())
            시총하한 = int(R.lineEdit_min.text().strip())
            포트폴리오수 = int(R.lineEdit_portsize.text().strip())

            r = CTickTradeRSI(sName=이름, UUID=uuid.uuid4().hex, kiwoom=self.kiwoom, parent=self)
            r.Setting(sScreenNo=스크린번호, 단위투자금=단위투자금, 시총상한=시총상한, 시총하한=시총하한, 포트폴리오수=포트폴리오수, 매수방법=매수방법, 매도방법=매도방법)

            self.robots.append(r)

    def RobotEdit_TickTradeRSI(self, robot):
        R = 화면_TickTradeRSI(parent=self)
        R.lineEdit_name.setText(robot.sName)
        R.lineEdit_screen_number.setText('{:04d}'.format(robot.sScreenNo))
        R.lineEdit_unit.setText(str(robot.단위투자금 // 10000))
        R.lineEdit_portsize.setText(str(robot.포트폴리오수))
        R.lineEdit_max.setText(str(robot.시총상한))
        R.lineEdit_min.setText(str(robot.시총하한))
        R.comboBox_buy_sHogaGb.setCurrentIndex(R.comboBox_buy_sHogaGb.findText(robot.매수방법, flags=Qt.MatchContains))
        R.comboBox_sell_sHogaGb.setCurrentIndex(R.comboBox_sell_sHogaGb.findText(robot.매도방법, flags=Qt.MatchContains))

        if R.exec_():
            이름 = R.lineEdit_name.text()
            스크린번호 = int(R.lineEdit_screen_number.text())
            단위투자금 = int(R.lineEdit_unit.text()) * 10000
            매수방법 = R.comboBox_buy_sHogaGb.currentText().strip()[0:2]
            매도방법 = R.comboBox_sell_sHogaGb.currentText().strip()[0:2]

            시총상한 = int(R.lineEdit_max.text().strip())
            시총하한 = int(R.lineEdit_min.text().strip())
            포트폴리오수 = int(R.lineEdit_portsize.text().strip())

            robot.sName = 이름
            robot.Setting(sScreenNo=스크린번호, 단위투자금=단위투자금, 시총상한=시총상한, 시총하한=시총하한, 포트폴리오수=포트폴리오수, 매수방법=매수방법, 매도방법=매도방법)

    """
    def RobotAdd_TradeSuperValue(self):
        print("MainWindow : RobotAdd_TradeSuperValue")
        try:
            스크린번호 = self.GetUnAssignedScreenNumber()
            R = 화면_TradeSuperValue(parent=self)
            R.lineEdit_screen_number.setText('{:04d}'.format(스크린번호))
            if R.exec_():
                # 이름 = R.lineEdit_name.text()
                # 스크린번호 = int(R.lineEdit_screen_number.text())
                단위투자금 = int(R.lineEdit_unit.text()) * 10000
                매수방법 = R.comboBox_buy_condition.currentText().strip()[0:2]
                목표율 = float(R.lineEdit_targetrate.text().strip())
                손절율 = float(R.lineEdit_sellrate.text().strip())
                최대보유일 = int(R.lineEdit_holdday.text().strip())
                매도방법 = R.comboBox_sell_condition.currentText().strip()[0:2]

                strategy_list = list(R.data['매수전략'].unique())
                for strategy in strategy_list:
                    스크린번호 = self.GetUnAssignedScreenNumber()
                    이름 = strategy+'_'+ R.lineEdit_name.text()
                    종목리스트 = R.data[R.data['매수전략'] == strategy]
                    # print(종목리스트)
                    r = CTradeSuperValue(sName=이름, UUID=uuid.uuid4().hex, kiwoom=self.kiwoom, parent=self)
                    # r.Setting(sScreenNo=스크린번호, 단위투자금=단위투자금, 매수방법=매수방법, 목표율=목표율, 손절율=손절율, 최대보유일=최대보유일, 매도방법=매도방법, 종목리스트=종목리스트)
                    r.Setting(sScreenNo=스크린번호, 단위투자금=단위투자금, 매수방법=매수방법, 매도방법=매도방법, 종목리스트=종목리스트)
                    self.robots.append(r)

        except Exception as e:
            print('RobotAdd_TradeSuperValue', e)

    """

    def RobotAdd_TradeSuperValue(self):
        print("MainWindow : RobotAdd_TradeSuperValue")
        try:
            스크린번호 = self.GetUnAssignedScreenNumber()
            R = 화면_TradeSuperValue(parent=self)
            R.lineEdit_screen_number.setText('{:04d}'.format(스크린번호))
            if R.exec_():
                매수방법 = R.comboBox_buy_condition.currentText().strip()[0:2]
                매도방법 = R.comboBox_sell_condition.currentText().strip()[0:2]

                strategy_list = list(R.data['매수전략'].unique())
                print(strategy_list)
                for strategy in strategy_list:
                    스크린번호 = self.GetUnAssignedScreenNumber()
                    print("a")
                    이름 = str(strategy)+'_'+ R.lineEdit_name.text()
                    print("b")
                    종목리스트 = R.data[R.data['매수전략'] == strategy]
                    print(종목리스트)
                    r = CTradeSuperValue(sName=이름, UUID=uuid.uuid4().hex, kiwoom=self.kiwoom, parent=self)
                    r.Setting(sScreenNo=스크린번호, 매수방법=매수방법, 매도방법=매도방법, 종목리스트=종목리스트)
                    self.robots.append(r)

        except Exception as e:
            print('RobotAdd_TradeSuperValue', e)

    def RobotAutoAdd_TradeSuperValue(self, data, strategy):
        print("MainWindow : RobotAutoAdd_TradeSuperValue")
        try:
            스크린번호 = self.GetUnAssignedScreenNumber()
            이름 = strategy + '_TradeSuperValue'
            매수방법 = '00'
            매도방법 = '00'
            종목리스트 = data
            print('추가 종목리스트')
            print(종목리스트)

            r = CTradeSuperValue(sName=이름, UUID=uuid.uuid4().hex, kiwoom=self.kiwoom, parent=self)
            r.Setting(sScreenNo=스크린번호, 매수방법=매수방법, 매도방법=매도방법, 종목리스트=종목리스트)
            self.robots.append(r)
            print('로봇 자동추가 완료')
            Telegram('[XTrader]로봇 자동추가 완료')

        except Exception as e:
            print('RobotAutoAdd_TradeSuperValue', e)
            Telegram('[XTrader]로봇 자동추가 실패', e)

    """
    def RobotEdit_TradeSuperValue(self, robot):
        R = 화면_TradeSuperValue(parent=self)
        R.lineEdit_name.setText(robot.sName)
        R.lineEdit_screen_number.setText('{:04d}'.format(robot.sScreenNo))
        R.lineEdit_unit.setText(str(robot.단위투자금 // 10000))
        R.comboBox_buy_condition.setCurrentIndex(R.comboBox_buy_condition.findText(robot.매수방법, flags=Qt.MatchContains))
        R.lineEdit_targetrate.setText(str(robot.목표율))
        R.lineEdit_sellrate.setText(str(robot.손절율))
        R.lineEdit_holdday.setText(str(robot.최대보유일))
        R.comboBox_sell_condition.setCurrentIndex(R.comboBox_sell_condition.findText(robot.매도방법, flags=Qt.MatchContains))

        if R.exec_():
            이름 = R.lineEdit_name.text()
            스크린번호 = int(R.lineEdit_screen_number.text())
            단위투자금 = int(R.lineEdit_unit.text()) * 10000
            매수방법 = R.comboBox_buy_condition.currentText().strip()[0:2]
            목표율 = float(R.lineEdit_targetrate.text())
            손절율 = float(R.lineEdit_sellrate.text())
            최대보유일 = int(R.lineEdit_holdday.text())
            매도방법 = R.comboBox_sell_condition.currentText().strip()[0:2]
            종목리스트 = R.data

            robot.sName = 이름
            robot.Setting(sScreenNo=스크린번호, 단위투자금=단위투자금, 매수방법=매수방법, 목표율=목표율, 손절율=손절율, 최대보유일=최대보유일, 매도방법=매도방법, 종목리스트=종목리스트)
    """

    def RobotEdit_TradeSuperValue(self, robot):
        R = 화면_TradeSuperValue(parent=self)
        R.lineEdit_name.setText(robot.sName)
        R.lineEdit_screen_number.setText('{:04d}'.format(robot.sScreenNo))
        R.comboBox_buy_condition.setCurrentIndex(R.comboBox_buy_condition.findText(robot.매수방법, flags=Qt.MatchContains))
        R.comboBox_sell_condition.setCurrentIndex(R.comboBox_sell_condition.findText(robot.매도방법, flags=Qt.MatchContains))

        if R.exec_():
            이름 = R.lineEdit_name.text()
            스크린번호 = int(R.lineEdit_screen_number.text())
            매수방법 = R.comboBox_buy_condition.currentText().strip()[0:2]
            매도방법 = R.comboBox_sell_condition.currentText().strip()[0:2]
            종목리스트 = R.data
            print(종목리스트)

            robot.sName = 이름
            robot.Setting(sScreenNo=스크린번호, 매수방법=매수방법, 매도방법=매도방법, 종목리스트=종목리스트)

    def RobotAutoEdit_TradeSuperValue(self, robot, data):
        print("MainWindow : RobotAutoEdit_TradeSuperValue")
        try:
            이름 = robot.sName
            스크린번호 = int('{:04d}'.format(robot.sScreenNo))
            매수방법 = robot.매수방법
            매도방법 = robot.매도방법
            종목리스트 = data
            print('편집 종목리스트')
            print(종목리스트)

            robot.sName = 이름
            robot.Setting(sScreenNo=스크린번호, 매수방법=매수방법, 매도방법=매도방법, 종목리스트=종목리스트)
            print('로봇 자동편집 완료')
            Telegram('[XTrader]로봇 자동편집 완료')
        except Exception as e:
            print('RobotAutoAdd_TradeSuperValue', e)
            Telegram('[XTrader]로봇 자동편집 실패', e)

    def RobotAdd_TradeCondition(self):
        print("MainWindow : RobotAdd_TradeCondition")
        스크린번호 = self.GetUnAssignedScreenNumber()
        R = 화면_TradeCondition(sScreenNo=스크린번호, kiwoom=self.kiwoom, parent=self)
        R.lineEdit_screen_number.setText('{:04d}'.format(스크린번호))

        if R.exec_():
            이름 = R.lineEdit_name.text()
            스크린번호 = int(R.lineEdit_screen_number.text())
            단위투자비율 = int(R.lineEdit_unit.text())
            매수방법 = R.comboBox_buy_sHogaGb.currentText().strip()[0:2]
            매도방법 = R.comboBox_sell_sHogaGb.currentText().strip()[0:2]
            포트폴리오수 = int(R.lineEdit_portsize.text().strip())
            조건식인덱스 = R.df_condition['Index'][R.comboBox_condition.currentIndex()] # 조건식의 인덱스 넘김
            조건식명 = R.df_condition['Name'][R.comboBox_condition.currentIndex()] # 조건식의 이름 넘김
            종목리스트 = R.data

            r = CTradeCondition(sName=이름, UUID=uuid.uuid4().hex, kiwoom=self.kiwoom, parent=self)
            r.Setting(sScreenNo=스크린번호, 단위투자비율=단위투자비율, 포트폴리오수=포트폴리오수, 조건식인덱스 = 조건식인덱스, 조건식명 = 조건식명, 매수방법=매수방법,매도방법=매도방법, 종목리스트=종목리스트)

            self.robots.append(r)

    def RobotEdit_TradeCondition(self, robot):
        R = 화면_TradeCondition(sScreenNo=robot.sScreenNo, kiwoom=self.kiwoom, parent=self)
        R.lineEdit_name.setText(robot.sName)
        R.lineEdit_screen_number.setText('{:04d}'.format(robot.sScreenNo))
        R.lineEdit_unit.setText(str(robot.단위투자비율))
        R.lineEdit_portsize.setText(str(robot.포트폴리오수))
        R.comboBox_buy_sHogaGb.setCurrentIndex(R.comboBox_buy_sHogaGb.findText(robot.매수방법, flags=Qt.MatchContains))
        R.comboBox_sell_sHogaGb.setCurrentIndex(
            R.comboBox_sell_sHogaGb.findText(robot.매도방법, flags=Qt.MatchContains))

        if R.exec_():
            이름 = R.lineEdit_name.text()
            스크린번호 = int(R.lineEdit_screen_number.text())
            단위투자비율 = int(R.lineEdit_unit.text())
            매수방법 = R.comboBox_buy_sHogaGb.currentText().strip()[0:2]
            매도방법 = R.comboBox_sell_sHogaGb.currentText().strip()[0:2]
            포트폴리오수 = int(R.lineEdit_portsize.text().strip())
            조건식인덱스 = R.df_condition['Index'][R.comboBox_condition.currentIndex()]  # 조건식의 인덱스 넘김
            조건식명 = R.df_condition['Name'][R.comboBox_condition.currentIndex()]  # 조건식의 이름 넘김
            종목리스트 = R.data

            robot.sName = 이름
            robot.Setting(sScreenNo=스크린번호, 단위투자비율=단위투자비율, 포트폴리오수=포트폴리오수, 조건식인덱스=조건식인덱스, 조건식명=조건식명, 매수방법=매수방법, 매도방법=매도방법, 종목리스트=종목리스트)

    # -------------------------------------------
    # UI 관련함수
    # -------------------------------------------
    # 종목코드 생성
    def StockCodeBuild(self, to_db=False):
        try:
            result = []

            markets = [['0', 'KOSPI'], ['10', 'KOSDAQ'], ['8', 'ETF']]
            for [marketcode, marketname] in markets:
                codelist = self.kiwoom.dynamicCall('GetCodeListByMarket(QString)', [
                    marketcode])  # sMarket – 0:장내, 3:ELW, 4:뮤추얼펀드, 5:신주인수권, 6:리츠, 8:ETF, 9:하이일드펀드, 10:코스닥, 30:제3시장
                codes = codelist.split(';')

                for code in codes:
                    if code is not '':
                        종목명 = self.kiwoom.dynamicCall('GetMasterCodeName(QString)', [code])
                        주식수 = self.kiwoom.dynamicCall('GetMasterListedStockCnt(QString)', [code])
                        감리구분 = self.kiwoom.dynamicCall('GetMasterConstruction(QString)',
                                                       [code])  # 감리구분 – 정상, 투자주의, 투자경고, 투자위험, 투자주의환기종목
                        상장일 = datetime.datetime.strptime(
                            self.kiwoom.dynamicCall('GetMasterListedStockDate(QString)', [code]), '%Y%m%d')
                        전일종가 = int(self.kiwoom.dynamicCall('GetMasterLastPrice(QString)', [code]))
                        종목상태 = self.kiwoom.dynamicCall('GetMasterStockState(QString)', [
                            code])  # 종목상태 – 정상, 증거금100%, 거래정지, 관리종목, 감리종목, 투자유의종목, 담보대출, 액면분할, 신용가능

                        result.append([marketname, code, 종목명, 주식수, 감리구분, 상장일, 전일종가, 종목상태])

            df_code = DataFrame(data=result, columns=['시장구분', '종목코드', '종목명', '주식수', '감리구분', '상장일', '전일종가', '종목상태'])
            # df.set_index('종목코드', inplace=True)

            if to_db == True:
                df_code['상장일'] = df_code['상장일'].apply(lambda x: (x.to_pydatetime()).strftime('%Y-%m-%d %H:%M:%S'))
                # print(df_code.values.tolist())
                conn = sqliteconn()
                df_code.to_sql('종목코드', conn, if_exists='replace', index=False)
                conn.close()

            return df_code

        except Exception as e:
            print('StockCodeBuild', e)

    # 유틸리티 함수
    def kiwoom_doc(self):
        kiwoom = QAxContainer.QAxWidget("KHOPENAPI.KHOpenAPICtrl.1")
        _doc = kiwoom.generateDocumentation()
        f = open("openapi_doc.html", 'w')
        f.write(_doc)
        f.close()

    def ReguestPriceDaily(self, _repeat=0):
        logger.info("일별가격정보백업: %s" % self.종목코드)
        self.statusbar.showMessage("주식일봉백업: %s" % self.종목코드)
        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, Qstring)', "종목코드", self.종목코드[0])
        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, Qstring)', "기준일자", self.기준일자)
        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, Qstring)', "수정주가구분", '1')
        ret = self.kiwoom.dynamicCall('CommRqData(QString, QString, int, QString)', "주식일봉차트조회", "OPT10081", _repeat,
                                      '{:04d}'.format(self.ScreenNumber))

    def ReguestPriceMin(self, _repeat=0):
        # logger.info("주식분봉백업: %s" % self.종목코드)
        self.statusbar.showMessage("주식분봉백업: %s" % self.종목코드)
        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, Qstring)', "종목코드", self.종목코드[0])
        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, Qstring)', "틱범위", self.틱범위)
        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, Qstring)', "수정주가구분", '1')
        ret = self.kiwoom.dynamicCall('CommRqData(QString, QString, int, QString)', "주식분봉차트조회", "OPT10080", _repeat,
                                      '{:04d}'.format(self.ScreenNumber))

    def RequestInvestorDaily(self, _repeat=0):
        logger.info("종목별투자자백업: %s" % self.종목코드)
        self.statusbar.showMessage("종목별투자자백업: %s" % self.종목코드)
        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, Qstring)', "일자", self.기준일자)
        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, Qstring)', "종목코드", self.종목코드[0])
        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, int)', "금액수량구분", 2)  # 1:금액, 2:수량
        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, int)', "매매구분", 0)  # 0:순매수, 1:매수, 2:매도
        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, int)', "단위구분", 1)  # 1000:천주, 1:단주
        ret = self.kiwoom.dynamicCall('CommRqData(QString, QString, int, QString)', "종목별투자자조회", "OPT10060", _repeat,
                                      '{:04d}'.format(self.ScreenNumber))

if __name__ == "__main__":
    # 1.로그 인스턴스를 만든다.
    logger = logging.getLogger('stocktrader')
    # 2.formatter를 만든다.
    formatter = logging.Formatter('[%(levelname)s|%(filename)s:%(lineno)s]%(asctime)s>%(message)s')

    loggerLevel = logging.DEBUG
    filename = "LOG/stocktrader.log"

    # 스트림과 파일로 로그를 출력하는 핸들러를 각각 만든다.
    filehandler = logging.FileHandler(filename)
    streamhandler = logging.StreamHandler()

    # 각 핸들러에 formatter를 지정한다.
    filehandler.setFormatter(formatter)
    streamhandler.setFormatter(formatter)

    # 로그 인스턴스에 스트림 핸들러와 파일 핸들러를 붙인다.
    logger.addHandler(filehandler)
    logger.addHandler(streamhandler)
    logger.setLevel(loggerLevel)
    logger.debug("=============================================================================")
    logger.info("LOG START")

    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(True)

    window = MainWindow()
    window.show()

    Telegram("[XTrader]프로그램이 실행되었습니다.")

    # 프로그램 실행 후 3초 후에 한번 신호 받고, 그 다음 1초 마다 신호를 계속 받음
    QTimer().singleShot(3, window.OnQApplicationStarted)  # 3초 후에 한번만(singleShot) 신호받음 : MainWindow.OnQApplicationStarted 실행

    clock = QtCore.QTimer()
    clock.timeout.connect(window.OnClockTick)  # 1초마다 현재시간 읽음 : MainWindow.OnClockTick 실행
    clock.start(1000)  # 기존값 1000, 1초마다 신호받음

    sys.exit(app.exec_())
