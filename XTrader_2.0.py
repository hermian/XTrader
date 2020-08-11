
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

# Google SpreadSheet Read/Write
import gspread # (추가 설치 모듈)
from oauth2client.service_account import ServiceAccountCredentials # (추가 설치 모듈)
from df2gspread import df2gspread as d2g # (추가 설치 모듈)
from string import ascii_uppercase # 알파벳 리스트

from bs4 import BeautifulSoup
import requests

import logging
import logging.handlers

import sqlite3

import telepot # 텔레그램봇(추가 설치 모듈)

import csv

# Google Spreadsheet Setting *******************************
scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
json_file_name = './secret/xtrader-276902-f5a8b77e2735.json'

credentials = ServiceAccountCredentials.from_json_keyfile_name(json_file_name, scope)
gc = gspread.authorize(credentials)

# XTrader-Stocklist URL
spreadsheet_url = 'https://docs.google.com/spreadsheets/d/1XE4sk0vDw4fE88bYMDZuJbnP4AF9CmRYHKY6fCXABw4/edit#gid=0' # Sheeet

# spreadsheet 연결 및 worksheet setting
doc = gc.open_by_url(spreadsheet_url)

shortterm_buy_sheet = doc.worksheet('매수모니터링')
shortterm_sell_sheet = doc.worksheet('매도모니터링')
shortterm_strategy_sheet = doc.worksheet('ST bot')
shortterm_history_sheet = doc.worksheet('매매이력')

shortterm_history_cols = ['번호', '종목명', '매수가', '매수수량', '매수일', '매수조건', '매도가', '매도수량',
                          '매도일', '매도전략', '매도구간', '수익률(계산)','수익률', '수익금', '세금+수수료', '확정 수익금']

# 구글 스프레드시트 업데이트를 위한 알파벳리스트(열 이름 얻기위함)
alpha_list = list(ascii_uppercase)


# SQLITE DB Setting *****************************************
DATABASE = 'stockdata.db'
def sqliteconn():
    conn = sqlite3.connect(DATABASE)
    return conn

# DB에서 종목명으로 종목코드, 종목영, 시장구분 반환
def get_code(종목명체크):
    # 종목명이 띄워쓰기, 대소문자 구분이 잘못될 것을 감안해서
    # DB 저장 시 종목명체크 컬럼은 띄워쓰기 삭제 및 소문자로 저장됨
    # 구글에서 받은 종목명을 띄워쓰기 삭제 및 소문자로 바꿔서 종목명체크와 일치하는 데이터 저장
    # 종목명은 DB에 있는 정상 종목명으로 사용하도록 리턴
    종목명체크 = 종목명체크.lower().replace(' ', '')
    query = """
                select 종목코드, 종목명, 시장구분
                from 종목코드
                where (종목명체크 = '%s')
            """ % (종목명체크)
    conn = sqliteconn()
    df = pd.read_sql(query, con=conn)
    conn.close()

    return list(df[['종목코드', '종목명', '시장구분']].values)[0]

# 종목코드가 int형일 경우 정상적으로 반환
def fix_stockcode(data):
    if len(data)< 6:
        for i in range(6 - len(data)):
            data = '0'+data
    return data

# 구글 스프레드 시트 Import후 DataFrame 반환
def import_googlesheet():
    try:
        # 1. 매수 모니터링 시트 체크 및 매수 종목 선정
        row_data = shortterm_buy_sheet.get_all_values() # 구글 스프레드시트 '매수모니터링' 시트 데이터 get

        # 작성 오류 체크를 위한 주요 항목의 위치(index)를 저장
        idx_strategy = row_data[0].index('기본매도전략')
        idx_buyprice = row_data[0].index('매수가1')
        idx_sellprice = row_data[0].index('목표가')

        # DB에서 받아올 종목코드와 시장 컬럼 추가
        # 번호, 종목명, 매수모니터링, 비중, 시가위치, 매수가1, 매수가2, 매수가3, 기존매도전략, 목표가
        row_data[0].insert(2, '종목코드')
        row_data[0].insert(3, '시장')

        for row in row_data[1:]:
            try:
                code, name, market = get_code(row[1])  # 종목명으로 종목코드, 종목명, 시장 받아서(get_code 함수) 추가
            except Exception as e:
                name = ''
                code = ''
                market = ''
                print('구글 매수모니터링 시트 종목명 오류 : %s' % (row[1]))
                logger.error('구글 매수모니터링 시트 오류 : %s' % (row[1]))
                Telegram('[XTrader]구글 매수모니터링 시트 오류 : %s' % (row[1]))

            row[1] = name # 정상 종목명으로 저장
            row.insert(2, code)
            row.insert(3, market)

        data = pd.DataFrame(data=row_data[1:], columns=row_data[0])

        # 사전 데이터 정리
        data = data[(data['매수모니터링'] == '1') & (data['종목코드']!= '')]
        data = data[row_data[0][:row_data[0].index('목표가')+1]]
        del data['매수모니터링']

        data.to_csv('%s_googlesheetdata.csv'%(datetime.date.today().strftime('%Y%m%d')), encoding='euc-kr', index=False)

        # 2. 매도 모니터링 시트 체크(번호, 종목명, 보유일, 매도전략, 매도가)
        row_data = shortterm_sell_sheet.get_all_values()  # 구글 스프레드시트 '매도모니터링' 시트 데이터 get

        # 작성 오류 체크를 위한 주요 항목의 위치(index)를 저장
        idx_holding = row_data[0].index('보유일')
        idx_strategy = row_data[0].index('매도전략')
        idx_loss = row_data[0].index('손절가')
        idx_sellprice = row_data[0].index('목표가')

        if len(row_data) > 1:
            for row in row_data[1:]:
                try:
                    code, name, market = get_code(row[1])  # 종목명으로 종목코드, 종목명, 시장 받아서(get_code 함수) 추가
                    if row[idx_holding] == '' : raise Exception('보유일 오류')
                    if row[idx_strategy] == '': raise Exception('매도전략 오류')
                    if row[idx_loss] == '': raise Exception('손절가 오류')
                    if row[idx_strategy] == '4' and row[idx_sellprice] == '': raise Exception('목표가 오류')
                except Exception as e:
                    if str(e) != '보유일 오류' and str(e) != '매도전략 오류' and str(e) != '손절가 오류'and str(e) != '목표가 오류': e = '종목명 오류'
                    print('구글 매도모니터링 시트 오류 : %s, %s' % (row[1], e))
                    logger.error('구글 매도모니터링 시트 오류 : %s, %s' % (row[1], e))
                    Telegram('[XTrader]구글 매도모니터링 시트 오류 : %s, %s' % (row[1], e))

        # print(data)
        print('[XTrader]구글 시트 확인 완료')
        # Telegram('[XTrader]구글 시트 확인 완료')
        # logger.info('[XTrader]구글 시트 확인 완료')

        return data

    except Exception as e:
        # 구글 시트 import error시 에러 없어을 때 백업한 csv 읽어옴
        print("import_googlesheet Error : %s"%e)
        logger.error("import_googlesheet Error : %s"%e)
        backup_file = datetime.date.today().strftime('%Y%m%d') + '_googlesheetdata.csv'
        if backup_file in os.listdir():
            data = pd.read_csv(backup_file, encoding='euc-kr')
            data = data.fillna('')
            data = data.astype(str)
            data['종목코드'] = data['종목코드'].apply(fix_stockcode)

            print("import googlesheet backup_file")
            logger.info("import googlesheet backup_file")

            return data


# Telegram Setting *****************************************
with open('./secret/telegram_token.txt', mode='r') as tokenfile:
    TELEGRAM_TOKEN = tokenfile.readline().strip()
with open('./secret/chatid.txt', mode='r') as chatfile:
    CHAT_ID = int(chatfile.readline().strip())
bot = telepot.Bot(TELEGRAM_TOKEN)

with open('./secret/Telegram.txt', mode='r') as tokenfile:
    r = tokenfile.read()
    TELEGRAM_TOKEN_yoo = r.split('\n')[0].split(', ')[1]
    CHAT_ID_yoo = r.split('\n')[1].split(', ')[1]
bot_yoo = telepot.Bot(TELEGRAM_TOKEN_yoo)

telegram_enable = False
def Telegram(str, send='all'):
    try:
        if telegram_enable == True:
            if send == 'mc':
                bot.sendMessage(CHAT_ID, str)
            else:
                bot.sendMessage(CHAT_ID, str)
                bot_yoo.sendMessage(CHAT_ID_yoo, str)
        else:
            pass
    except Exception as e:
        Telegram('[XTrader]Telegram Error : %s' % e, send='mc')


# 매수 후 보유기간 계산 *****************************************
today = datetime.date.today()
def holdingcal(base_date, excluded=(6, 7)):  # 2018-06-23
    yy = int(base_date[:4])  # 연도
    mm = int(base_date[5:7])  # 월
    dd = int(base_date[8:10])  # 일

    base_d = datetime.date(yy, mm, dd)

    delta = 0
    while base_d <= today:
        if base_d.isoweekday() not in excluded:
            delta += 1
        base_d += datetime.timedelta(days=1)

    return delta

# 호가 계산(상한가, 현재가) *************************************
def hogacal(price, diff, market, option):
    # diff 0 : 상한가 호가, -1 : 상한가 -1호가
    if option == '현재가':
        cal_price = price
    elif option == '상한가':
        cal_price = price * 1.3

    if cal_price < 1000:
        hogaunit = 1
    elif cal_price < 5000:
        hogaunit = 5
    elif cal_price < 10000:
        hogaunit = 10
    elif cal_price < 50000:
        hogaunit = 50
    elif cal_price < 100000 and market == "KOSPI":
        hogaunit = 100
    elif cal_price < 500000 and market == "KOSPI":
        hogaunit = 500
    elif cal_price >= 500000 and market == "KOSPI":
        hogaunit = 1000
    elif cal_price >= 50000 and market == "KOSDAQ":
        hogaunit = 100

    cal_price = int(cal_price / hogaunit) * hogaunit + (hogaunit * diff)

    return cal_price


로봇거래계좌번호 = None

주문딜레이 = 0.25
초당횟수제한 = 5

## 키움증권 제약사항 - 3.7초에 한번 읽으면 지금까지는 괜찮음
주문지연 = 3700 # 3.7초

로봇스크린번호시작 = 9000
로봇스크린번호종료 = 9999

# Table View 데이터 정리
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


# 포트폴리오에 사용되는 주식정보 클래스
# TradeShortTerm용 포트폴리오
class CPortStock_ShortTerm(object):
    def __init__(self, 번호, 매수일, 종목코드, 종목명, 시장, 매수가, 매수조건, 보유일, 매도전략, 매도구간별조건, 매도구간=1, 매도가=0, 수량=0):
        self.번호 = 번호
        self.매수일 = 매수일
        self.종목코드 = 종목코드
        self.종목명 = 종목명
        self.시장 = 시장
        self.매수가 = 매수가
        self.매수조건 = 매수조건
        self.보유일 = 보유일
        self.매도전략 = 매도전략
        self.매도구간별조건 = 매도구간별조건
        self.매도구간 = 매도구간
        self.매도가 = 매도가
        self.수량 = 수량

        if self.매도전략 == '2' or self.매도전략 == '3':
            self.목표도달 = False # 목표가(매도가) 도달 체크(False 상태로 구간 컷일경우 전량 매도)
            self.매도조건 = '' # 구간매도 : B, 목표매도 : T
        elif self.매도전략 == '4':
            self.sellcount = 0
            self.매도단위수량 = 0 # 전략4의 기본 매도 단위는 보유수량의 1/3
            self.익절가1도달 = False
            self.익절가2도달 = False
            self.목표가도달 = False

# CTrade 거래로봇용 베이스클래스 : OpenAPI와 붙어서 주문을 내는 등을 하는 클래스
class CTrade(object):
    def __init__(self, sName, UUID, kiwoom=None, parent=None):
        """
        :param sName: 로봇이름
        :param UUID: 로봇구분용 id
        :param kiwoom: 키움OpenAPI
        :param parent: 나를 부른 부모 - 보통은 메인윈도우
        """
        # print("CTrade : __init__")

        self.sName = sName
        self.UUID = UUID

        self.sAccount = None  # 거래용계좌번호
        self.kiwoom = kiwoom
        self.parent = parent

        self.running = False  # 실행상태

        self.portfolio = dict()  # 포트폴리오 관리 {'종목코드':종목정보}
        self.현재가 = dict()  # 각 종목의 현재가

    # 계좌 보유 종목 받음
    def InquiryList(self, _repeat=0):
        # print("CTrade : InquiryList")
        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, Qstring)', "계좌번호", self.sAccount)
        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, Qstring)', "비밀번호입력매체구분", '00')
        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, Qstring)', "조회구분", '1')
        ret = self.kiwoom.dynamicCall('CommRqData(QString, QString, int, QString)', "계좌평가잔고내역요청", "opw00018", _repeat, '{:04d}'.format(self.sScreenNo))

        self.InquiryLoop = QEventLoop()  # 로봇에서 바로 쓸 수 있도록하기 위해서 계좌 조회해서 종목을 받고나서 루프해제시킴
        self.InquiryLoop.exec_()

    # 금일 매도 종목에 대해서 수익률, 수익금, 수수료 요청(일별종목별실현손익요청)
    def DailyProfit(self, 금일매도종목):
        _repeat = 0
        # self.sAccount = 로봇거래계좌번호
        # self.sScreenNo = self.ScreenNumber
        시작일자 = datetime.date.today().strftime('%Y%m%d')

        cnt = 1
        for 종목코드 in 금일매도종목:
            # print(self.sScreenNo, 종목코드, 시작일자)
            self.update_cnt = len(금일매도종목) - cnt
            cnt += 1
            ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, Qstring)', "계좌번호", self.sAccount)
            ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, Qstring)', "종목코드", 종목코드)
            ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, Qstring)', "시작일자", 시작일자)
            ret = self.kiwoom.dynamicCall('CommRqData(QString, QString, int, QString)', "일자별종목별실현손익요청", "OPT10072",
                                          _repeat, '{:04d}'.format(self.sScreenNo))

            self.DailyProfitLoop = QEventLoop()  # 로봇에서 바로 쓸 수 있도록하기 위해서 계좌 조회해서 종목을 받고나서 루프해제시킴
            self.DailyProfitLoop.exec_()

    # 일별종목별실현손익 응답 결과 구글 업로드
    def DailyProfitUpload(self, 매도결과):
        # 매도결과 ['종목명','체결량','매입단가','체결가','당일매도손익','손익율','당일매매수수료','당일매매세금']
        print(매도결과)

        if self.sName == 'TradeShortTerm':
            history_sheet = shortterm_history_sheet
            history_cols = shortterm_history_cols

        code_row = history_sheet.findall(매도결과[0])[-1].row

        계산수익률 = round((int(float(매도결과[3])) / int(float(매도결과[2])) - 1) * 100, 2)

        cell = alpha_list[history_cols.index('매수가')] + str(code_row)  # 매입단가
        history_sheet.update_acell(cell, int(float(매도결과[2])))

        cell = alpha_list[history_cols.index('매도가')] + str(code_row)  # 체결가
        history_sheet.update_acell(cell, int(float(매도결과[3])))

        cell = alpha_list[history_cols.index('수익률(계산)')] + str(code_row)  # 수익률 계산
        history_sheet.update_acell(cell, 계산수익률)

        cell = alpha_list[history_cols.index('수익률')] + str(code_row)  # 손익율
        history_sheet.update_acell(cell, 매도결과[5])

        cell = alpha_list[history_cols.index('수익금')] + str(code_row)  # 손익율
        history_sheet.update_acell(cell, int(float(매도결과[4])))

        cell = alpha_list[history_cols.index('세금+수수료')] + str(code_row)  # 당일매매수수료 + 당일매매세금
        history_sheet.update_acell(cell, int(float(매도결과[6])) + int(float(매도결과[7])))

        self.DailyProfitLoop.exit()

        if self.update_cnt == 0:
            print('금일 실현 손익 구글 업로드 완료')

            Telegram("[XTrader]금일 실현 손익 구글 업로드 완료")
            logger.info("[XTrader]금일 실현 손익 구글 업로드 완료")

    # 포트폴리오의 상태
    def GetStatus(self):
        # print("CTrade : GetStatus")
        try:
            result = []
            for p, v in self.portfolio.items():
                result.append('%s(%s)[P%s/V%s/D%s]' % (v.종목명.strip(), v.종목코드, v.매수가, v.수량, v.매수일))

            return [self.__class__.__name__, self.sName, self.UUID, self.sScreenNo, self.running, len(self.portfolio), ','.join(result)]
        except Exception as e:
            print('CTrade_GetStatus Error', e)
            logger.error('CTrade_GetStatus Error : %s' % e)

    def GenScreenNO(self):
        """
        :return: 키움증권에서 요구하는 스크린번호를 생성
        """
        # print("CTrade : GenScreenNO")
        self.SmallScreenNumber += 1
        if self.SmallScreenNumber > 9999:
            self.SmallScreenNumber = 0

        return self.sScreenNo * 10000 + self.SmallScreenNumber

    def GetLoginInfo(self, tag):
        """
        :param tag:
        :return: 로그인정보 호출
        """
        # print("CTrade : GetLoginInfo")
        return self.kiwoom.dynamicCall('GetLoginInfo("%s")' % tag)

    def KiwoomConnect(self):
        """
        :return: 키움증권OpenAPI의 CallBack에 대응하는 처리함수를 연결
        """
        # print("CTrade : KiwoomConnect")
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
            print("CTrade : KiwoomConnect Error :", e)

        # logger.info("%s : connected" % self.sName)

    def KiwoomDisConnect(self):
        """
        :return: Callback 연결해제
        """
        # print("CTrade : KiwoomDisConnect")
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
        # print("CTrade : KiwoomAccount")
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
        # print("CTrade : KiwoomSendOrder")
        try:
            order = self.kiwoom.dynamicCall(
                'SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)',
                [sRQName, sScreenNo, sAccNo, nOrderType, sCode, nQty, nPrice, sHogaGb, sOrgOrderNo])
            return order
        except Exception as e:
            print('CTradeShortTerm_KiwoomSendOrder Error ', e)
            Telegram('[XTrader]CTradeShortTerm_KiwoomSendOrder Error: %s' % e, send='mc')
            logger.error('CTradeShortTerm_KiwoomSendOrder Error : %s' % e)
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
        # print("CTrade : KiwoomSetRealReg")
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
        # print("CTrade : KiwoomSetRealRemove")
        ret = self.kiwoom.dynamicCall('SetRealRemove(QString, QString)', sScreenNo, sCode)
        return ret

    def OnEventConnect(self, nErrCode):
        """
        OpenAPI 메뉴얼 참조
        :param nErrCode:
        :return:
        """
        # print("CTrade : OnEventConnect")
        logger.debug('OnEventConnect', nErrCode)

    def OnReceiveMsg(self, sScrNo, sRQName, sTRCode, sMsg):
        """
        OpenAPI 메뉴얼 참조
        :param sScrNo:
        :param sRQName:
        :param sTRCode:
        :param sMsg:
        :return:
        """
        # print("CTrade : OnReceiveMsg")
        logger.debug('OnReceiveMsg [%s] [%s] [%s] [%s]' % (sScrNo, sRQName, sTRCode, sMsg))
        # self.InquiryLoop.exit()

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
        # print('CTrade : OnReceiveTrData')
        try:
            logger.debug('OnReceiveTrData [%s] [%s] [%s] [%s] [%s] [%s] [%s] [%s] [%s] ' % (sScrNo, sRQName, sTRCode, sRecordName, sPreNext, nDataLength, sErrorCode, sMessage, sSPlmMsg))

            if self.sScreenNo != int(sScrNo[:4]):
                return

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

            if sRQName == "일자별종목별실현손익요청":
                try:
                    data_idx = ['종목명', '체결량', '매입단가', '체결가', '당일매도손익', '손익율', '당일매매수수료', '당일매매세금']

                    result = []
                    for idx in data_idx:
                        data = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode,
                                                       "",
                                                       sRQName, 0, idx)
                        result.append(data.strip())

                    self.DailyProfitUpload(result)

                except Exception as e:
                    print(e)
                    logger.error('일자별종목별실현손익요청 Error : %s' % e)

        except Exception as e:
            print('CTradeShortTerm_OnReceiveTrData Error ', e)
            Telegram('[XTrader]CTradeShortTerm_OnReceiveTrData Error : %s' % e, send='mc')
            logger.error('CTradeShortTerm_OnReceiveTrData Error : %s' % e)

    def OnReceiveChejanData(self, sGubun, nItemCnt, sFidList):
        """
        OpenAPI 메뉴얼 참조
        :param sGubun:
        :param nItemCnt:
        :param sFidList:
        :return:
        """
        # logger.debug('OnReceiveChejanData [%s] [%s] [%s]' % (sGubun, nItemCnt, sFidList))

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
        # print("CTrade : OnReceiveChejanData")
        try:
            # 접수
            if sGubun == "0":
                # logger.debug('OnReceiveChejanData: 접수 [%s] [%s] [%s]' % (sGubun, nItemCnt, sFidList))

                화면번호 = self.kiwoom.dynamicCall('GetChejanData(QString)', 920)

                if len(화면번호.replace(' ','')) == 0 : # 로봇 실행중 영웅문으로 주문 발생 시 화면번호가 '    '로 들어와 에러발생함 방지
                    print('다른 프로그램을 통한 거래 발생')
                    Telegram('다른 프로그램을 통한 거래 발생', send='mc')
                    logger.info('다른 프로그램을 통한 거래 발생')
                    return
                elif self.sScreenNo != int(화면번호[:4]):
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

                logger.debug('접수 - 주문상태:{주문상태} 계좌번호:{계좌번호} 체결시간:{체결시간} 주문번호:{주문번호} 체결번호:{체결번호} 종목코드:{종목코드} 종목명:{종목명} 체결량:{체결량} 체결가:{체결가} 단위체결가:{단위체결가} 주문수량:{주문수량} 체결수량:{체결수량} 단위체결량:{단위체결량} 미체결수량:{미체결수량} 당일매매수수료:{당일매매수수료} 당일매매세금:{당일매매세금}'.format(**param))

                if param["주문상태"] == "접수":
                    self.접수처리(param)
                if param["주문상태"] == "체결": # 매도의 경우 체결로 안들어옴
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

                logger.debug('잔고통보 - 계좌번호:{계좌번호} 종목명:{종목명} 보유수량:{보유수량} 매입단가:{매입단가} 총매입가:{총매입가} 손익율:{손익율} 당일총매도손익:{당일총매도손익} 당일순매수량:{당일순매수량}'.format(**param))

                self.잔고처리(param)

            # 특이신호
            if sGubun == "3":
                logger.debug('OnReceiveChejanData: 특이신호 [%s] [%s] [%s]' % (sGubun, nItemCnt, sFidList))

                pass

        except Exception as e:
            print('CTradeShortTerm_OnReceiveChejanData Error ', e)
            Telegram('[XTrader]CTradeShortTerm_OnReceiveChejanData Error : %s' % e, send='mc')
            logger.error('CTradeShortTerm_OnReceiveChejanData Error : %s' % e)

    def OnReceiveRealData(self, sRealKey, sRealType, sRealData):
        """
        OpenAPI 메뉴얼 참조
        :param sRealKey:
        :param sRealType:
        :param sRealData:
        :return:
        """
        # logger.debug('OnReceiveRealData [%s] [%s] [%s]' % (sRealKey, sRealType, sRealData))
        _now = datetime.datetime.now()
        try:
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

                self.실시간데이터처리(param)

        except Exception as e:
            print('CTradeShortTerm_OnReceiveRealData Error ', e)
            Telegram('[XTrader]CTradeShortTerm_OnReceiveRealData Error : %s' % e, send='mc')
            logger.error('CTradeShortTerm_OnReceiveRealData Error : %s' % e)

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
        try:
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

        except Exception as e:
            print('CTradeShortTerm_정액매수 Error ', e)
            Telegram('[XTrader]CTradeShortTerm_정액매수 Error : %s' % e, send='mc')
            logger.error('CTradeShortTerm_정액매수 Error : %s' % e)

    def 정량매도(self, sRQName, 종목코드, 매도가, 수량):
        # sRQName = '정량매도%s' % self.sScreenNo
        try:
            sScreenNo = self.GenScreenNO()
            sAccNo = self.sAccount
            nOrderType = 2  # (1:신규매수, 2:신규매도 3:매수취소, 4:매도취소, 5:매수정정, 6:매도정정)
            sCode = 종목코드
            nQty = 수량
            nPrice = 매도가
            sHogaGb = self.매도방법  # 00:지정가, 03:시장가, 05:조건부지정가, 06:최유리지정가, 07:최우선지정가, 10:지정가IOC, 13:시장가IOC, 16:최유리IOC, 20:지정가FOK, 23:시장가FOK, 26:최유리FOK, 61:장개시전시간외, 62:시간외단일가매매, 81:시간외종가
            if sHogaGb in ['03', '06', '07']:
                nPrice = 0
            sOrgOrderNo = 0

            ret = self.parent.KiwoomSendOrder(sRQName, sScreenNo, sAccNo, nOrderType, sCode, nQty, nPrice, sHogaGb,
                                              sOrgOrderNo)

            return ret
        except Exception as e:
            print('CTradeShortTerm_정량매도 Error ', e)
            Telegram('[XTrader]CTradeShortTerm_정량매도 Error : %s' % e, send='mc')
            logger.error('CTradeShortTerm_정량매도 Error : %s' % e)

    def 정액매도(self, sRQName, 종목코드, 매도가, 수량):
        # sRQName = '정액매도%s' % self.sScreenNo
        sScreenNo = self.GenScreenNO()
        sAccNo = self.sAccount
        nOrderType = 2  # (1:신규매수, 2:신규매도 3:매수취소, 4:매도취소, 5:매수정정, 6:매도정정)
        sCode = 종목코드
        nQty = 수량
        nPrice = 매도가
        sHogaGb = '00' #self.매도방법  # 00:지정가, 03:시장가, 05:조건부지정가, 06:최유리지정가, 07:최우선지정가, 10:지정가IOC, 13:시장가IOC, 16:최유리IOC, 20:지정가FOK, 23:시장가FOK, 26:최유리FOK, 61:장개시전시간외, 62:시간외단일가매매, 81:시간외종가
        if sHogaGb in ['03', '06', '07']:
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

    def OnReceiveTrData(self, sScrNo, sRQName, sTRCode, sRecordName, sPreNext, nDataLength, sErrorCode, sMessage, sSPlmMsg):
        # logger.debug('OnReceiveTrData [%s] [%s] [%s] [%s] [%s] [%s] [%s] [%s] [%s] ' % (sScrNo, sRQName, sTRCode, sRecordName, sPreNext, nDataLength, sErrorCode, sMessage, sSPlmMsg))
        print('화면_분별주가 : OnReceiveTrData')
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
            # df = DataFrame(data=self.result, columns=self.columns)
            # df.to_csv('분봉.csv', encoding='euc-kr')
            if sPreNext == '2':
                QTimer.singleShot(주문지연, lambda: self.Request(_repeat=2))
            else:
                df = DataFrame(data=self.result, columns=self.columns)
                df.to_csv('분봉.csv', encoding='euc-kr', index=False)
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
                df_new = df[['종목코드'] + self.columns]
                self.model.update(df_new)
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


Ui_TradeShortTerm, QtBaseClass_TradeShortTerm = uic.loadUiType("./UI/TradeShortTerm.ui")
class 화면_TradeShortTerm(QDialog, Ui_TradeShortTerm):
    def __init__(self, parent):
        super(화면_TradeShortTerm, self).__init__(parent)
        self.setupUi(self)

        self.parent = parent

        self.model = PandasModel()
        self.tableView.setModel(self.model)

        self.result = []

    def inquiry(self):
        # Google spreadsheet 사용
        try:
            self.data = import_googlesheet()
            print(self.data)

            self.model.update(self.data)
            for i in range(len(self.data)):
                self.tableView.resizeColumnToContents(i)

        except Exception as e:
            print('화면_TradeShortTerm : inquiry Error ', e)
            logger.error('화면_TradeShortTerm : inquiry Error : %s' % e)


class CTradeShortTerm(CTrade):  # 로봇 추가 시 __init__ : 복사, Setting, 초기조건:전략에 맞게, 데이터처리~Run:복사
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
        self.매수모니터링체크 = False

        self.SmallScreenNumber = 9999

        self.d = today

    # 구글 스프레드시트에서 읽은 DataFrame에서 로봇별 종목리스트 셋팅
    def set_stocklist(self, data):
        self.Stocklist = dict()
        self.Stocklist['컬럼명'] = list(data.columns)
        for 종목코드 in data['종목코드'].unique():
            temp_list = data[data['종목코드'] == 종목코드].values[0]
            self.Stocklist[종목코드] = {
                '번호': temp_list[self.Stocklist['컬럼명'].index('번호')],
                '종목명': temp_list[self.Stocklist['컬럼명'].index('종목명')],
                '종목코드': 종목코드,
                '시장': temp_list[self.Stocklist['컬럼명'].index('시장')],
                '투자비중': float(temp_list[self.Stocklist['컬럼명'].index('비중')]),  # 저장 후 setting 함수에서 전략의 단위투자금을 곱함
                '시가위치': list(map(float, temp_list[self.Stocklist['컬럼명'].index('시가위치')].split(','))),
                '매수가': list(
                    int(float(temp_list[list(data.columns).index(col)].replace(',', ''))) for col in data.columns if
                    '매수가' in col and temp_list[list(data.columns).index(col)] != ''),
                '매도전략': temp_list[self.Stocklist['컬럼명'].index('기본매도전략')],
                '매도가': list(
                    int(float(temp_list[list(data.columns).index(col)].replace(',', ''))) for col in data.columns if
                    '목표가' in col and temp_list[list(data.columns).index(col)] != '')
            }
        return self.Stocklist

    # RobotAdd 함수에서 초기화 다음 셋팅 실행해서 설정값 넘김
    def Setting(self, sScreenNo, 매수방법='00', 매도방법='03', 종목리스트=pd.DataFrame()):
        try:
            self.sScreenNo = sScreenNo
            self.실시간종목리스트 = []
            self.매수방법 = 매수방법
            self.매도방법 = 매도방법
            self.종목리스트 = 종목리스트

            self.Stocklist = self.set_stocklist(self.종목리스트)  # 번호, 종목명, 종목코드, 시장, 비중, 시가위치, 매수가, 매도전략, 매도가
            self.Stocklist['전략'] = {
                '단위투자금': '',
                '모니터링종료시간': '',
                '보유일': '',
                '투자금비중': '',
                '매도구간별조건': [],
                '전략매도가': [],
            }

            row_data = shortterm_strategy_sheet.get_all_values()

            for data in row_data:
                if data[0] == '단위투자금':
                    self.Stocklist['전략']['단위투자금'] = int(data[1])
                elif data[0] == '매수모니터링 종료시간':
                    self.Stocklist['전략']['모니터링종료시간'] = data[1] + ':00'
                elif data[0] == '보유일':
                    self.Stocklist['전략']['보유일'] = int(data[1])
                elif data[0] == '투자금 비중':
                    self.Stocklist['전략']['투자금비중'] = float(data[1][:-1])
                # elif data[0] == '손절율':
                #     self.Stocklist['전략']['매도구간별조건'].append(float(data[1][:-1]))
                # elif data[0] == '시가 위치':
                #     self.Stocklist['전략']['시가위치'] = list(map(int, data[1].split(',')))
                elif '구간' in data[0]:
                    if data[0][-1] != '1' and data[0][-1] != '2':
                        self.Stocklist['전략']['매도구간별조건'].append(float(data[1][:-1]))
                elif '손절가' == data[0]:
                    self.Stocklist['전략']['전략매도가'].append(float(data[1].replace('%', '')))
                elif '본전가' == data[0]:
                    self.Stocklist['전략']['전략매도가'].append(float(data[1].replace('%', '')))
                elif '익절가' in data[0]:
                    self.Stocklist['전략']['전략매도가'].append(float(data[1].replace('%', '')))

            self.Stocklist['전략']['매도구간별조건'].insert(0, self.Stocklist['전략']['전략매도가'][0])  # 손절가
            self.Stocklist['전략']['매도구간별조건'].insert(1, self.Stocklist['전략']['전략매도가'][1])  # 본전가

            for code in self.Stocklist.keys():
                if code == '컬럼명' or code == '전략':
                    continue
                else:
                    self.Stocklist[code]['단위투자금'] = int(
                        self.Stocklist[code]['투자비중'] * self.Stocklist['전략']['단위투자금'])
                    self.Stocklist[code]['시가체크'] = False
                    self.Stocklist[code]['매수상한도달'] = False
                    self.Stocklist[code]['매수조건'] = 0
                    self.Stocklist[code]['매수총수량'] = 0  # 분할매수에 따른 수량체크
                    self.Stocklist[code]['매수수량'] = 0  # 분할매수 단위
                    self.Stocklist[code]['매수주문완료'] = 0  # 분할매수에 따른 매수 주문 수
                    self.Stocklist[code]['매수가전략'] = len(self.Stocklist[code]['매수가'])  # 매수 전략에 따른 매수가 지정 수량
                    if self.Stocklist[code]['매도전략'] == '4':
                        self.Stocklist[code]['매도가'].append(self.Stocklist['전략']['전략매도가'])
            print(self.Stocklist)

        except Exception as e:
            print('CTradeShortTerm_Setting Error :', e)
            Telegram('[XTrader]CTradeShortTerm_Setting Error : %s' % e, send='mc')
            logger.error('CTradeShortTerm_Setting Error : %s' % e)

    # 수동 포트폴리오 생성
    def manual_portfolio(self):
        self.portfolio = dict()
        self.Stocklist = {
            '097800': {'번호': '7.099', '종목명': '윈팩', '종목코드': '097800', '시장': 'KOSDAQ', '매수전략': '10', '매수가': [3219],
                       '매수조건': 1, '수량': 310, '매도전략': '4', '매도가': [3700], '매수일': '2020/05/29 09:22:39'},

            '297090': {'번호': '7.101', '종목명': '씨에스베어링', '종목코드': '297090', '시장': 'KOSDAQ', '매수전략': '10', '매수가': [5000],
                       '매수조건': 3, '수량': 15, '매도전략': '2', '매도가': [], '매수일': '2020/06/03 09:12:15'},

            '053610': {'번호': '6.154', '종목명': '프로텍', '종목코드': '053610', '시장': 'KOSDAQ', '매수전략': '10', '매수가': [9500],
                       '매수조건': 2, '수량': 26, '매도전략': '4', '매도가': [9900], '매수일': '2020/06/03 09:12:15'}
        }

        self.strategy = {'전략': {'단위투자금': 200000, '모니터링종료시간': '10:30:00', '보유일': 20,
                                '투자금비중': 70.0, '매도구간별조건': [-2.7, 0.3, -3.0, -4.0, -5.0, -7.0],
                                '전략매도가': [-2.7, 0.3, 3.0, 6.0]}}

        for code in list(self.Stocklist.keys()):
            self.portfolio[code] = CPortStock_ShortTerm(번호=self.Stocklist[code]['번호'], 종목코드=code,
                                                        종목명=self.Stocklist[code]['종목명'],
                                                        시장=self.Stocklist[code]['시장'],
                                                        매수가=self.Stocklist[code]['매수가'][0],
                                                        매수조건=self.Stocklist[code]['매수조건'],
                                                        보유일=self.strategy['전략']['보유일'],
                                                        매도전략=self.Stocklist[code]['매도전략'],
                                                        매도가=self.Stocklist[code]['매도가'],
                                                        매도구간별조건=self.strategy['전략']['매도구간별조건'], 매도구간=1,
                                                        수량=self.Stocklist[code]['수량'],
                                                        매수일=self.Stocklist[code]['매수일'])

    # google spreadsheet 매매이력 생성
    def save_history(self, code, status):
        # 매매이력 sheet에 해당 종목(매수된 종목)이 있으면 row를 반환 아니면 예외처리 -> 신규 매수로 처리
        if status == '매도모니터링':
            row = []
            row.append(self.portfolio[code].번호)
            row.append(self.portfolio[code].종목명)
            row.append(self.portfolio[code].매수가)

            shortterm_sell_sheet.append_row(row)

        try:
            code_row = shortterm_history_sheet.findall(self.portfolio[code].종목명)[
                -1].row  # 종목명이 있는 모든 셀을 찾아서 맨 아래에 있는 셀을 선택
            cell = alpha_list[shortterm_history_cols.index('매도가')] + str(code_row)  # 매수 이력에 있는 종목이 매도가 되었는지 확인
            sell_price = shortterm_history_sheet.acell(str(cell)).value

            # 매도 이력은 추가 매도(매도전략2의 경우)나 신규 매도인 경우라 매도 이력 유무와 상관없음
            if status == '매도':  # 매도 이력은 포트폴리오에서 종목 pop을 하므로 Stocklist 데이터 사용
                cell = alpha_list[shortterm_history_cols.index('매도가')] + str(code_row)
                shortterm_history_sheet.update_acell(cell, self.Stocklist[code]['매도체결가'])

                cell = alpha_list[shortterm_history_cols.index('매도수량')] + str(code_row)
                shortterm_history_sheet.update_acell(cell, self.Stocklist[code]['매도수량'])

                cell = alpha_list[shortterm_history_cols.index('매도일')] + str(code_row)
                shortterm_history_sheet.update_acell(cell, datetime.datetime.now().strftime('%Y/%m/%d %H:%M:%S'))

                cell = alpha_list[shortterm_history_cols.index('매도전략')] + str(code_row)
                shortterm_history_sheet.update_acell(cell, self.Stocklist[code]['매도전략'])

                cell = alpha_list[shortterm_history_cols.index('매도구간')] + str(code_row)
                shortterm_history_sheet.update_acell(cell, self.Stocklist[code]['매도구간'])

                계산수익률 = round((self.Stocklist[code]['매도체결가'] / self.Stocklist[code]['매수가'] - 1) * 100, 2)
                cell = alpha_list[shortterm_history_cols.index('수익률(계산)')] + str(code_row)  # 수익률 계산
                shortterm_history_sheet.update_acell(cell, 계산수익률)

            # 매수 이력은 있으나 매도 이력이 없음 -> 매도 전 추가 매수
            if sell_price == '':
                if status == '매수':  # 포트폴리오 데이터 사용
                    cell = alpha_list[shortterm_history_cols.index('매수가')] + str(code_row)
                    shortterm_history_sheet.update_acell(cell, self.portfolio[code].매수가)

                    cell = alpha_list[shortterm_history_cols.index('매수수량')] + str(code_row)
                    shortterm_history_sheet.update_acell(cell, self.portfolio[code].수량)

                    cell = alpha_list[shortterm_history_cols.index('매수일')] + str(code_row)
                    shortterm_history_sheet.update_acell(cell, self.portfolio[code].매수일)

                    cell = alpha_list[shortterm_history_cols.index('매수조건')] + str(code_row)
                    shortterm_history_sheet.update_acell(cell, self.portfolio[code].매수조건)

            else:  # 매도가가 기록되어 거래가  완료된 종목으로 판단하여 예외발생으로 신규 매수 추가함
                raise Exception('매매완료 종목')

        except Exception as e:
            try:
                logger.debug('CTradeShortTerm_save_history Error1 : 종목명:%s, %s' % (self.portfolio[code].종목명, e))
                row = []
                row_buy = []
                if status == '매수':
                    print('%s 매수이력 row 내용 생성' % self.portfolio[code].종목명)
                    row.append(self.portfolio[code].번호)
                    row.append(self.portfolio[code].종목명)
                    row.append(self.portfolio[code].매수가)
                    row.append(self.portfolio[code].수량)
                    row.append(self.portfolio[code].매수일)
                    row.append(self.portfolio[code].매수조건)

                print('%s 매수이력 row 내용 생성완료' % self.portfolio[code].종목명)
                shortterm_history_sheet.append_row(row)
            except Exception as e:
                print('CTradeShortTerm_save_history Error2 : 종목명:%s, %s' % (self.portfolio[code].종목명, e))
                Telegram('[XTrade]CTradeShortTerm_save_history Error2 : 종목명:%s, %s' % (self.portfolio[code].종목명, e),
                         send='mc')
                logger.debug('CTradeShortTerm_save_history Error2 : 종목명:%s, %s' % (self.portfolio[code].종목명, e))

    # 매수 전략별 매수 조건 확인
    def buy_strategy(self, code, price):
        result = False
        condition = self.Stocklist[code]['매수조건']  # 초기값 0
        qty = self.Stocklist[code]['매수수량']  # 초기값 0

        현재가, 시가, 고가, 저가, 전일종가 = price  # 시세 = [현재가, 시가, 고가, 저가, 전일종가]

        매수가 = self.Stocklist[code]['매수가']  # [매수가1, 매수가2, 매수가3]
        시가위치하한 = self.Stocklist[code]['시가위치'][0]
        시가위치상한 = self.Stocklist[code]['시가위치'][1]

        # 1. 금일시가 위치 체크(초기 한번)하여 매수조건(1~6)과 주문 수량 계산
        if self.Stocklist[code]['시가체크'] == False:  # 종목별로 초기에 한번만 시가 위치 체크를 하면 되므로 별도 함수 미사용
            매수가.append(시가)
            매수가.sort(reverse=True)
            band = 매수가.index(시가)  # band = 0 : 매수가1 이상, band=1: 매수가1, 2 사이, band=2: 매수가2,3 사이
            매수가.remove(시가)

            if band == len(매수가):  # 매수가 지정한 구간보다 시가가 아래일 경우로 초기값이 result=False, condition=0 리턴
                self.Stocklist[code]['시가체크'] = True
                self.Stocklist[code]['매수조건'] = 0
                self.Stocklist[code]['매수수량'] = 0
                return False, 0, 0
            else:
                # 단위투자금으로 매수가능한 총 수량 계산, band = 0 : 매수가1, band=1: 매수가2, band=2: 매수가3 로 계산
                self.Stocklist[code]['매수총수량'] = self.Stocklist[code]['단위투자금'] // 매수가[band]
                if band == 0:  # 시가가 매수가1보다 높은 경우
                    # 시가가 매수가1의 시가범위에 포함 : 조건 1, 2, 3
                    if 매수가[band] * (1 + 시가위치하한 / 100) <= 시가 and 시가 < 매수가[band] * (1 + 시가위치상한 / 100):
                        condition = len(매수가)
                        self.Stocklist[code]['매수가전략'] = len(매수가)
                        qty = self.Stocklist[code]['매수총수량'] // condition
                    else:  # 시가 위치에 미포함
                        self.Stocklist[code]['시가체크'] = True
                        self.Stocklist[code]['매수조건'] = 0
                        self.Stocklist[code]['매수수량'] = 0
                        return False, 0, 0
                else:  # 시가가 매수가 중간인 경우 - 매수가1&2사이(band 1) : 조건 4,5 / 매수가2&3사이(band 2) : 조건 6
                    for i in range(band):  # band 1일 경우 매수가 1은 불필요하여 삭제, band 2 : 매수가 1, 2 삭제(band수 만큼 삭제 실행)
                        매수가.pop(0)
                    if 매수가[0] * (1 + 시가위치하한 / 100) <= 시가:  # 시가범위 포함
                        # 조건 4 = 매수가길이 1 + band 1 + 2(=band+1) -> 4 = 1 + 2*1 + 1
                        # 조건 5 = 매수가길이 2 + band 1 + 2(=band+1) -> 5 = 2 + 2*1 + 1
                        # 조건 6 = 매수가길이 1 + band 2 + 3(=band+1) -> 6 = 1 + 2*2 + 1
                        condition = len(매수가) + (2 * band) + 1
                        self.Stocklist[code]['매수가전략'] = len(매수가)
                        qty = self.Stocklist[code]['매수총수량'] // (condition % 2 + 1)
                    else:
                        self.Stocklist[code]['시가체크'] = True
                        self.Stocklist[code]['매수조건'] = 0
                        self.Stocklist[code]['매수수량'] = 0
                        return False, 0, 0

            self.Stocklist[code]['시가체크'] = True
            self.Stocklist[code]['매수조건'] = condition
            self.Stocklist[code]['매수수량'] = qty
        else:  # 시가 위치 체크를 한 두번째 데이터 이후에는 condition이 0이면 바로 매수 불만족 리턴시킴
            if condition == 0:  # condition 0은 매수 조건 불만족
                return False, 0, 0

        # 매수조건 확정, 매수 수량 계산 완료
        # 매수상한에 미도달한 상태로 매수가로 내려왔을 때 매수
        # 현재가가 해당조건에서의 시가위치 상한 이상으로 오르면 매수상한도달을 True로 해서 매수하지 않게 함
        if 현재가 >= 매수가[0] * (1 + 시가위치상한 / 100): self.Stocklist[code]['매수상한도달'] = True

        if self.Stocklist[code]['매수주문완료'] < self.Stocklist[code]['매수가전략'] and self.Stocklist[code]['매수상한도달'] == False:
            if 현재가 == 매수가[0]:
                result = True

                self.Stocklist[code]['매수주문완료'] += 1

                print("매수모니터링 만족_종목:%s, 시가:%s, 조건:%s, 현재가:%s, 체크결과:%s, 수량:%s" % (
                    self.Stocklist[code]['종목명'], 시가, condition, 현재가, result, qty))
                logger.debug("매수모니터링 만족_종목:%s, 시가:%s, 조건:%s, 현재가:%s, 체크결과:%s, 수량:%s" % (
                    self.Stocklist[code]['종목명'], 시가, condition, 현재가, result, qty))

        return result, condition, qty

    # 매도 구간 확인
    def profit_band_check(self, 현재가, 매수가):
        band_list = [0, 3, 5, 10, 15, 25]
        # print('현재가, 매수가', 현재가, 매수가)

        ratio = round((현재가 - 매수가) / 매수가 * 100, 2)
        # print('ratio', ratio)

        if ratio < 3:
            return 1
        elif ratio in band_list:
            return band_list.index(ratio) + 1
        else:
            band_list.append(ratio)
            band_list.sort()
            band = band_list.index(ratio)
            band_list.remove(ratio)
            return band

    # 매도 전략별 매도 조건 확인
    def sell_strategy(self, code, price):
        # print('%s 매도 조건 확인' % code)
        try:
            result = False
            band = self.portfolio[code].매도구간  # 이전 매도 구간 받음
            매도방법 = self.매도방법  # '03' : 시장가
            qty_ratio = 1  # 매도 수량 결정 : 보유수량 * qty_ratio

            현재가, 시가, 고가, 저가, 전일종가 = price  # 시세 = [현재가, 시가, 고가, 저가, 전일종가]
            매수가 = self.portfolio[code].매수가

            # 전략 1, 2, 3과 4 별도 체크
            strategy = self.portfolio[code].매도전략
            # 전략 1, 2, 3
            if strategy != '4':
                # 매도를 위한 수익률 구간 체크(매수가 대비 현재가의 수익률 조건에 다른 구간 설정)
                new_band = self.profit_band_check(현재가, 매수가)
                if (hogacal(시가, 0, self.portfolio[code].시장, '상한가')) <= 현재가:
                    band = 7

                if band < new_band:  # 이전 구간보다 현재 구간이 높을 경우(시세가 올라간 경우)만
                    band = new_band  # 구간을 현재 구간으로 변경(반대의 경우는 구간 유지)

                if band == 1 and 현재가 <= 매수가 * (1 + (self.portfolio[code].매도구간별조건[0] / 100)):
                    result = True
                elif band == 2 and 현재가 <= 매수가 * (1 + (self.portfolio[code].매도구간별조건[1] / 100)):
                    result = True
                elif band == 3 and 현재가 <= 고가 * (1 + (self.portfolio[code].매도구간별조건[2] / 100)):
                    result = True
                elif band == 4 and 현재가 <= 고가 * (1 + (self.portfolio[code].매도구간별조건[3] / 100)):
                    result = True
                elif band == 5 and 현재가 <= 고가 * (1 + (self.portfolio[code].매도구간별조건[4] / 100)):
                    result = True
                elif band == 6 and 현재가 <= 고가 * (1 + (self.portfolio[code].매도구간별조건[5] / 100)):
                    result = True
                elif band == 7 and 현재가 >= (hogacal(시가, -3, self.Stocklist[code]['시장'], '상한가')):
                    매도방법 = '00'  # 지정가
                    result = True

                self.portfolio[code].매도구간 = band  # 포트폴리오에 매도구간 업데이트

                try:
                    if strategy == '2' or strategy == '3':  # 매도전략 2(기존 5)
                        if strategy == '2':
                            목표가 = self.portfolio[code].매도가[0]
                        elif strategy == '3':
                            목표가 = (hogacal(시가 * 1.1, 0, self.Stocklist[code]['시장'], '현재가'))
                        매도조건 = self.portfolio[code].매도조건  # 매도가 실행된 조건 '': 매도 전, 'B':구간매도, 'T':목표가매도

                        target_band = self.profit_band_check(목표가, 매수가)

                        if band < target_band:  # 현재가구간이 목표가구간 미만일때 전량매도
                            qty_ratio = 1

                        else:  # 현재가구간이 목표가구간 이상일 때
                            if 현재가 == 목표가:  # 목표가 도달 시 절반 매도
                                self.portfolio[code].목표도달 = True  # 목표가 도달 여부 True
                                if 매도조건 == '':  # 매도이력이 없는 경우 목표가매도 'T', 절반 매도
                                    self.portfolio[code].매도조건 = 'T'
                                    result = True
                                    if self.portfolio[code].수량 == 1:
                                        qty_ratio = 1
                                    else:
                                        qty_ratio = 0.5
                                elif 매도조건 == 'B':  # 구간 매도 이력이 있을 경우 절반매도가 된 상태이므로 남은 전량매도
                                    result = True
                                    qty_ratio = 1
                                elif 매도조건 == 'T':  # 목표가 매도 이력이 있을 경우 매도미실행
                                    result = False

                            else:  # 현재가가 목표가가 아닐 경우 구간 매도 실행(매도실행여부는 결정된 상태)
                                if self.portfolio[code].목표도달 == False:  # 목표가 도달을 못한 경우면 전량매도
                                    qty_ratio = 1
                                else:
                                    if 매도조건 == '':  # 매도이력이 없는 경우 구간매도 'B', 절반 매도
                                        self.portfolio[code].매도조건 = 'B'
                                        if self.portfolio[code].수량 == 1:
                                            qty_ratio = 1
                                        else:
                                            qty_ratio = 0.5
                                    elif 매도조건 == 'B':  # 구간 매도 이력이 있을 경우 매도미실행
                                        result = False
                                    elif 매도조건 == 'T':  # 목표가 매도 이력이 있을 경우 전량매도
                                        qty_ratio = 1

                except Exception as e:
                    print('sell_strategy 매도전략 2 Error :', e)
                    logger.error('CTradeShortTerm_sell_strategy 종목 : %s 매도전략 2 Error : %s' % (code, e))
                    Telegram('[XTrader]CTradeShortTerm_sell_strategy 종목 : %s 매도전략 2 Error : %s' % (code, e), send='mc')
                    result = False
                    return 매도방법, result, qty_ratio

                # print('종목코드 : %s, 현재가 : %s, 시가 : %s, 고가 : %s, 매도구간 : %s, 결과 : %s' % (code, 현재가, 시가, 고가, band, result))
                return 매도방법, result, qty_ratio

            # 전략 4(지정가 00 매도)
            else:
                매도방법 = '00'  # 지정가
                try:
                    # 전략 4의 매도가 = [목표가(원), [손절가(%), 본전가(%), 1차익절가(%), 2차익절가(%)]]
                    # 1. 매수 후 손절가까지 하락시 매도주문 -> 손절가, 전량매도로 끝
                    if 현재가 <= 매수가 * (1 + self.portfolio[code].매도가[1][0] / 100):
                        self.portfolio[code].매도구간 = 0
                        result = True
                        qty_ratio = 1
                    # 2. 1차익절가 도달시 매도주문 -> 1차익절가, 1/3 매도
                    elif self.portfolio[code].익절가1도달 == False and 현재가 >= 매수가 * (
                            1 + self.portfolio[code].매도가[1][2] / 100):
                        self.portfolio[code].매도구간 = 1
                        self.portfolio[code].익절가1도달 = True
                        result = True
                        if self.portfolio[code].수량 == 1:
                            qty_ratio = 1
                        elif self.portfolio[code].수량 == 2:
                            qty_ratio = 0.5
                        else:
                            qty_ratio = 0.3
                    # 3. 2차익절가 도달못하고 본전가까지 하락 또는 고가 -3%까지시 매도주문 -> 1차익절가, 나머지 전량 매도로 끝
                    elif self.portfolio[code].익절가1도달 == True and self.portfolio[code].익절가2도달 == False and (
                            (현재가 <= 매수가 * (1 + self.portfolio[code].매도가[1][1] / 100)) or (현재가 <= 고가 * 0.97)):
                        self.portfolio[code].매도구간 = 1.5
                        result = True
                        qty_ratio = 1
                    # 4. 2차 익절가 도달 시 매도주문 -> 2차 익절가, 1/3 매도
                    elif self.portfolio[code].익절가1도달 == True and self.portfolio[code].익절가2도달 == False and 현재가 >= 매수가 * (
                            1 + self.portfolio[code].매도가[1][3] / 100):
                        self.portfolio[code].매도구간 = 2
                        self.portfolio[code].익절가2도달 = True
                        result = True
                        if self.portfolio[code].수량 == 1:
                            qty_ratio = 1
                        else:
                            qty_ratio = 0.5
                    # 5. 목표가 도달못하고 2차익절가까지 하락 시 매도주문 -> 2차익절가, 나머지 전량 매도로 끝
                    elif self.portfolio[code].익절가2도달 == True and self.portfolio[code].목표가도달 == False and (
                            (현재가 <= 매수가 * (1 + self.portfolio[code].매도가[1][2] / 100)) or (현재가 <= 고가 * 0.97)):
                        self.portfolio[code].매도구간 = 2.5
                        result = True
                        qty_ratio = 1
                    # 6. 목표가 도달 시 매도주문 -> 목표가, 나머지 전량 매도로 끝
                    elif self.portfolio[code].목표가도달 == False and 현재가 >= self.portfolio[code].매도가[0]:
                        self.portfolio[code].매도구간 = 3
                        self.portfolio[code].목표가도달 = True
                        result = True
                        qty_ratio = 1

                    return 매도방법, result, qty_ratio

                except Exception as e:
                    print('sell_strategy 매도전략 4 Error :', e)
                    logger.error('CTradeShortTerm_sell_strategy 종목 : %s 매도전략 4 Error : %s' % (code, e))
                    Telegram('[XTrader]CTradeShortTerm_sell_strategy 종목 : %s 매도전략 4 Error : %s' % (code, e), send='mc')
                    result = False
                    return 매도방법, result, qty_ratio

        except Exception as e:
            print('CTradeShortTerm_sell_strategy Error ', e)
            Telegram('[XTrader]CTradeShortTerm_sell_strategy Error : %s' % e, send='mc')
            logger.error('CTradeShortTerm_sell_strategy Error : %s' % e)
            result = False
            qty_ratio = 1
            return 매도방법, result, qty_ratio

    # 보유일 전략 : 보유기간이 보유일 이상일 경우 전량 매도 실행(Mainwindow 타이머에서 시간 체크)
    def hold_strategy(self):
        if self.holdcheck == True:
            print('보유일 만기 매도 체크')
            try:
                for code in list(self.portfolio.keys()):
                    보유기간 = holdingcal(self.portfolio[code].매수일)
                    print('종목명 : %s, 보유일 : %s, 보유기간 : %s' % (self.portfolio[code].종목명, self.portfolio[code].보유일, 보유기간))
                    if 보유기간 >= int(self.portfolio[code].보유일) and self.주문실행중_Lock.get('S_%s' % code) is None and \
                            self.portfolio[code].수량 != 0:
                        self.portfolio[code].매도구간 = 0
                        (result, order) = self.정량매도(sRQName='S_%s' % code, 종목코드=code, 매도가=self.portfolio[code].매수가,
                                                    수량=self.portfolio[code].수량)

                        if result == True:
                            self.주문실행중_Lock['S_%s' % code] = True
                            Telegram('[XTrader]정량매도(보유일만기) : 종목코드=%s, 종목명=%s, 수량=%s' % (
                                code, self.portfolio[code].종목명, self.portfolio[code].수량))
                            logger.info('정량매도(보유일만기) : 종목코드=%s, 종목명=%s, 수량=%s' % (
                                code, self.portfolio[code].종목명, self.portfolio[code].수량))
                        else:
                            Telegram('[XTrader]정액매도실패(보유일만기) : 종목코드=%s, 종목명=%s, 수량=%s' % (
                                code, self.portfolio[code].종목명, self.portfolio[code].수량))
                            logger.info('정량매도실패(보유일만기) : 종목코드=%s, 종목명=%s, 수량=%s' % (
                                code, self.portfolio[code].종목명, self.portfolio[code].수량))
            except Exception as e:
                print("hold_strategy Error :", e)

    # 포트폴리오 생성
    def set_portfolio(self, code, buyprice, condition):
        try:
            self.portfolio[code] = CPortStock_ShortTerm(번호=self.Stocklist[code]['번호'], 종목코드=code,
                                                        종목명=self.Stocklist[code]['종목명'],
                                                        시장=self.Stocklist[code]['시장'], 매수가=buyprice,
                                                        매수조건=condition, 보유일=self.Stocklist['전략']['보유일'],
                                                        매도전략=self.Stocklist[code]['매도전략'],
                                                        매도가=self.Stocklist[code]['매도가'],
                                                        매도구간별조건=self.Stocklist['전략']['매도구간별조건'],
                                                        매수일=datetime.datetime.now().strftime('%Y/%m/%d %H:%M:%S'))

            self.Stocklist[code]['매수일'] = datetime.datetime.now().strftime('%Y/%m/%d %H:%M:%S')  # 매매이력 업데이트를 위해 매수일 추가

        except Exception as e:
            print('CTradeShortTerm_set_portfolio Error ', e)
            Telegram('[XTrader]CTradeShortTerm_set_portfolio Error : %s' % e, send='mc')
            logger.error('CTradeShortTerm_set_portfolio Error : %s' % e)

    # Robot_Run이 되면 실행됨 - 매수/매도 종목을 리스트로 저장
    def 초기조건(self, codes):
        # 매수총액 계산하기
        # 금일매도종목 리스트 변수 초기화
        # 매도할종목 : 포트폴리오에 있던 종목 추가
        # 매수할종목 : 구글에서 받은 종목 추가
        self.parent.statusbar.showMessage("[%s] 초기조건준비" % (self.sName))

        self.금일매도종목 = []  # 장 마감 후 금일 매도한 종목에 대해서 매매이력 정리 업데이트(매도가, 손익률 등)
        self.매도할종목 = []
        self.매수할종목 = []
        self.매수총액 = 0
        self.holdcheck = False

        for code in codes:  # 구글 시트에서 import된 매수 모니커링 종목은 '매수할종목'에 추가
            self.매수할종목.append(code)

        # 포트폴리오에 있는 종목은 매도 관련 전략 재확인(구글시트) 및 '매도할종목'에 추가
        if len(self.portfolio) > 0:
            row_data = shortterm_sell_sheet.get_all_values()
            idx_holding = row_data[0].index('보유일')
            idx_strategy = row_data[0].index('매도전략')
            idx_loss = row_data[0].index('손절가')
            idx_sellprice = row_data[0].index('목표가')
            for row in row_data[1:]:
                code, name, market = get_code(row[1])  # 종목명으로 종목코드, 종목명, 시장 받아서(get_code 함수) 추가
                if code in list(self.portfolio.keys()):
                    self.portfolio[code].보유일 = row[idx_holding]
                    self.portfolio[code].매도전략 = row[idx_strategy]
                    self.portfolio[code].매도가 = []  # 매도 전략 변경에 따라 매도가 초기화
                    if self.portfolio[code].매도전략 == '4':  # 매도가 = [목표가(원), [손절가(%), 본전가(%), 1차익절가(%), 2차익절가(%)]]
                        self.portfolio[code].매도가.append(int(float(row[idx_sellprice].replace(',', ''))))
                        self.portfolio[code].매도가.append(self.Stocklist['전략']['전략매도가'])
                        self.portfolio[code].매도가[1][0] = float(row[idx_loss].replace('%', ''))

                        self.portfolio[code].sellcount = 0
                        self.portfolio[code].매도단위수량 = 0  # 전략4의 기본 매도 단위는 보유수량의 1/3
                        self.portfolio[code].익절가1도달 = False
                        self.portfolio[code].익절가2도달 = False
                        self.portfolio[code].목표가도달 = False
                    else:  # 매도구간별조건 = [손절가(%), 본전가(%), 구간3 고가대비(%), 구간4 고가대비(%), 구간5 고가대비(%), 구간6 고가대비(%)]
                        self.portfolio[code].매도구간별조건[0] = float(row[idx_loss][:-1])
                        if self.portfolio[code].매도전략 == '2' or self.portfolio[code].매도전략 == '3':
                            self.portfolio[code].목표도달 = False  # 목표가(매도가) 도달 체크(False 상태로 구간 컷일경우 전량 매도)
                            self.portfolio[code].매도조건 = ''  # 구간매도 : B, 목표매도 : T

        for port_code in list(self.portfolio.keys()):
            # 로봇 시작 시 포트폴리오 종목의 매도구간(전일 매도모니터링)을 1로 초기화
            # 구간이 내려가는 건 반영하지 않으므로 초기화를 시켜서 다시 구간 체크 시작하기 위함
            self.portfolio[port_code].매도구간 = 1  # 매도 구간은 로봇 실행 시 마다 초기화시킴

            # 매수총액계산
            self.매수총액 += (self.portfolio[port_code].매수가 * self.portfolio[port_code].수량)

            # 포트폴리오에 있는 종목이 구글에서 받아서 만든 Stocklist에 없을 경우만 추가함
            # 이 조건이 없을 경우 구글에서 받은 전략들이 아닌 과거 전략이 포트폴리오에서 넘어감
            # 근데 포트폴리오에 있는 종목을 왜 Stocklist에 넣어야되는지 모르겠음(내가 하고도...)
            if port_code not in list(self.Stocklist.keys()):
                self.Stocklist[port_code] = {
                    '번호': self.portfolio[port_code].번호,
                    '종목명': self.portfolio[port_code].종목명,
                    '종목코드': self.portfolio[port_code].종목코드,
                    '시장': self.portfolio[port_code].시장,
                    '매수조건': self.portfolio[port_code].매수조건,
                    '매수가': self.portfolio[port_code].매수가,
                    '매도전략': self.portfolio[port_code].매도전략,
                    '매도가': self.portfolio[port_code].매도가
                }

            self.매도할종목.append(port_code)

        # for stock in df_keeplist['종목번호'].values: # 보유 종목 체크해서 매도 종목에 추가 → 로봇이 두개 이상일 경우 중복되므로 미적용
        #     self.매도할종목.append(stock)
        # 종목명 = df_keeplist[df_keeplist['종목번호']==stock]['종목명'].values[0]
        # 매입가 = df_keeplist[df_keeplist['종목번호']==stock]['매입가'].values[0]
        # 보유수량 = df_keeplist[df_keeplist['종목번호']==stock]['보유수량'].values[0]
        # print('종목코드 : %s, 종목명 : %s, 매입가 : %s, 보유수량 : %s' %(stock, 종목명, 매입가, 보유수량))
        # self.portfolio[stock] = CPortStock_ShortTerm(종목코드=stock, 종목명=종목명, 매수가=매입가, 수량=보유수량, 매수일='')

    def 실시간데이터처리(self, param):
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

                종목명 = self.parent.CODE_POOL[종목코드][1]  # pool[종목코드] = [시장구분, 종목명, 주식수, 전일종가, 시가총액]
                전일종가 = self.parent.CODE_POOL[종목코드][3]
                시세 = [현재가, 시가, 고가, 저가, 전일종가]

                self.parent.statusbar.showMessage("[%s] %s %s %s %s" % (체결시간, 종목코드, 종목명, 현재가, 전일대비))
                self.wr.writerow([체결시간, 종목코드, 종목명, 현재가, 전일대비])

                # 매수 조건
                # 매수모니터링 종료 시간 확인
                if current_time < self.Stocklist['전략']['모니터링종료시간']:
                    if 종목코드 in self.매수할종목 and 종목코드 not in self.금일매도종목:
                        # 매수총액 + 종목단위투자금이 투자총액보다 작음 and 매수주문실행중Lock에 없음 -> 추가매수를 위해서 and 포트폴리오에 없음 조건 삭제
                        if (self.매수총액 + self.Stocklist[종목코드]['단위투자금'] < self.투자총액) and self.주문실행중_Lock.get(
                                'B_%s' % 종목코드) is None and len(
                            self.Stocklist[종목코드]['매수가']) > 0:  # and self.portfolio.get(종목코드) is None
                            # 매수 전략별 모니터링 체크
                            buy_check, condition, qty = self.buy_strategy(종목코드, 시세)
                            if buy_check == True and (self.Stocklist[종목코드]['단위투자금'] // 현재가 > 0):
                                (result, order) = self.정량매수(sRQName='B_%s' % 종목코드, 종목코드=종목코드, 매수가=현재가, 수량=qty)

                                if result == True:
                                    if self.portfolio.get(종목코드) is None:  # 포트폴리오에 없으면 신규 저장
                                        self.set_portfolio(종목코드, 현재가, condition)

                                    self.주문실행중_Lock['B_%s' % 종목코드] = True
                                    Telegram('[XTrader]매수주문 : 종목코드=%s, 종목명=%s, 매수가=%s, 매수조건=%s, 매수수량=%s' % (
                                        종목코드, 종목명, 현재가, condition, qty))
                                    logger.info('매수주문 : 종목코드=%s, 종목명=%s, 매수가=%s, 매수조건=%s, 매수수량=%s' % (
                                        종목코드, 종목명, 현재가, condition, qty))

                                else:
                                    Telegram('[XTrader]매수실패 : 종목코드=%s, 종목명=%s, 매수가=%s, 매수조건=%s' % (
                                        종목코드, 종목명, 현재가, condition))
                                    logger.info('매수실패 : 종목코드=%s, 종목명=%s, 매수가=%s, 매수조건=%s' % (종목코드, 종목명, 현재가, condition))
                else:
                    if self.매수모니터링체크 == False:
                        for code in self.매수할종목:
                            if self.portfolio.get(code) is not None and code not in self.매도할종목:
                                Telegram('[XTrader]매수모니터링마감 : 종목코드=%s, 종목명=%s 매도모니터링 전환' % (종목코드, 종목명))
                                logger.info('매수모니터링마감 : 종목코드=%s, 종목명=%s 매도모니터링 전환' % (종목코드, 종목명))
                                self.매수할종목.remove(code)
                                self.매도할종목.append(code)

                        self.매수모니터링체크 = True
                        logger.info('매도할 종목 :%s' % self.매도할종목)

                # 매도 조건
                if 종목코드 in self.매도할종목:
                    # 포트폴리오에 있음 and 매도주문실행중Lock에 없음 and 매수주문실행중Lock에 없음
                    if self.portfolio.get(종목코드) is not None and self.주문실행중_Lock.get(
                            'S_%s' % 종목코드) is None:  # and self.주문실행중_Lock.get('B_%s' % 종목코드) is None:
                        # 매도 전략별 모니터링 체크
                        매도방법, sell_check, ratio = self.sell_strategy(종목코드, 시세)
                        if sell_check == True:
                            if 매도방법 == '00':
                                (result, order) = self.정액매도(sRQName='S_%s' % 종목코드, 종목코드=종목코드, 매도가=현재가,
                                                            수량=round(self.portfolio[종목코드].수량 * ratio))
                            else:
                                (result, order) = self.정량매도(sRQName='S_%s' % 종목코드, 종목코드=종목코드, 매도가=현재가,
                                                            수량=round(self.portfolio[종목코드].수량 * ratio))

                            if result == True:
                                self.주문실행중_Lock['S_%s' % 종목코드] = True
                                Telegram('[XTrader]매도주문 : 종목코드=%s, 종목명=%s, 매도가=%s, 매도전략=%s, 매도구간=%s, 수량=%s' % (
                                    종목코드, 종목명, 현재가, self.portfolio[종목코드].매도전략, self.portfolio[종목코드].매도구간,
                                    int(self.portfolio[종목코드].수량 * ratio)))
                                if self.portfolio[종목코드].매도전략 == '2':
                                    logger.info(
                                        '매도주문 : 종목코드=%s, 종목명=%s, 매도가=%s, 매도전략=%s, 매도구간=%s, 목표도달=%s, 매도조건=%s, 수량=%s' % (
                                            종목코드, 종목명, 현재가, self.portfolio[종목코드].매도전략, self.portfolio[종목코드].매도구간,
                                            self.portfolio[종목코드].목표도달, self.portfolio[종목코드].매도조건,
                                            int(self.portfolio[종목코드].수량 * ratio)))
                                else:
                                    logger.info('매도주문 : 종목코드=%s, 종목명=%s, 매도가=%s, 매도전략=%s, 매도구간=%s, 수량=%s' % (
                                        종목코드, 종목명, 현재가, self.portfolio[종목코드].매도전략, self.portfolio[종목코드].매도구간,
                                        int(self.portfolio[종목코드].수량 * ratio)))
                            else:
                                Telegram(
                                    '[XTrader]매도실패 : 종목코드=%s, 종목명=%s, 매도가=%s, 매도전략=%s, 매도구간=%s, 수량=%s' % (종목코드, 종목명,
                                                                                                          현재가,
                                                                                                          self.portfolio[
                                                                                                              종목코드].매도전략,
                                                                                                          self.portfolio[
                                                                                                              종목코드].매도구간,
                                                                                                          self.portfolio[
                                                                                                              종목코드].수량 * ratio))
                                logger.info('매도실패 : 종목코드=%s, 종목명=%s, 매도가=%s, 매도전략=%s, 매도구간=%s, 수량=%s' % (종목코드, 종목명,
                                                                                                         현재가,
                                                                                                         self.portfolio[
                                                                                                             종목코드].매도전략,
                                                                                                         self.portfolio[
                                                                                                             종목코드].매도구간,
                                                                                                         self.portfolio[
                                                                                                             종목코드].수량 * ratio))


        except Exception as e:
            print('CTradeShortTerm_실시간데이터처리 Error : %s, %s' % (종목명, e))
            Telegram('[XTrader]CTradeShortTerm_실시간데이터처리 Error : %s, %s' % (종목명, e), send='mc')
            logger.error('CTradeShortTerm_실시간데이터처리 Error :%s, %s' % (종목명, e))

    def 접수처리(self, param):
        pass

    def 체결처리(self, param):
        종목코드 = param['종목코드']
        주문번호 = param['주문번호']
        self.주문결과[주문번호] = param

        주문수량 = int(param['주문수량'])
        미체결수량 = int(param['미체결수량'])
        체결가 = int(0 if (param['체결가'] is None or param['체결가'] == '') else param['체결가'])  # 매입가 동일
        단위체결량 = int(0 if (param['단위체결량'] is None or param['단위체결량'] == '') else param['단위체결량'])
        당일매매수수료 = int(0 if (param['당일매매수수료'] is None or param['당일매매수수료'] == '') else param['당일매매수수료'])
        당일매매세금 = int(0 if (param['당일매매세금'] is None or param['당일매매세금'] == '') else param['당일매매세금'])

        # 매수
        if param['매도수구분'] == '2':
            if self.주문번호_주문_매핑.get(주문번호) is not None:
                주문 = self.주문번호_주문_매핑[주문번호]
                매수가 = int(주문[2:])
                # 단위체결가 = int(0 if (param['단위체결가'] is None or param['단위체결가'] == '') else param['단위체결가'])

                # logger.debug('매수-------> %s %s %s %s %s' % (param['종목코드'], param['종목명'], 매수가, 주문수량 - 미체결수량, 미체결수량))

                P = self.portfolio.get(종목코드)
                if P is not None:
                    P.종목명 = param['종목명']
                    P.매수가 = 체결가  # 단위체결가
                    P.수량 += 단위체결량  # 추가 매수 대비해서 기존 수량에 체결된 수량 계속 더함(주문수량 - 미체결수량)

                    P.매수일 = datetime.datetime.now().strftime('%Y/%m/%d %H:%M:%S')
                else:
                    logger.error('ERROR 포트에 종목이 없음 !!!!')

                if 미체결수량 == 0:
                    try:
                        self.주문실행중_Lock.pop(주문)
                        if self.Stocklist[종목코드]['매수주문완료'] >= self.Stocklist[종목코드]['매수가전략']:
                            self.매수할종목.remove(종목코드)
                            self.매도할종목.append(종목코드)

                            Telegram('[XTrader]분할 매수 완료_종목명:%s, 종목코드:%s 매수가:%s, 수량:%s' % (P.종목명, 종목코드, P.매수가, P.수량))
                            logger.info('분할 매수 완료_종목명:%s, 종목코드:%s 매수가:%s, 수량:%s' % (P.종목명, 종목코드, P.매수가, P.수량))

                        self.Stocklist[종목코드]['수량'] = P.수량
                        self.Stocklist[종목코드]['매수가'].pop(0)

                        self.매수총액 += (P.매수가 * P.수량)
                        logger.debug('체결처리완료_종목명:%s, 매수총액계산완료:%s' % (P.종목명, self.매수총액))

                        self.save_history(종목코드, status='매수')
                        Telegram('[XTrader]매수체결완료_종목명:%s, 매수가:%s, 수량:%s' % (P.종목명, P.매수가, P.수량))
                        logger.info('매수체결완료_종목명:%s, 매수가:%s, 수량:%s' % (P.종목명, P.매수가, P.수량))
                    except Exception as e:
                        Telegram('[XTrader]체결처리_매수 에러 종목명:%s, %s ' % (P.종목명, e), send='mc')
                        logger.error('체결처리_매수 에러 종목명:%s, %s ' % (P.종목명, e))

    def 잔고처리(self, param):
        # print('CTradeShortTerm : 잔고처리')

        종목코드 = param['종목코드']
        P = self.portfolio.get(종목코드)
        if P is not None:
            P.매수가 = int(0 if (param['매입단가'] is None or param['매입단가'] == '') else param['매입단가'])
            P.수량 = int(0 if (param['보유수량'] is None or param['보유수량'] == '') else param['보유수량'])
            if P.수량 == 0:
                self.portfolio.pop(종목코드)
                self.매도할종목.remove(종목코드)
                if 종목코드 not in self.금일매도종목: self.금일매도종목.append(종목코드)
                logger.info('잔고처리_포트폴리오POP %s ' % 종목코드)

        # 메인 화면에 반영
        self.parent.RobotView()

    def Run(self, flag=True, sAccount=None):
        self.running = flag
        ret = 0
        # self.manual_portfolio()

        for code in list(self.portfolio.keys()):
            print(self.portfolio[code].__dict__)
            logger.info(self.portfolio[code].__dict__)
            # if code == '051490': self.portfolio.pop(code)
        if flag == True:
            print("%s ROBOT 실행" % (self.sName))
            try:
                Telegram("[XTrader]%s ROBOT 실행" % (self.sName))

                self.sAccount = sAccount

                self.투자총액 = floor(int(d2deposit.replace(",", "")) * (self.Stocklist['전략']['투자금비중'] / 100))

                print('로봇거래계좌 : ', 로봇거래계좌번호)
                print('D+2 예수금 : ', int(d2deposit.replace(",", "")))
                print('투자 총액 : ', self.투자총액)
                print('Stocklist : ', self.Stocklist)

                # self.최대포트수 = floor(int(d2deposit.replace(",", "")) / self.단위투자금 / len(self.parent.robots))
                # print(self.최대포트수)

                self.주문결과 = dict()
                self.주문번호_주문_매핑 = dict()
                self.주문실행중_Lock = dict()

                codes = list(self.Stocklist.keys())
                codes.remove('전략')
                codes.remove('컬럼명')
                self.초기조건(codes)

                print("매도 : ", self.매도할종목)
                print("매수 : ", self.매수할종목)
                print("매수총액 : ", self.매수총액)

                print("포트폴리오 매도모니터링 수정")
                for code in list(self.portfolio.keys()):
                    print('종목명 : %s, 보유일 : %s, 매도전략 : %s, 매도가 : %s' % (
                        self.portfolio[code].종목명, self.portfolio[code].보유일, self.portfolio[code].매도전략,
                        self.portfolio[code].매도가))

                self.실시간종목리스트 = self.매도할종목 + self.매수할종목

                logger.info("오늘 거래 종목 : %s %s" % (self.sName, ';'.join(self.실시간종목리스트) + ';'))
                self.KiwoomConnect()  # MainWindow 외에서 키움 API구동시켜서 자체적으로 API데이터송수신가능하도록 함
                if len(self.실시간종목리스트) > 0:
                    self.f = open('data_result.csv', 'a', newline='')
                    self.wr = csv.writer(self.f)
                    self.wr.writerow(['체결시간', '종목코드', '종목명', '현재가', '전일대비'])

                    ret = self.KiwoomSetRealReg(self.sScreenNo, ';'.join(self.실시간종목리스트) + ';')
                    logger.debug("실시간데이타요청 등록결과 %s" % ret)

            except Exception as e:
                print('CTradeShortTerm_Run Error :', e)
                Telegram('[XTrader]CTradeShortTerm_Run Error : %s' % e, send='mc')
                logger.error('CTradeShortTerm_Run Error : %s' % e)

        else:
            Telegram("[XTrader]%s ROBOT 실행 중지" % (self.sName))
            print('Stocklist : ', self.Stocklist)

            ret = self.KiwoomSetRealRemove(self.sScreenNo, 'ALL')

            self.f.close()
            del self.f
            del self.wr

            if self.portfolio is not None:
                for code in list(self.portfolio.keys()):
                    if self.portfolio[code].수량 == 0:
                        self.portfolio.pop(code)

            if len(self.매도할종목) > 0:
                # 매도모니터링 시트 기존 자료 삭제
                num_data = shortterm_sell_sheet.get_all_values()
                for i in range(len(num_data)):
                    shortterm_sell_sheet.delete_rows(2)

                for code in self.매도할종목:
                    self.save_history(code, status='매도모니터링')

            if len(self.금일매도종목) > 0:
                try:
                    Telegram("[XTrader]%s 금일 매도 종목 손익 Upload : %s" % (self.sName, self.금일매도종목))
                    logger.info("%s 금일 매도 종목 손익 Upload : %s" % (self.sName, self.금일매도종목))
                    self.parent.statusbar.showMessage("금일 매도 종목 손익 Upload")
                    self.DailyProfit(self.금일매도종목)
                except Exception as e:
                    print('%s 금일매도종목 결과 업로드 Error : %s' % (self.sName, e))
                finally:
                    del self.DailyProfitLoop  # 금일매도결과 업데이트 시 QEventLoop 사용으로 로봇 저장 시 pickcle 에러 발생하여 삭제시킴

            self.KiwoomDisConnect()  # 로봇 클래스 내에서 일별종목별실현손익 데이터를 받고나서 연결 해제시킴

            # 메인 화면에 반영
            self.parent.RobotView()


##################################################################################
# 메인
##################################################################################
Ui_MainWindow, QtBaseClass_MainWindow = uic.loadUiType("./UI/XTrader_MainWindow.ui")
class MainWindow(QMainWindow, Ui_MainWindow):
    def __init__(self):
        # 화면을 보여주기 위한 코드
        super().__init__()
        QMainWindow.__init__(self)
        Ui_MainWindow.__init__(self)

        self.UI_setting()

        # 현재 시간 받음
        self.시작시각 = datetime.datetime.now()

        # 메인윈도우가 뜨고 키움증권과 붙이기 위한 작업
        self.KiwoomAPI()  # 키움 ActiveX를 메모리에 올림
        self.KiwoomConnect()  # 메모리에 올라온 ActiveX와 내가 만든 함수 On시리즈와 연결(콜백 : 이벤트가 오면 나를 불러줘)
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

        # self.portfolio_model.update((DataFrame(columns=['종목코드', '종목명', '매수가', '수량', '매수일'])))

        self.robot_columns = ['Robot타입', 'Robot명', 'RobotID', '스크린번호', '실행상태', '포트수', '포트폴리오']

        # TODO: 주문제한 설정
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.limit_per_second)  # 초당 4번
        # QtCore.QObject.connect(self.timer, QtCore.SIGNAL("timeout()"), self.limit_per_second)
        self.timer.start(1000)  # 1초마다 리셋

        self.주문제한 = 0
        self.조회제한 = 0
        self.금일백업작업중 = False
        self.종목선정작업중 = False

        self.DailyData = False  # 관심종목 일봉 업데이트
        self.InvestorData = False  # 관심종목 종목별투자자 업데이트

        self.df_daily = DataFrame()
        self.df_weekly = DataFrame()
        self.df_monthly = DataFrame()
        self.df_investor = DataFrame()                           
        self._login = False

        self.KiwoomLogin()  # 프로그램 실행 시 자동로그인
        self.CODE_POOL = self.get_code_pool()  # DB 종목데이블에서 시장구분, 코드, 종목명, 주식수, 전일종가 읽어옴

    # 화면 Setting
    def UI_setting(self):
        self.setupUi(self)
        self.setWindowTitle("XTrader")
        self.setWindowIcon(QIcon('./PNG/icon_stock.png'))

        self.actionLogin.setIcon(QIcon('./PNG/Internal.png'))
        self.actionLogout.setIcon(QIcon('./PNG/External.png'))
        self.actionExit.setIcon(QIcon('./PNG/Approval.png'))

        self.actionAccountDialog.setIcon(QIcon('./PNG/Sales Performance.png'))
        self.actionMinutePrice.setIcon(QIcon('./PNG/Candle Sticks.png'))
        self.actionDailyPrice.setIcon(QIcon('./PNG/Overtime.png'))
        self.actionInvestors.setIcon(QIcon('./PNG/Conference Call.png'))
        self.actionSectorView.setIcon(QIcon('./PNG/Organization.png'))
        self.actionSectorPriceView.setIcon(QIcon('./PNG/Ratings.png'))
        self.actionCodeBuild.setIcon(QIcon('./PNG/Inspection.png'))

        self.actionRobotOneRun.setIcon(QIcon('./PNG/Process.png'))
        self.actionRobotOneStop.setIcon(QIcon('./PNG/Cancel 2.png'))
        self.actionRobotMonitoringStop.setIcon(QIcon('./PNG/Cancel File.png'))
        self.actionRobotRun.setIcon(QIcon('./PNG/Checked.png'))
        self.actionRobotStop.setIcon(QIcon('./PNG/Cancel.png'))
        self.actionRobotRemove.setIcon(QIcon('./PNG/Delete File.png'))
        self.actionRobotClear.setIcon(QIcon('./PNG/Empty Trash.png'))
        self.actionRobotView.setIcon(QIcon('./PNG/Checked 2.png'))
        self.actionRobotSave.setIcon(QIcon('./PNG/Download.png'))

        self.actionTradeShortTerm.setIcon(QIcon('./PNG/Bullish.png'))

    # DB에 저장된 상장 종목 코드 읽음
    def get_code_pool(self):
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
    def Import_ShortTermStock(self, check):
        try:
            data = import_googlesheet()

            if check == False:
                # # 매수 전략별 별도 로봇 운영 시
                # # 매수 전략 확인
                # strategy_list = list(data['매수전략'].unique())
                #
                # # 로딩된 로봇을 robot_list에 저장
                # robot_list = []
                # for robot in self.robots:
                #     robot_list.append(robot.sName.split('_')[0])
                #
                # # 매수 전략별 로봇 자동 편집/추가
                # for strategy in strategy_list:
                #     df_stock = data[data['매수전략'] == strategy]
                #
                #     if strategy in robot_list:
                #         print('로봇 편집')
                #         Telegram('[XTrader]로봇 편집')
                #         for robot in self.robots:
                #             if robot.sName.split('_')[0] == strategy:
                #                 self.RobotAutoEdit_TradeShortTerm(robot, df_stock)
                #                 self.RobotView()
                #                 break
                #     else:
                #         print('로봇 추가')
                #         Telegram('[XTrader]로봇 추가')
                #         self.RobotAutoAdd_TradeShortTerm(df_stock, strategy)
                #         self.RobotView()

                # 로딩된 로봇을 robot_list에 저장
                robot_list = []
                for robot in self.robots:
                    robot_list.append(robot.sName)

                if 'TradeShortTerm' in robot_list:
                    for robot in self.robots:
                        if robot.sName == 'TradeShortTerm':
                            print('로봇 편집')
                            logger.debug('로봇 편집')
                            self.RobotAutoEdit_TradeShortTerm(robot, data)
                            self.RobotView()
                            break

                else:
                    print('로봇 추가')
                    logger.debug('로봇 추가')
                    self.RobotAutoAdd_TradeShortTerm(data)
                    self.RobotView()

                # print("로봇 준비 완료")
                # Slack('[XTrader]로봇 준비 완료')
                # logger.info("로봇 준비 완료")

        except Exception as e:
            print('MainWindow_Import_ShortTermStock Error', e)
            Telegram('[XTrader]MainWindow_Import_ShortTermStock Error : %s' % e, send='mc')
            logger.error('MainWindow_Import_ShortTermStock Error : %s' % e)

    # 프로그램 실행 3초 후 실행
    def OnQApplicationStarted(self):
        # 1. 8시 58분 이전일 경우 5분 단위 구글시트 오퓨 체크 타이머 시작시킴
        current = datetime.datetime.now()
        current_time = current.strftime('%H:%M:%S')
        if '07:00:00' <= current_time and current_time <= '08:58:00':
            print('구글 시트 오류 체크 시작')
            # Telegram('[XTrader]구글 시트 오류 체크 시작')
            self.statusbar.showMessage("구글 시트 오류 체크 시작")

            self.checkclock = QTimer(self)
            self.checkclock.timeout.connect(self.OnGoogleCheck)  # 5분마다 구글 시트 읽음 : MainWindow.OnGoogleCheck 실행
            self.checkclock.start(300000)  # 300000초마다 타이머 작동

        # 2. DB에 저장된 로봇 정보받아옴
        global 로봇거래계좌번호
        try:
            with sqlite3.connect(DATABASE) as conn:
                cursor = conn.cursor()

                cursor.execute("select value from Setting where keyword='robotaccount'")

                for row in cursor.fetchall():
                    # _temp = base64.decodestring(row[0])  # base64에 text화해서 암호화 : DB에 잘 넣기 위함
                    _temp = base64.decodebytes(row[0])
                    로봇거래계좌번호 = pickle.loads(_temp)
                    print('로봇거래계좌번호', 로봇거래계좌번호)

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

        # 8시 32분 : 종목 데이블 생성
        if current_time == '08:32:00':
            print('종목테이블 생성')
            # Slack('[XTrader]종목테이블 생성')
            self.StockCodeBuild(to_db=True)
            self.CODE_POOL = self.get_code_pool()  # DB 종목데이블에서 시장구분, 코드, 종목명, 주식수, 전일종가 읽어옴
            self.statusbar.showMessage("종목테이블 생성")

        # 8시 59분 : 구글 시트 종목 Import
        if current_time == '08:59:00':
            print('구글 시트 오류 체크 중지')
            # Telegram('[XTrader]구글 시트 오류 체크 중지')
            self.checkclock.stop()

            robot_list = []
            for robot in self.robots:
                robot_list.append(robot.sName)

            if 'TradeShortTerm' in robot_list:
                print('구글시트 Import')
                Telegram('[XTrader]구글시트 Import')
                self.Import_ShortTermStock(check=False)

                self.statusbar.showMessage('구글시트 Import')

        # 8시 59분 30초 : 로봇 실행
        if '08:59:30' <= current_time and current_time < '08:59:40':
            try:
                if len(self.robots) > 0:
                    for r in self.robots:
                        if r.running == False:  # 로봇이 실행중이 아니면
                            r.Run(flag=True, sAccount=로봇거래계좌번호)
                            self.RobotView()
            except Exception as e:
                print('Robot Auto Run Error', e)
                Telegram('[XTrader]Robot Auto Run Error : %s' % e, send='mc')
                logger.error('Robot Auto Run Error : %s' % e)

        # 15시 29분 :TradeShortTerm 보유일 만기 매도 전략 체크 후 주문
        if current_time >= '15:29:00' and current_time < '15:29:30':
            if len(self.robots) > 0:
                for r in self.robots:
                    if r.sName == 'TradeShortTerm':
                        if r.holdcheck == False:
                            r.holdcheck = True
                            r.hold_strategy()

        # 16시 00분 : 로봇 정지
        if '16:00:00' <= current_time and current_time < '16:00:05':
            self.RobotStop()

        # 16시 05분 : 프로그램 종료
        if '16:05:00' <= current_time and current_time < '16:05:05':
            Telegram("[XTrade]프로그램을 종료합니다.")
            quit()

    # 주문 제한 초기화
    def limit_per_second(self):
        self.주문제한 = 0
        self.조회제한 = 0
        # logger.info("초당제한 주문 클리어")

    # 5분 마다 실행 : 구글 스프레드 시트 오류 확인
    def OnGoogleCheck(self):
        self.Import_ShortTermStock(check=True)

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
        elif _action == "actionMinutePrice":
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
        elif _action == "actionAccountDialog":  # 계좌정보조회
            if self.dialog.get('계좌정보조회') is not None:  # dialog : __init__()에 dict로 정의됨
                try:
                    self.dialog['계좌정보조회'].show()
                except Exception as e:
                    self.dialog['계좌정보조회'] = 화면_계좌정보(sScreenNo=7000, kiwoom=self.kiwoom,
                                                    parent=self)  # self는 메인윈도우, 계좌정보윈도우는 자식윈도우/부모는 메인윈도우
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
        elif _action == "actionTradeShortTerm":
            self.RobotAdd_TradeShortTerm()
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
            self.RobotOneMonitoringStop()
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
        elif _action == "actionTest":
            import_googlesheet()

    # 키움증권 OpenAPI
    # 키움API ActiveX를 메모리에 올림
    def KiwoomAPI(self):
        self.kiwoom = QAxWidget("KHOPENAPI.KHOpenAPICtrl.1")

    # 메모리에 올라온 ActiveX와 On시리즈와 붙임(콜백 : 이벤트가 오면 나를 불러줘)
    def KiwoomConnect(self):
        self.kiwoom.OnEventConnect[int].connect(
            self.OnEventConnect)  # 키움의 OnEventConnect와 이 프로그램의 OnEventConnect 함수와 연결시킴
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
        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, Qstring)', "계좌번호", self.sAccount)
        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, Qstring)', "비밀번호입력매체구분", '00')
        ret = self.kiwoom.dynamicCall('SetInputValue(Qstring, Qstring)', "조회구분", '1')
        ret = self.kiwoom.dynamicCall('CommRqData(QString, QString, int, QString)', "계좌평가잔고내역요청", "opw00018",
                                      _repeat, '{:04d}'.format(self.ScreenNumber))

        self.InquiryLoop = QEventLoop()  # 로봇에서 바로 쓸 수 있도록하기 위해서 계좌 조회해서 종목을 받고나서 루프해제시킴
        self.InquiryLoop.exec_()

    # 계좌 번호 / D+2 예수금 받음
    def KiwoomAccount(self):
        ACCOUNT_CNT = self.kiwoom.dynamicCall('GetLoginInfo("ACCOUNT_CNT")')
        ACC_NO = self.kiwoom.dynamicCall('GetLoginInfo("ACCNO")')
        self.account = ACC_NO.split(';')[0:-1]
        self.sAccount = self.account[0]

        global Account
        Account = self.sAccount

        global 로봇거래계좌번호
        로봇거래계좌번호 = self.sAccount
        print('계좌 : ', self.sAccount)
        print('로봇계좌 : ', 로봇거래계좌번호)
        self.kiwoom.dynamicCall('SetInputValue(Qstring, Qstring)', "계좌번호", self.sAccount)
        self.kiwoom.dynamicCall('CommRqData(QString, QString, int, QString)', "d+2예수금요청", "opw00001", 0,
                                '{:04d}'.format(self.ScreenNumber))
        self.depositLoop = QEventLoop()  # self.d2_deposit를 로봇에서 바로 쓸 수 있도록하기 위해서 예수금을 받고나서 루프해제시킴
        self.depositLoop.exec_()

        # return (ACCOUNT_CNT, ACC_NO)

    def KiwoomSendOrder(self, sRQName, sScreenNo, sAccNo, nOrderType, sCode, nQty, nPrice, sHogaGb, sOrgOrderNo):
        if self.주문제한 < 초당횟수제한:
            Order = self.kiwoom.dynamicCall(
                'SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)',
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
        ret = self.kiwoom.dynamicCall('SetRealReg(QString, QString, QString, QString)', sScreenNo, sCode, '9001;10',
                                      sRealType)  # 10은 실시간FID로 메뉴얼에 나옴(현재가,체결가, 실시간종가)
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

            current = datetime.datetime.now().strftime('%H:%M:%S')
            if current <= '08:58:00':
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

    def OnReceiveTrData(self, sScrNo, sRQName, sTRCode, sRecordName, sPreNext, nDataLength, sErrorCode, sMessage, sSPlmMsg):
        # logger.debug('main:OnReceiveTrData [%s] [%s] [%s] [%s] [%s] [%s] [%s] [%s] [%s] ' % (sScrNo, sRQName, sTRCode, sRecordName, sPreNext, nDataLength, sErrorCode, sMessage, sSPlmMsg))
        # print("MainWindow : OnReceiveTrData")

        if self.ScreenNumber != int(sScrNo):
            return

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

        if sRQName == "주식일봉차트조회":
            try:
                self.주식일봉컬럼 = ['일자', '현재가', '거래량']  # ['일자', '현재가', '시가', '고가', '저가', '거래량',  '거래대금']
                # cnt = self.kiwoom.dynamicCall('GetRepeatCnt(QString, QString)', sTRCode, sRQName)
                cnt = self.AnalysisPriceList[3] + 30
                for i in range(0, cnt):
                    row = []
                    for j in self.주식일봉컬럼:
                        S = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "",
                                                    sRQName, i, j).strip().lstrip('0')
                        if len(S) > 0 and S[0] == '-':
                            S = '-' + S[1:].lstrip('0')
                        # if S == '': S = 0
                        # if j != '일자':S = int(float(S))
                        row.append(S)
                        # print(row)
                    self.종목일봉.append(row)

                df = DataFrame(data=self.종목일봉, columns=self.주식일봉컬럼)
                # df.to_csv('data.csv')

                try:
                    df.loc[df.현재가 == '', ['현재가']] = 0
                    df.loc[df.거래량 == '', ['거래량']] = 0
                except:
                    pass

                df = df.sort_values(by='일자').reset_index(drop=True)
                # df.to_csv('data.csv')

                self.UploadAnalysisData(data=df, 구분='일봉')

                if len(self.종목리스트) > 0:
                    self.종목코드 = self.종목리스트.pop(0)
                    QTimer.singleShot(주문지연, lambda: self.ReguestPriceDaily(_repeat=0))
                else:
                    print('일봉데이터 수신 완료')
                    self.DailyData = False
                    self.WeeklyData = True
                    self.MonthlyData = False
                    self.InvestorData = False
                    self.stock_analysis()

            except Exception as e:
                print('OnReceiveTrData_주식일봉차트조회 : ', self.종목코드, e)

        if sRQName == "주식주봉차트조회":
            try:
                self.주식주봉컬럼 = ['일자', '현재가']  # ['일자', '현재가', '시가', '고가', '저가', '거래량',  '거래대금']
                # cnt = self.kiwoom.dynamicCall('GetRepeatCnt(QString, QString)', sTRCode, sRQName)
                cnt = self.AnalysisPriceList[4]+5
                for i in range(0, cnt):
                    row = []
                    for j in self.주식주봉컬럼:
                        S = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "",
                                                    sRQName, i, j).strip().lstrip('0')
                        if len(S) > 0 and S[0] == '-':
                            S = '-' + S[1:].lstrip('0')
                        # if S == '': S = 0
                        # if j != '일자':S = int(float(S))
                        row.append(S)
                        # print(row)
                    self.종목주봉.append(row)

                df = DataFrame(data=self.종목주봉, columns=self.주식주봉컬럼)
                # df.to_csv('data.csv')

                try:
                    df.loc[df.현재가 == '', ['현재가']] = 0
                except:
                    pass

                df = df.sort_values(by='일자').reset_index(drop=True)
                # df.to_csv('data.csv')

                self.UploadAnalysisData(data=df, 구분='주봉')

                if len(self.종목리스트) > 0:
                    self.종목코드 = self.종목리스트.pop(0)
                    QTimer.singleShot(주문지연, lambda: self.ReguestPriceWeekly(_repeat=0))
                else:
                    print('주봉데이터 수신 완료')
                    self.DailyData = False
                    self.WeeklyData = False
                    self.MonthlyData = True
                    self.InvestorData = False
                    self.stock_analysis()

            except Exception as e:
                print('OnReceiveTrData_주식주봉차트조회 : ', self.종목코드, e)

        if sRQName == "주식월봉차트조회":
            try:
                self.주식월봉컬럼 = ['일자', '현재가']  # ['일자', '현재가', '시가', '고가', '저가', '거래량',  '거래대금']
                # cnt = self.kiwoom.dynamicCall('GetRepeatCnt(QString, QString)', sTRCode, sRQName)
                cnt = self.AnalysisPriceList[5]+5
                for i in range(0, cnt):
                    row = []
                    for j in self.주식월봉컬럼:
                        S = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "",
                                                    sRQName, i, j).strip().lstrip('0')
                        if len(S) > 0 and S[0] == '-':
                            S = '-' + S[1:].lstrip('0')
                        # if S == '': S = 0
                        # if j != '일자':S = int(float(S))
                        row.append(S)
                        # print(row)
                    self.종목월봉.append(row)

                df = DataFrame(data=self.종목월봉, columns=self.주식월봉컬럼)
                try:
                    df.loc[df.현재가 == '', ['현재가']] = 0
                except:
                    pass

                df = df.sort_values(by='일자').reset_index(drop=True)
                #df.to_csv('data.csv')

                self.UploadAnalysisData(data=df, 구분='월봉')

                if len(self.종목리스트) > 0:
                    self.종목코드 = self.종목리스트.pop(0)
                    QTimer.singleShot(주문지연, lambda: self.ReguestPriceMonthly(_repeat=0))
                else:
                    print('월봉데이터 수신 완료')
                    self.DailyData = False
                    self.WeeklyData = False
                    self.MonthlyData = False
                    self.InvestorData = True
                    self.stock_analysis()

            except Exception as e:
                print('OnReceiveTrData_주식월봉차트조회 : ', self.종목코드, e)

        if sRQName == "종목별투자자조회":
            self.종목별투자자컬럼 = ['일자', '기관계', '외국인투자자', '개인투자자']
            # ['일자', '현재가', '전일대비', '누적거래대금', '개인투자자', '외국인투자자', '기관계', '금융투자', '보험', '투신', '기타금융', '은행','연기금등', '국가', '내외국인', '사모펀드', '기타법인']
            try:
                # cnt = self.kiwoom.dynamicCall('GetRepeatCnt(QString, QString)', sTRCode, sRQName)
                cnt = 10
                for i in range(0, cnt):
                    row = []
                    for j in self.종목별투자자컬럼:
                        S = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "",
                                                    sRQName, i, j).strip().lstrip('0').replace('--', '-')
                        if S == '': S = '0'
                        row.append(S)
                    self.종목별투자자.append(row)

                df = DataFrame(data=self.종목별투자자, columns=self.종목별투자자컬럼)
                df['일자'] = df['일자'].apply(lambda x: x[0:4] + '-' + x[4:6] + '-' + x[6:])
                try:
                    df.ix[df.개인투자자 == '', ['개인투자자']] = 0
                    df.ix[df.외국인투자자 == '', ['외국인투자자']] = 0
                    df.ix[df.기관계 == '', ['기관계']] = 0
                except:
                    pass
                # df.dropna(inplace=True)
                df = df.sort_values(by='일자').reset_index(drop=True)
                #df.to_csv('종목별투자자.csv', encoding='euc-kr')

                self.UploadAnalysisData(data=df, 구분='종목별투자자')

                if len(self.종목리스트) > 0:
                    self.종목코드 = self.종목리스트.pop(0)
                    QTimer.singleShot(주문지연, lambda: self.RequestInvestorDaily(_repeat=0))
                else:
                    print('종목별투자자데이터 수신 완료')
                    self.end = datetime.datetime.now()
                    print('start :', self.start)
                    print('end :', self.end)
                    print('소요시간 :', self.end - self.start)
                    self.df_analysis = pd.merge(self.df_daily, self.df_weekly, on='종목코드', how='outer')
                    self.df_analysis = pd.merge(self.df_analysis, self.df_monthly, on='종목코드', how='outer')
                    self.df_analysis = pd.merge(self.df_analysis, self.df_investor, on='종목코드', how='outer')
                    self.df_analysis['우선순위'] = ''
                    self.df_analysis = self.df_analysis[
                        ['번호', '종목명', '우선순위', '일봉1', '일봉2', '일봉3', '일봉4', '주봉1', '월봉1', '거래량', '기관', '외인', '개인']]
                    print(self.df_analysis.head())
                    conn = sqliteconn()
                    self.df_analysis.to_sql('종목분석', conn, if_exists='replace', index=False)
                    query = """
                                select *
                                from 종목분석
                            """
                    df = pd.read_sql(query, con=conn)
                    df.to_csv('종목분석.csv', index=False, encoding='euc-kr')
                    conn.close()
                    Telegram('[XTrader]관심종목 데이터 업데이트 완료', send='mc')

            except Exception as e:
                print('OnReceiveTrData_종목별투자자조회 : ', self.종목코드, e)

        if sRQName == "d+2예수금요청":
            data = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName,
                                           0, "d+2추정예수금")

            # 입력된 문자열에 대해 lstrip 메서드를 통해 문자열 왼쪽에 존재하는 '-' 또는 '0'을 제거. 그리고 format 함수를 통해 천의 자리마다 콤마를 추가한 문자열로 변경
            strip_data = data.lstrip('-0')
            if strip_data == '':
                strip_data = '0'

            format_data = format(int(strip_data), ',d')
            if data.startswith('-'):
                format_data = '-' + format_data

            global d2deposit  # D+2 예수금 → 매수 가능 금액 계산을 위함
            d2deposit = format_data
            print("예수금 %s 저장 완료" % (d2deposit))
            self.depositLoop.exit()  # self.d2_deposit를 로봇에서 바로 쓸 수 있도록하기 위해서 예수금을 받고나서 루프해제시킴

        if sRQName == "계좌평가잔고내역요청":
            try:
                cnt = self.kiwoom.dynamicCall('GetRepeatCnt(QString, QString)', sTRCode, sRQName)

                global df_keeplist  # 계좌 보유 종목 리스트

                result = []

                cols = ['종목번호', '종목명', '보유수량', '매입가', '매입금액']  # , '평가금액', '수익률(%)', '평가손익', '매매가능수량']
                for i in range(0, cnt):
                    row = []
                    for j in cols:
                        # S = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "", sRQName, i, '종목번호').strip().lstrip('0')
                        S = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "",
                                                    sRQName, i, j).strip().lstrip('0')

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

        if sRQName == "일자별종목별실현손익요청":
            try:
                data_idx = ['종목명', '체결량', '매입단가', '체결가', '당일매도손익', '손익율', '당일매매수수료', '당일매매세금']

                result = []
                for idx in data_idx:
                    data = self.kiwoom.dynamicCall('CommGetData(QString, QString, QString, int, QString)', sTRCode, "",
                                                   sRQName, 0, idx)
                    result.append(data.strip())

                self.DailyProfitUpload(result)

            except Exception as e:
                print(e)
                logger.error('일자별종목별실현손익요청 Error : %s' % e)

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

    # Robot 함수
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

    def RobotRun(self):
        for r in self.robots:
            # r.초기조건()
            # logger.debug('%s %s %s %s' % (r.sName, r.UUID, len(r.portfolio), r.GetStatus()))
            print('RobotRun_로봇거래계좌번호 : ', 로봇거래계좌번호)
            r.Run(flag=True, sAccount=로봇거래계좌번호)

        self.statusbar.showMessage("RUN !!!")

    def RobotStop(self):
        try:
            for r in self.robots:
                if r.running == True:
                    r.Run(flag=False)
                    logger.info("전체 ROBOT 실행 중지시킵니다.")                                                                                                                                     

            self.RobotView()
            self.RobotSaveSilently()


            self.statusbar.showMessage("전체 ROBOT 실행 중지!!!")

        except Exception as e:
            print("Robot all stop error", e)
            Telegram('[XTrade]Robot all stop error : %s' % e)
            logger.error('Robot all stop error : %s' % e)

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

        # reply = QMessageBox.question(self,
        #                              "로봇 실행 중지", "로봇 실행을 중지할까요?\n%s" % robot_found.GetStatus(),
        #                              QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel)
        # if reply == QMessageBox.Cancel:
        #     pass
        # elif reply == QMessageBox.No:
        #     pass
        # elif reply == QMessageBox.Yes:
        try:
            if robot_found.running == True:
                robot_found.Run(flag=False)
                for code in list(robot_found.portfolio.keys()):
                    if robot_found.portfolio[code].수량 == 0:
                        robot_found.portfolio.pop(code)

            self.RobotView()
            self.RobotSaveSilently()
        except Exception as e:
            print("Robot one stop error", e)
            logger.error('Robot one stop error : %s' % e)

    def RobotOneMonitoringStop(self):
        print('RobotMonitoringStop')
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

        if robot_found.running == True:
            print('Robot_%s : 매수 모니터링 정지' % robot_found.sName)
            print(robot_found.매수할종목)
            robot_found.매수할종목 = []
            logger.info('Robot_%s : 매수 모니터링 정지' % robot_found.sName)
            Telegram('[XTrade]Robot_%s : 매수 모니터링 정지' % robot_found.sName)

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
        if len(self.robots) > 0:
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
            reply = QMessageBox.question(self,
                                         "Robot Save Error", "현재 설정된 로봇이 없습니다. DB를 삭제할까요?",
                                         QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel)
            if reply == QMessageBox.Cancel:
                pass
            elif reply == QMessageBox.No:
                pass
            elif reply == QMessageBox.Yes:
                try:
                    with sqlite3.connect(DATABASE) as conn:
                        cursor = conn.cursor()
                        cursor.execute('delete from Robots')
                        conn.commit()
                except Exception as e:
                    print('RobotSaveSilently', e)
                finally:
                    self.statusbar.showMessage("로봇 저장 완료")

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

                    uuid = r.UUID
                    strategy = r.__class__.__name__
                    name = r.sName
                    print('로봇 저장 : ', r.sName)

                    robot = pickle.dumps(r, protocol=pickle.HIGHEST_PROTOCOL, fix_imports=True)

                    # robot_encoded = base64.encodestring(robot)
                    robot_encoded = base64.encodebytes(robot)

                    cursor.execute("insert or replace into Robots(uuid, strategy, name, robot) values (?, ?, ?, ?)",
                                   [uuid, strategy, name, robot_encoded])
                    conn.commit()

        except Exception as e:
            print('RobotSaveSilently_Error : ', e)
            logger.error('RobotSaveSilently_Error : %s' % (e))

        finally:
            r.kiwoom = self.kiwoom
            r.parent = self
            print('로봇 저장 완료')
            self.statusbar.showMessage("로봇 저장 완료")

    def RobotView(self):
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

            if Robot타입 == 'CTradeShortTerm':
                self.RobotEdit_TradeShortTerm(robot_found)
        except Exception as e:
            print('RobotEdit', e)

    def RobotSelected(self, QModelIndex):
        # print(self.model._data[QModelIndex.row()])
        try:
            RobotName = self.model._data[QModelIndex.row():QModelIndex.row() + 1]['Robot명'].values[0]

            uuid = self.model._data[QModelIndex.row():QModelIndex.row() + 1]['RobotID'].values[0]
            portfolio = None
            for r in self.robots:
                if r.UUID == uuid:
                    portfolio = r.portfolio
                    # print(portfolio.items())

                    model = PandasModel()
                    result = []
                    if RobotName == 'TradeShortTerm':
                        self.portfolio_columns = ['번호', '종목코드', '종목명', '매수가', '수량', '매수조건', '매도전략', '보유일', '매수일']
                        for p, v in portfolio.items():
                            result.append((v.번호, v.종목코드, v.종목명.strip(), v.매수가, v.수량, v.매수조건, v.매도전략,  v.보유일, v.매수일))
                        self.portfolio_model.update((DataFrame(data=result, columns=self.portfolio_columns)))
                    elif RobotName == 'TradeCondition':
                        self.portfolio_columns = ['종목코드', '종목명', '매수가', '수량', '매수일']
                        for p, v in portfolio.items():
                            result.append((v.종목코드, v.종목명.strip(), v.매수가, v.수량, v.매수일))
                        self.portfolio_model.update((DataFrame(data=result, columns=self.portfolio_columns)))
                    break

        except Exception as e:
            print('robot_selected', e)

    def RobotDoubleClicked(self, QModelIndex):
        self.RobotEdit(QModelIndex)
        self.RobotView()

    def RobotCurrentIndex(self, index):
        self.tableView_robot_current_index = index

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

    # TradeShotTerm
    def RobotAdd_TradeShortTerm(self):
        # print("MainWindow : RobotAdd_TradeShortTerm")
        try:
            스크린번호 = self.GetUnAssignedScreenNumber()
            R = 화면_TradeShortTerm(parent=self)
            R.lineEdit_screen_number.setText('{:04d}'.format(스크린번호))
            if R.exec_():
                매수방법 = R.comboBox_buy_condition.currentText().strip()[0:2]
                매도방법 = R.comboBox_sell_condition.currentText().strip()[0:2]

                # strategy_list = list(R.data['매수전략'].unique())
                # print(strategy_list)
                # for strategy in strategy_list:
                #     스크린번호 = self.GetUnAssignedScreenNumber()
                #     print("a")
                #     이름 = str(strategy)+'_'+ R.lineEdit_name.text()
                #     print("b")
                #     종목리스트 = R.data[R.data['매수전략'] == strategy]
                #     print(종목리스트)
                #     r = CTradeShortTerm(sName=이름, UUID=uuid.uuid4().hex, kiwoom=self.kiwoom, parent=self)
                #     r.Setting(sScreenNo=스크린번호, 매수방법=매수방법, 매도방법=매도방법, 종목리스트=종목리스트)
                #     self.robots.append(r)
                스크린번호 = self.GetUnAssignedScreenNumber()
                이름 = R.lineEdit_name.text()
                종목리스트 = R.data
                r = CTradeShortTerm(sName=이름, UUID=uuid.uuid4().hex, kiwoom=self.kiwoom, parent=self)
                r.Setting(sScreenNo=스크린번호, 매수방법=매수방법, 매도방법=매도방법, 종목리스트=종목리스트)
                self.robots.append(r)

        except Exception as e:
            print('RobotAdd_TradeShortTerm', e)

    def RobotAutoAdd_TradeShortTerm(self, data):  # , strategy):
        # print("MainWindow : RobotAutoAdd_TradeShortTerm")
        try:
            스크린번호 = self.GetUnAssignedScreenNumber()
            # 이름 = strategy + '_TradeShortTerm'
            이름 = 'TradeShortTerm'
            매수방법 = '00'
            매도방법 = '03'
            종목리스트 = data
            print('추가 종목리스트')
            print(종목리스트)

            r = CTradeShortTerm(sName=이름, UUID=uuid.uuid4().hex, kiwoom=self.kiwoom, parent=self)
            r.Setting(sScreenNo=스크린번호, 매수방법=매수방법, 매도방법=매도방법, 종목리스트=종목리스트)
            self.robots.append(r)
            print('로봇 자동추가 완료')
            logger.info('로봇 자동추가 완료')
            Telegram('[XTrader]로봇 자동추가 완료')

        except Exception as e:
            print('RobotAutoAdd_TradeShortTerm', e)
            Telegram('[XTrader]로봇 자동추가 실패 : %s' % e, send='mc')

    def RobotEdit_TradeShortTerm(self, robot):
        R = 화면_TradeShortTerm(parent=self)
        R.lineEdit_name.setText(robot.sName)
        R.lineEdit_screen_number.setText('{:04d}'.format(robot.sScreenNo))
        R.comboBox_buy_condition.setCurrentIndex(R.comboBox_buy_condition.findText(robot.매수방법, flags=Qt.MatchContains))
        R.comboBox_sell_condition.setCurrentIndex(
            R.comboBox_sell_condition.findText(robot.매도방법, flags=Qt.MatchContains))

        if R.exec_():
            이름 = R.lineEdit_name.text()
            스크린번호 = int(R.lineEdit_screen_number.text())
            매수방법 = R.comboBox_buy_condition.currentText().strip()[0:2]
            매도방법 = R.comboBox_sell_condition.currentText().strip()[0:2]
            종목리스트 = R.data

            robot.sName = 이름
            robot.Setting(sScreenNo=스크린번호, 매수방법=매수방법, 매도방법=매도방법, 종목리스트=종목리스트)

    def RobotAutoEdit_TradeShortTerm(self, robot, data):
        # print("MainWindow : RobotAutoEdit_TradeShortTerm")
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
            logger.info('로봇 자동편집 완료')
            Telegram('[XTrader]로봇 자동편집 완료')
        except Exception as e:
            print('RobotAutoAdd_TradeShortTerm', e)
            Telegram('[XTrader]로봇 자동편집 실패 : %s' % e, send='mc')

    # 종목코드 생성
    def StockCodeBuild(self, to_db=False):
        try:
            result = []

            markets = [['0', 'KOSPI'], ['10', 'KOSDAQ']]#, ['8', 'ETF']]
            for [marketcode, marketname] in markets:
                codelist = self.kiwoom.dynamicCall('GetCodeListByMarket(QString)', [
                    marketcode])  # sMarket – 0:장내, 3:ELW, 4:뮤추얼펀드, 5:신주인수권, 6:리츠, 8:ETF, 9:하이일드펀드, 10:코스닥, 30:제3시장
                codes = codelist.split(';')

                for code in codes:
                    if code is not '':
                        종목명 = self.kiwoom.dynamicCall('GetMasterCodeName(QString)', [code])
                        종목명체크 = 종목명.lower().replace(' ', '')
                        주식수 = self.kiwoom.dynamicCall('GetMasterListedStockCnt(QString)', [code])
                        감리구분 = self.kiwoom.dynamicCall('GetMasterConstruction(QString)',
                                                       [code])  # 감리구분 – 정상, 투자주의, 투자경고, 투자위험, 투자주의환기종목
                        상장일 = datetime.datetime.strptime(
                            self.kiwoom.dynamicCall('GetMasterListedStockDate(QString)', [code]), '%Y%m%d')
                        전일종가 = int(self.kiwoom.dynamicCall('GetMasterLastPrice(QString)', [code]))
                        종목상태 = self.kiwoom.dynamicCall('GetMasterStockState(QString)', [
                            code])  # 종목상태 – 정상, 증거금100%, 거래정지, 관리종목, 감리종목, 투자유의종목, 담보대출, 액면분할, 신용가능

                        result.append([marketname, code, 종목명, 종목명체크, 주식수, 감리구분, 상장일, 전일종가, 종목상태])

            df_code = DataFrame(data=result,
                                columns=['시장구분', '종목코드', '종목명', '종목명체크', '주식수', '감리구분', '상장일', '전일종가', '종목상태'])
            # df.set_index('종목코드', inplace=True)

            if to_db == True:
                # 테마코드
                themecodes = []
                ret = self.kiwoom.dynamicCall('GetThemeGroupList(int)', [1]).split(';')
                for item in ret:
                    [code, name] = item.split('|')
                    themecodes.append([code, name])
                df_theme = DataFrame(data=themecodes, columns=['테마코드', '테마명'])

                # 테마구성종목
                themestocks = []
                for code, name in themecodes:
                    codes = self.kiwoom.dynamicCall('GetThemeGroupCode(QString)', [code]).replace('A', '').split(';')
                    for c in codes:
                        themestocks.append([code, c])
                df_themecode = DataFrame(data=themestocks, columns=['테마코드', '종목코드'])

                # 종목코드와 테마명 합침
                df_thememerge = pd.merge(df_theme, df_themecode, on='테마코드', how='outer')

                df_code['상장일'] = df_code['상장일'].apply(lambda x: (x.to_pydatetime()).strftime('%Y-%m-%d %H:%M:%S'))

                df_final = pd.merge(df_code, df_thememerge, on='종목코드', how='outer')
                df_final = df_final[['시장구분', '종목코드', '종목명', '종목명체크', '주식수', '감리구분', '상장일', '전일종가', '종목상태', '테마명']]

                conn = sqliteconn()
                df_final.to_sql('종목코드', conn, if_exists='replace', index=False)
                conn.close()

            return df_code

        except Exception as e:
            print('StockCodeBuild Error :', e)

if __name__ == "__main__":
    # 1.로그 인스턴스를 만든다.
    logger = logging.getLogger('XTrader')
    # 2.formatter를 만든다.
    formatter = logging.Formatter('[%(levelname)s|%(filename)s:%(lineno)s]%(asctime)s>%(message)s')

    loggerLevel = logging.DEBUG
    filename = "LOG/XTrader.log"

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

    current = datetime.datetime.now().strftime('%H:%M:%S')
    if current <= '08:58:00':
        Telegram("[XTrader]프로그램이 실행되었습니다.")


    # 프로그램 실행 후 3초 후에 한번 신호 받고, 그 다음 1초 마다 신호를 계속 받음
    QTimer().singleShot(3, window.OnQApplicationStarted)  # 3초 후에 한번만(singleShot) 신호받음 : MainWindow.OnQApplicationStarted 실행

    clock = QtCore.QTimer()
    clock.timeout.connect(window.OnClockTick)  # 1초마다 현재시간 읽음 : MainWindow.OnClockTick 실행
    clock.start(1000)  # 기존값 1000, 1초마다 신호받음

    sys.exit(app.exec_())
