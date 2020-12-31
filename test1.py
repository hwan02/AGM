import sys
from PyQt5.QtWidgets import *
import win32com.client
import ctypes
import time

################################################
# PLUS 공통 OBJECT
g_objCodeMgr = win32com.client.Dispatch('CpUtil.CpCodeMgr')
g_objCpStatus = win32com.client.Dispatch('CpUtil.CpCybos')
g_objCpTrade = win32com.client.Dispatch('CpTrade.CpTdUtil')
 
 
################################################


################################################
# PLUS 실행 기본 체크 함수
def InitPlusCheck():
    # 프로세스가 관리자 권한으로 실행 여부
    if ctypes.windll.shell32.IsUserAnAdmin():
        print('정상: 관리자권한으로 실행된 프로세스입니다.')
    else:
        print('오류: 일반권한으로 실행됨. 관리자 권한으로 실행해 주세요')
        return False
 
    # 연결 여부 체크
    if (g_objCpStatus.IsConnect == 0):
        print("PLUS가 정상적으로 연결되지 않음. ")
        return False
 
    # 주문 관련 초기화
    if (g_objCpTrade.TradeInit(0) != 0):
        print("주문 초기화 실패")
        return False
 
    return True
 
 
################################################


################################################
# CpEvent: 실시간 이벤트 수신 클래스
class CpEvent:
    def set_params(self, client, name, caller):
        self.client = client  # CP 실시간 통신 object
        self.name = name  # 서비스가 다른 이벤트를 구분하기 위한 이름
        self.caller = caller  # callback 을 위해 보관
 
    def OnReceived(self):
        # 실시간 처리 - 현재가 체결 데이터
        if self.name == 'stockcur':
            code = self.client.GetHeaderValue(0)  
            name = self.client.GetHeaderValue(1)  
            timess = self.client.GetHeaderValue(18)  # 초
            exFlag = self.client.GetHeaderValue(19)  # 예상체결 플래그
            cprice = self.client.GetHeaderValue(13)  # 현재가
            diff = self.client.GetHeaderValue(2)  # 대비
            cVol = self.client.GetHeaderValue(17)  # 순간체결수량
            vol = self.client.GetHeaderValue(9)  # 거래량
 
            if exFlag != ord('2'):
                return
 
            item = {}
            item['code'] = code
            item['time'] = timess
            item['diff'] = diff
            item['cur'] = cprice
            item['vol'] = cVol
 
            # 현재가 업데이트
            self.caller.updateCurData(item)
 
            return
 
 
################################################
# plus 실시간 수신 base 클래스
class CpPublish:
    def __init__(self, name, serviceID):
        self.name = name
        self.obj = win32com.client.Dispatch(serviceID)
        self.bIsSB = False
 
    def Subscribe(self, var, caller):
        if self.bIsSB:
            self.Unsubscribe()
 
        if (len(var) > 0):
            self.obj.SetInputValue(0, var)
 
        handler = win32com.client.WithEvents(self.obj, CpEvent)
        handler.set_params(self.obj, self.name, caller)
        self.obj.Subscribe()
        self.bIsSB = True
 
    def Unsubscribe(self):
        if self.bIsSB:
            self.obj.Unsubscribe()
        self.bIsSB = False
 
 
################################################
# CpPBStockCur: 실시간 현재가 요청 클래스
class CpPBStockCur(CpPublish):
    def __init__(self):
        super().__init__('stockcur', 'DsCbo1.StockCur')

# MACD 지표 계산
class CMACD:
    def __init__(self):
        self.objSeries = win32com.client.Dispatch("CpIndexes.CpSeries")
        self.objIndex = win32com.client.Dispatch("CpIndexes.CpIndex")
 
    # MACD 계산
    def makeMACD(self):
        result = {}
        # 지표 계산 object
        self.objIndex.series = self.objSeries
        self.objIndex.put_IndexKind("MACD")  # 계산할 지표: MACD
        self.objIndex.put_IndexDefault("MACD")  # MACD 지표 기본 변수 자동 세팅
 
        print("MACD 변수", self.objIndex.get_Term1(), self.objIndex.get_Term2(), self.objIndex.get_Signal())
 
        # 지표 데이터 계산 하기
        self.objIndex.Calculate()
 
        cntofIndex = self.objIndex.ItemCount
        print("지표 개수:  ", cntofIndex)
        indexName = ["MACD", "SIGNAL", "OSC"]
 
        result['MACD'] = []
        result['SIGNAL'] = []
        result['OSC'] = []
        for index in range(cntofIndex):
            cnt = self.objIndex.GetCount(index)
            for j in range(cnt):
                value = self.objIndex.GetResult(index, j)
                result[indexName[index]].append(value)

        print('MACD %.2f SIGNLA %.2f OSC %.2f' % (result['MACD'][-1], result['SIGNAL'][-1], result['OSC'][-1]))
        return (True, result)
 

 # 분차트 관리 클래스
#   주어진 주기로 분차트 조회 , 실시간 분차트 데이터 생성, MACD 계산 호출
class CMinchartData:
    def __init__(self):
        self.objCur = {}
        self.data = {}
        self.code = {}
        self.objMACD = CMACD()
        self.LASTTIME = 1530
 
        # 오늘 날짜
        now = time.localtime()
        self.todayDate = now.tm_year * 10000 + now.tm_mon * 100 + now.tm_mday
        print(self.todayDate)
 
    def MonCode(self, code):
        self.data = {}
        self.code = code
 
        self.data['MACD'] = []
        self.data['SIGNAL'] = []
        self.data['OSC'] = []
 
        # MACD 계산 하기
        ret, result = self.objMACD.makeMACD()
 
        self.data['MACD'] = result['MACD']
        self.data['SIGNAL'] = result['SIGNAL']
        self.data['OSC'] = result['OSC']
 
        # 실시간 시세 요청
        if (code not in self.objCur):
            self.objCur[code] = CpPBStockCur()
            self.objCur[code].Subscribe(code, self)
 
    def stop(self):
        for k, v in self.objCur.items():
            v.Unsubscribe()
        self.objCur = {}
 
    def printdata(self):
        for i in range(len(self.code)):
            print(
                  self.data['MACD'][i],
                  self.data['SIGNAL'][i],
                  self.data['OSC'][i])
 
 # 테스트를 위한 메인 화면
class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()
 
        # plus 상태 체크
        if InitPlusCheck() == False:
            exit()
        # 원하는 종목 받아와서 for문 돌리기 
        self.minData = CMinchartData(5)
        self.minData.MonCode('A069500')
 
        
        nH = 20
 
        btnPrint = QPushButton('print', self)
        btnPrint.move(20, nH)
        btnPrint.clicked.connect(self.btnPrint_clicked)
        nH += 50
 
        btnExit = QPushButton('종료', self)
        btnExit.move(20, nH)
        btnExit.clicked.connect(self.btnExit_clicked)
        nH += 50
 
    def btnPrint_clicked(self):
        self.minData.printdata()
        return
 
    def btnExit_clicked(self):
        self.minData.stop()
        exit()
        return
 
 
if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    app.exec_()
 