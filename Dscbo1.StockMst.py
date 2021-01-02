import sys
import win32com.client
import ctypes

################################################
# PLUS 공통 OBJECT
g_objCodeMgr = win32com.client.Dispatch('CpUtil.CpCodeMgr')
g_objCpStatus = win32com.client.Dispatch('CpUtil.CpCybos')
g_objCpTrade = win32com.client.Dispatch('CpTrade.CpTdUtil')
g_objDscbo1 = win32com.client.Dispatch('Dscbo1.StockMst')
g_objCpSysDib = win32com.client.Dispatch('CpSysDib.CpMarketWatch')
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
g_objDscbo1.SetInputValue(0, 'A005930')
g_objDscbo1.BlockRequest()
print(g_objDscbo1.GetHeaderValue(0))
code = g_objDscbo1.GetHeaderValue(0)  #종목코드
name= g_objDscbo1.GetHeaderValue(1)  # 종목명
time= g_objDscbo1.GetHeaderValue(4)  # 시간
cprice= g_objDscbo1.GetHeaderValue(11) # 종가
diff= g_objDscbo1.GetHeaderValue(12)  # 대비
open= g_objDscbo1.GetHeaderValue(13)  # 시가
high= g_objDscbo1.GetHeaderValue(14)  # 고가
low= g_objDscbo1.GetHeaderValue(15)   # 저가
offer = g_objDscbo1.GetHeaderValue(16)  #매도호가
bid = g_objDscbo1.GetHeaderValue(17)   #매수호가
vol= g_objDscbo1.GetHeaderValue(47)   #거래량
vol_value= g_objDscbo1.GetHeaderValue(48)  #거래대금

print("코드", code)
print("이름", name)
print("시간", time)
print("종가", cprice)
print("대비", diff)
print("시가", open)
print("고가", high)
print("저가", low)
print("매도호가", offer)
print("매수호가", bid)
print("거래량", vol)
print("거래대금", vol_value)
 
g_objCpSysDib.SetInputValue(0, 'A005930') 
g_objCpSysDib.BlockRequest()
value = g_objCpSysDib.GetDataValue(4,13)
print(value)