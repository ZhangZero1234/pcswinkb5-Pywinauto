# _*_ coding:UTF-8 _*_
import time
from pywinauto.application import Application
from pywinauto.keyboard import SendKeys
from pykeyboard import *
import xlrd
from math import ceil, floor
from datetime import date,datetime
# 启动PCOMM
k = PyKeyboard()

type_speed = .02

def key_input(str):
	for i in range(0,len(str)):
		time.sleep(.02)
		k.press_key(str[i]) # 模拟键盘按H键
		k.release_key(str[i]) # 模拟键盘松开H键

def key_tabs(num):
	for i in range(0,num):
		SendKeys('{TAB}')
		# time.sleep(type_speed)
def key_del(num):
	for i in range(0,num):
		SendKeys('{DELETE}')

def key_f_nine(num):
	for i in range(0,num):
		SendKeys('{F9}')

app = Application().start(r"C:\\Program Files (x86)\\IBM\\Personal Communications\\pcsfe.exe")
# app.IBMPersonalCommunicationsSessionManager.wait('visible')
# time.sleep(1)
app.IBMPersonalCommunicationsSessionManager.NewSession.set_focus()
app.IBMPersonalCommunicationsSessionManager.NewSession.click()
time.sleep(2)
app.IBMPersonalCommunicationsSessionManager.close()
app = Application().connect(title_re=".*Session A.*")
# 0x0228E550
print(app)
app.CustomizeCommnnication.LinkParameters.set_focus()
app.CustomizeCommnnication.LinkParameters.click()
time.sleep(.200)
app.Telnet3270.TypeKeys('aaa1.au.ibm.com')
time.sleep(.200)
app.Telnet3270.OK.click()
time.sleep(.200)
app.CustomizeCommnnication.OK.click()
dlg = app.window(title_re="Session A.*")
time.sleep(3)
dlg.set_focus()

# k.type_string('cicsa1p2'[::-1])
key_input("cicsa1p2")				
time.sleep(type_speed)
SendKeys('{ENTER}')


# dlg.print_control_identifiers()

# 读取Excel内容 
data = xlrd.open_workbook('Manual invoice V1.0.xlsm')# 打开Excel文件读取数据
table = data.sheets()[0]          #通过索引顺序获取
UseID = table.row(7)[4].value
Password = table.row(8)[4].value
BPICN = ceil(table.row(9)[4].value)
BPCode = table.row(10)[4].value

# print(BPICN)
# print(BPCode)
time.sleep(2)
key_input(UseID)
time.sleep(type_speed)
key_tabs(1)
key_input(Password)
time.sleep(type_speed)
SendKeys('{ENTER}')
time.sleep(1)
key_input("6")
time.sleep(type_speed)
SendKeys('{ENTER}')
time.sleep(type_speed)

nrows = table.nrows #行数 不为空的行数
# print(nrows)

# for i in range(13,nrows):
key_input("1")
time.sleep(type_speed)
SendKeys('{ENTER}')


time.sleep(1)
key_input(str(BPICN))
time.sleep(type_speed)
key_input(str(BPCode))
time.sleep(type_speed)
key_tabs(3)
key_input("sw1")
time.sleep(type_speed)
key_input("ppaot")
time.sleep(type_speed)
key_tabs(6)

key_input(str(ceil(table.row(13)[6].value))) #key_input(table.row(i)[6].value) Gi
time.sleep(type_speed)
key_tabs(9)

# key_input(str(ceil(table.row(13)[4].value))) #key_input(table.row(i)[4].value) Ei
date_value  = xlrd.xldate_as_tuple(table.row(13)[4].value,data.datemode)
date(*date_value[:3]).strftime('%Y-%m-%d')
# print(date(*date_value[:3]).strftime('%Y-%m-%d'))
key_input(date(*date_value[:3]).strftime('%Y-%m-%d'))
time.sleep(type_speed)
key_tabs(15)
key_input("  ppa") #key_input(table.row(i)[4].value) Ei
time.sleep(type_speed)
SendKeys('{ENTER}')
time.sleep(type_speed)
SendKeys('{F9}')
time.sleep(type_speed)
key_del(20)
time.sleep(type_speed)
key_input(str(ceil(table.row(13)[5].value))+"000") #key_input(table.row(i)[5].value) Fi
time.sleep(type_speed)
key_tabs(2)

# 获得Agreement值的前10位
fn = str(table.row(13)[7].value)[0:10] #str(ceil(table.row(i)[7].value))[0:10] 截取前十位字符串
key_input(fn)

key_tabs(1) # 9 39 
time.sleep(type_speed)
key_input(str(table.row(13)[8].value)) #key_input(table.row(i)[8].value) Ii
time.sleep(type_speed)

key_tabs(3) # 18 13
IN = "invoice#"+str(ceil(table.row(13)[3].value)) #IN = "invoice#"+str(ceil(table.row(i)[3].value))
key_input(IN)
time.sleep(type_speed)
SendKeys('{ENTER}')
time.sleep(type_speed)
SendKeys('{F4}')
time.sleep(type_speed)
key_f_nine(4)
time.sleep(type_speed)
SendKeys('{F3}')

