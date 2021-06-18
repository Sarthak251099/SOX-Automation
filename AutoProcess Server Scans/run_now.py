from windows_prod_run import windowsProdRun
from windows_preprod_run import windowsPreProdRun
from linux_run import linuxRun
from log import log
import datetime

today = datetime.date.today()
today = today.strftime('%m%d%Y')
now = datetime.datetime.now()
current_time = now.strftime("%H:%M:%S")
log("\nDate : "+today+" || Time : "+current_time)
log('Pre Production:')
windowsPreProdRun('C:\\Python v2_sarthak\\Input Folder\\PreProd')
log('Production:')
windowsProdRun('C:\\Python v2_sarthak\\Input Folder\\Prod')
log('Linux:')
linuxRun('C:\\Python v2_sarthak\\Input Folder\\Linux')
