import os
import shutil
import sys
from run_filter_linux import runFilterLinux
import datetime

def processFilteringLinux(path):
    if not os.path.exists(path+'\\IS Developer Access'):
        os.makedirs(path+'\\IS Developer Access')
        input_path = path+'\\ScanResults.xlsx'
        output_path = path+'\\IS Developer Access'
        shutil.copyfile(input_path,output_path+'\\IS Developers with Direct Access '+datetime.date.today().strftime('%m%d%Y')+'.xlsx')
        runFilterLinux(input_path, output_path)
    else:
        sys.exit(f'IS Developer Access Folder already exists inside {path}')