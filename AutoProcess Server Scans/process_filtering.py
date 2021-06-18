import os
import shutil
import sys
from run_filter import run_filter
import datetime

def processFiltering(path):
    i=1
    while os.path.isfile(path + '\\group-user with AD info '+str(i)+'.xlsx'):
        if not os.path.exists(path+'\\IS Developer Access '+str(i)):
            os.makedirs(path+'\\IS Developer Access '+str(i))
            input_path = path+'\\group-user with AD info '+str(i)+'.xlsx'
            output_path = path+'\\IS Developer Access '+str(i)
            shutil.copyfile(input_path,output_path+'\\IS Developers with Direct Access '+datetime.date.today().strftime('%m%d%Y')+'.xlsx')
            shutil.copyfile(input_path,output_path+'\\IS Developers with Access thru AD Group '+datetime.date.today().strftime('%m%d%Y')+'.xlsx')
            run_filter(input_path, output_path)
        else:
            print(f'IS Developer Access {str(i)} Folder already exists inside {path}')
            return
        i+=1