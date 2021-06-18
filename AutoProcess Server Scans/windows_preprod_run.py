import os
import shutil
from WindowsACLscanPreProd import serverScanPreProd
import sys
import datetime
from process_filtering import processFiltering
from log import log

def windowsPreProdRun(input_path):
    today = datetime.date.today()
    today = today.strftime('%m%d%Y')

    scan_folder = 'Scans '+today
    new_input_path = input_path + '\\' + scan_folder
    newScanFound  = False

    # check if new csv files are present
    for f in os.listdir(input_path):
        file_name, ext = os.path.splitext(f)
        try:
            if not ext:
                pass
            elif ext in ('.csv','.xls','.xlsx'):
                newScanFound = True
                break
        except:
            print('Some error in PreProd')
        
    # stop script if no new scans are available
    if newScanFound is False:
        print('No new scans were found in PreProd')
        log('No new scans were found in PreProd')
        return
    else:
        log('New server scans found in Preprod')
        if not os.path.exists(new_input_path):
            os.makedirs(new_input_path) #creating new folder Scan <date>
        else:
            print('Scans for this date are already processed in PreProd')
            log('Scans for this date are already processed in PreProd')
            return

    log('Server List:')

    # move new found scans in Scan <date>
    for f in os.listdir(input_path):
        filename, file_ext = os.path.splitext(f)
        try:
            if not file_ext:
                pass
            elif file_ext in ('.xls','.csv','.xlsx'):
                log(filename)
                shutil.move(
                    os.path.join(input_path,f'{filename}{file_ext}'),
                    os.path.join(input_path,scan_folder,f'{filename}{file_ext}')
                )
        except:
            print("Error in PreProd")
            log('Error in moving files in PreProd')

    # running jonathan scripts on the new scans
    serverScanPreProd(new_input_path)

    # check if group user with ad info file is created or not
    john_result_path = new_input_path+'\\results'
    req_file_created = False
    if os.path.isfile(john_result_path+'\\group-user with AD info 1.xlsx'):
        req_file_created = True

    if req_file_created:
        processFiltering(john_result_path)
    else:
        print("No group-user with AD info 1 file was found in "+new_input_path)
        log("No group-user with AD info 1 file was found in "+new_input_path)



