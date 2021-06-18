import os
import shutil
from UnixACLscan import serverScanLinux
import sys
import datetime
from process_filtering_linux import processFilteringLinux
from log import log

def linuxRun(input_path):
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
            elif ext in ('.txt','.out'):
                newScanFound = True
                break
        except:
            print('Some error')
        
    # stop script if no new scans are available
    if newScanFound is False:
        print('No new scans were found in Linux')
        log('No new scans were found in Linux')
        return
    else:
        log('New server scans found in Linux')
        if not os.path.exists(new_input_path):
            os.makedirs(new_input_path) #creating new folder Scan <date>
        else:
            print('Scans for this date are already processed in Linux')
            log('Scans for this date are already processed in Linux')
            return


    # move new found scans in Scan <date>   
    log('Server List:')
    for f in os.listdir(input_path):
        filename, file_ext = os.path.splitext(f)
        try:
            if not file_ext:
                pass
            elif file_ext in ('.txt','.out'):
                log(filename)
                shutil.move(
                    os.path.join(input_path,f'{filename}{file_ext}'),
                    os.path.join(input_path,scan_folder,f'{filename}{file_ext}')
                )
        except:
            print("Error in Linux file moving")
            log("Error in Linux file moving")
            return

    # running jonathan scripts on the new scans
    serverScanLinux(new_input_path)

    # check if scanResults file is created or not
    req_file_created = False
    if os.path.isfile(new_input_path+'\\ScanResults.xlsx'):
        req_file_created = True

    if req_file_created:
        processFilteringLinux(new_input_path)
    else:
        print('File ScanResults was not found in '+new_input_path)
        log('File ScanResults was not found in '+new_input_path)
        return



