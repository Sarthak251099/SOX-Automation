import schedule
import time
from main_run import run
schedule.every(24).hours.do(run)
while True:
    schedule.run_pending()
    time.sleep(1)