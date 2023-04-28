import subprocess
import threading
import time

from apscheduler.schedulers.background import BackgroundScheduler
from pytz import timezone


# define a function to start the scheduler in a separate thread
def start_scheduler():
    # create a new scheduler instance and set the timezone
    scheduler = BackgroundScheduler(timezone=timezone('US/Eastern'))

    # define the job function to run main_process.py in a new process
    def run_main_process():
        subprocess.Popen(['C:\\Users\\lmcglaughlin\\PycharmProjects\\outlook_data\\venv\\Scripts\\python.exe', 'main_process.py'], cwd='C:\\Users\\lmcglaughlin\\PycharmProjects\\outlook_data')

    # add the job to the scheduler to run every 5 minutes between 7:00 AM and 7:00 PM, Monday through Friday
    scheduler.add_job(run_main_process, 'cron', minute='*/5', hour='7-18', day_of_week='mon-fri', max_instances=1)

    # start the scheduler
    scheduler.start()

    # enter an infinite loop to keep the scheduler running
    while True:
        time.sleep(1)

# start the scheduler thread
scheduler_thread = threading.Thread(target=start_scheduler)
scheduler_thread.start()

# wait for the scheduler thread to complete
scheduler_thread.join()
