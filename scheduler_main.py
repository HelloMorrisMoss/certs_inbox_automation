import subprocess
import threading
import time

from apscheduler.schedulers.background import BackgroundScheduler
from pytz import timezone

from untracked_config.scheduling_data import MAIN_PROCESS_DIR_PATH, MAIN_PROCESS_FILENAME, PYTHON_EXE_PATH, \
    SCHEDULING_PARAMETERS


# define the job function to run main_process.py in a new process
def run_main_process():
    subprocess.Popen([PYTHON_EXE_PATH, MAIN_PROCESS_FILENAME], cwd=MAIN_PROCESS_DIR_PATH)

# define a function to start the scheduler in a separate thread
def start_scheduler():
    # create a new scheduler instance and set the timezone
    scheduler = BackgroundScheduler(timezone=timezone('US/Eastern'))

    # add the job to the scheduler using the scheduling parameters from the configuration file
    scheduler.add_job(run_main_process, **SCHEDULING_PARAMETERS)

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
