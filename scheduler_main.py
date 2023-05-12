import subprocess
import threading
import time

from apscheduler.schedulers.background import BackgroundScheduler
from pytz import timezone

from untracked_config.scheduling_data import MAIN_PROCESS_DIR_PATH, MAIN_PROCESS_FILENAME, PYTHON_EXE_PATH, \
    SCHEDULING_PARAMETERS


# define the job function to run main_process.py in a new process
def run_main_process() -> None:
    """Run the main process.

    This function executes the main process by spawning a new process using `subprocess.Popen()`. It runs the main
    process file (`MAIN_PROCESS_FILENAME`) using the specified Python executable path (`PYTHON_EXE_PATH`) and sets the
    current working directory to `MAIN_PROCESS_DIR_PATH`.

    Returns:
        None
    """
    subprocess.Popen([PYTHON_EXE_PATH, MAIN_PROCESS_FILENAME], cwd=MAIN_PROCESS_DIR_PATH)


# define a function to start the scheduler in a separate thread
def start_scheduler() -> None:
    """Start the scheduler to run the main process job.

    This function creates a new scheduler instance, sets the timezone to 'US/Eastern', adds the main process job to the
    scheduler using the scheduling parameters from the configuration file, and starts the scheduler. It then enters an
    infinite loop to keep the scheduler running.

    Returns:
        None
    """
    # Create a new scheduler instance and set the timezone
    scheduler = BackgroundScheduler(timezone=timezone('US/Eastern'))

    # Add the job to the scheduler using the scheduling parameters from the configuration file
    scheduler.add_job(run_main_process, **SCHEDULING_PARAMETERS)

    # Start the scheduler
    scheduler.start()

    # Enter an infinite loop to keep the scheduler running
    while True:
        time.sleep(1)



# start the scheduler thread
scheduler_thread = threading.Thread(target=start_scheduler)
scheduler_thread.start()

# wait for the scheduler thread to complete
scheduler_thread.join()
