import os

# define the relative path to the Python executable to use to call the main_process
PYTHON_EXE_REL_PATH = 'venv/Scripts/python.exe'

# define the filename of the main_process.py file
MAIN_PROCESS_FILENAME = 'main_process.py'

# define the scheduling parameters
SCHEDULING_PARAMETERS = {
    'trigger': 'cron',
    'minute': '*/5',
    'hour': '9-13',
    'day_of_week': 'mon-fri',
    'max_instances': 1
}

# get the absolute path to the directory containing this module
MODULE_DIR = os.getcwd()

# calculate the absolute path to the Python executable
PYTHON_EXE_PATH = os.path.join(MODULE_DIR, PYTHON_EXE_REL_PATH)

# calculate the absolute path to the directory containing the main_process.py file
MAIN_PROCESS_DIR_PATH = MODULE_DIR
