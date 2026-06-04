import multiprocessing
multiprocessing.freeze_support()

import sys
import os
from streamlit.web import cli as stcli

if __name__ == '__main__':
    if getattr(sys, 'frozen', False):
        app_path = os.path.join(sys._MEIPASS, 'app.py')
    else:
        app_path = os.path.join(os.path.dirname(__file__), 'app.py')

    sys.argv = [
        "streamlit", "run", app_path,
        "--global.developmentMode", "false",   # ← 关键：禁用开发模式
        "--server.port", "8501",
        "--server.headless", "true",
        "--browser.gatherUsageStats", "false",
        "--server.enableXsrfProtection", "false"
    ]
    sys.exit(stcli.main())