"""
Author		: paradowski.michal@outlook.com
Description	: main executable of SQDoc - MSSQL database documentation script
Updates:

* 2024-07-26 - v1.0 - creation
"""

# import generic libraries
import sys
import time
from datetime import datetime

# import engine modules
import fetcher as f
import builder as b
import utility as u

# global variables
_main_path = "C:\\SQDoc"
_logs_path = "C:\\SQDoc\\logs"
_db_name = "Neo_DB"


class SQDoc:
    """
    Main executable.
    """
    def __init__(self):
        """
        Initialize class instance.
        """
        print("""
__________________________________________

███████╗ ██████╗ ██████╗  ██████╗  ██████╗
██╔════╝██╔═══██╗██╔══██╗██╔═══██╗██╔════╝
███████╗██║   ██║██║  ██║██║   ██║██║     
╚════██║██║▄▄ ██║██║  ██║██║   ██║██║     
███████║╚██████╔╝██████╔╝╚██████╔╝╚██████╗
╚══════╝ ╚══▀▀═╝ ╚═════╝  ╚═════╝  ╚═════╝
__________________________________________
  MSSQL Database documentation processor
        """)

        # Startup
        self.utils = u.MyUtils()
        self.log = self.utils.get_logger(f"{_logs_path}\\{datetime.now().strftime("%Y-%m-%d_%H-%M-%S")}.log")
        try:
            self.db_name = _db_name
            self.log.info(f"----------")
            self.log.info(f"new script execution")
            b.execute(f.execute(self), self)
        except Exception as exc:
            self.log.warn(f'Unspecified script exception: {exc}')
            time.sleep(5)
            sys.exit(1)


def main():
    SQDoc()


if __name__ == '__main__':
    main()

