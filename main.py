"""
Author		: paradowski.michal@outlook.com
Description	: main executable of SQDoc - MSSQL database documentation script
Updates:

* 2024-07-26 - v1.0 - creation
* 2024-08-27 - v1.0 - code cleanup, adding document property settings
"""

# import generic libraries
import os
import sys
import time
from datetime import date, datetime

# import engine modules
import fetcher as f
import builder as b
import utility as u

# global variables
_main_path = "C:\\SQDoc"
_logs_path = "C:\\SQDoc\\logs"
_docs_path = "C:\\SQDoc\\docx"
_db_name = "Neo_DB"
db_conn_string = "Driver={SQL Server};Server=G02PLXN08339\\SQLEXPRESS;Database=Neo_DB;Trusted_Connection=yes;"

# printed document properties
_doc_properties = [['Owner:', 'ECS'], ['Author:', 'Michal Paradowski'], ['E-mail:', 'michal.paradowski@fujitsu.com'],
                   ['Version:', '1.0'], ['Status:', 'Final'], ['Created on:', date.today().strftime("%Y-%m-%d")]]

# printed document content setting
# set sections available for printing to be included
_doc_content = {'db_configuration': True, 'db_tables': True, 'db_procedures': True}


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

███████╗ ██████╗ ██████╗ 
██╔════╝██╔═══██╗██╔══██╗ ██████╗  ██████╗
███████╗██║   ██║██║  ██║██╔═══██╗██╔════╝
╚════██║██║▄▄ ██║██║  ██║██║   ██║██║     
███████║╚██████╔╝██████╔╝╚██████╔╝╚██████╗
╚══════╝ ╚══▀▀═╝ ╚═════╝  ╚═════╝  ╚═════╝
__________________________________________
  MSSQL Database documentation processor
        """)

        # import utility methods class
        self.utils = u.MyUtils()

        # setup log
        _log_path = f"{_logs_path}\\{datetime.now().strftime("%Y-%m-%d_%H-%M-%S")}.log"
        self.log = self.utils.get_logger(_log_path)
        print(f"{self.utils.timestamp()} log file: {_log_path}\n----------")
        try:
            # directory checks
            os.makedirs(_logs_path, exist_ok=True)
            os.makedirs(_docs_path, exist_ok=True)
            # set document properties
            self.doc_content = _doc_content
            self.doc_properties = _doc_properties
            # set utilities
            self.db_name = _db_name
            self.export = _docs_path
            # proceed with db data fetch and export
            self.log.info(f"----------")
            self.log.info(f"new script execution")
            b.execute(f.execute(self), self)
        except Exception as exc:
            self.log.warn(f'Unspecified script exception: {exc}')
            time.sleep(5)
            sys.exit(1)


if __name__ == '__main__':
    SQDoc()
