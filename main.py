"""
Description	: main executable of SQDoc - MSSQL database documentation script
"""

# import generic libraries
import os
import socket
from datetime import date

# import engine modules
import fetcher as f
import builder as b
import utility as u

# global variables
PATH_MAIN = "C:\\Temp\\SQDoc"
PATH_LOGS = "C:\\Temp\\SQDoc\\Logs"
PATH_DOCS = "C:\\Temp\\SQDoc\\Docx"
DB_NAME = "Neo_DB"
CONN_STRING = (f"Driver={{SQL Server}};"
               f"Server={socket.gethostname()}\\SQLEXPRESS;"
               f"Database={DB_NAME};"
               f"Trusted_Connection=yes;"
               )

# printed document properties
DOC_PROPERTIES = [
    ['Owner:', 'Owner unit'],
    ['Author:', 'Author Name'],
    ['E-mail:', 'dbadmin@domain.com'],
    ['Version:', '1.0'],
    ['Status:', 'Final'],
    ['Created on:', date.today().strftime("%Y-%m-%d")]
]

# printed document content setting
# set sections available for printing to be included
DOC_CONTENT = {
    'db configuration': True,
    'db tables': True,
    'db procedures': True
}


def main() -> None:
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
    utils = u.MyUtils()

    # setup log
    log = utils.create_log()
    try:
        # directory checks
        os.makedirs(PATH_LOGS, exist_ok=True)
        os.makedirs(PATH_DOCS, exist_ok=True)
        # proceed with db data fetch and export
        log.info(f"----------")
        log.info(f"new script execution")
        settings = {'doc content': DOC_CONTENT,
                    'doc properties': DOC_PROPERTIES,
                    'utils': utils
                    }
        # get data related to chosen settings
        data = f.MyFetcher(settings, log)
        # format data into a report
        b.MyPrinter(settings, data, log)
    except Exception as exc:
        log.error(f'Unspecified script exception: {exc}')
        exit()


if __name__ == '__main__':
    main()
