"""
Description	: utility methods for use across script functionalities
"""

import sys
import main
import pyodbc
import logging
from typing import Any
from pathlib import Path
from datetime import datetime


class MyUtils:
    """
    Utility methods class.
    """

    def __init__(self):
        """
        Initialize class instance.
        """

    @staticmethod
    def create_log() -> logging.Logger:
        """
        Create and return app log file object.

        :return: log object
        :rtype: logging.Logger
        """
        # log parameters
        log_name = f"{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.log"

        try:
            # adjust target logs directory as needed - by default: subdirectory in a script location
            log_directory = Path(f"{main.PATH_LOGS}")
            log_directory.mkdir(parents=True, exist_ok=True)
            log_path = log_directory / log_name
            # create logger
            logger = logging.getLogger(log_name)
            logger.setLevel(logging.INFO)
            # file handler setting
            if not logger.handlers:
                file_handler = logging.FileHandler(log_path, mode='a', encoding='utf-8')
                file_handler.setFormatter(logging.Formatter(
                    '%(asctime)s - %(levelname)s - %(message)s',
                    datefmt='%Y-%m-%d %H:%M:%S'
                ))
                logger.addHandler(file_handler)
                # console output handler setting
                console_handler = logging.StreamHandler(sys.stdout)
                console_handler.setFormatter(logging.Formatter(
                    '%(asctime)s - %(levelname)s - %(message)s',
                    datefmt='%Y-%m-%d %H:%M:%S'
                ))
                logger.addHandler(console_handler)
            return logger
        except Exception as exc:
            print(f"Cannot create local log object: {str(exc)}; script will now exit.")
            exit()

    @staticmethod
    def get_data(db_conn: pyodbc.Connection, query: str, log: logging.Logger) -> list[Any] | None:
        """
        Default function to obtain raw data from db based on a provided query.

        :param db_conn: database connection object
        :param str query: query string
        :param log: log object
        :type db_conn: pyodbc.Connection
        :type log: logging.Logger
        :return: list of records - result of query
        :rtype: list[Any] or None
        :raise pyodbc.ProgrammingError: syntax error - debug query if valid
        :raise pyodbc.InterfaceError: database connection error
        :raise pyodbc.DataError: data error, verify if query can be executed
        :raise pyodbc.IntegrityError: possible keys / constraints corruption - verify query and data state in database
        :raise pyodbc.Error: general pyodbc error - debug connection state manually
        :raise Exception: unexpected exception - debug code validity
        """
        result_list = []
        try:
            cursor = db_conn.cursor()
            cursor.execute(query)
            for record in cursor.fetchall():
                result_list.append(list(record))
            cursor.close()
            return result_list
        except pyodbc.ProgrammingError as exc:
            log.error(f'> SQL syntax error: {exc}')
        except pyodbc.InterfaceError as exc:
            log.error(f"> connection error: {exc}")
        except pyodbc.DataError as exc:
            log.error(f"> data error: {exc}")
        except pyodbc.IntegrityError as exc:
            log.error(f"> integrity error: {exc}")
        except pyodbc.Error as exc:
            log.error(f"> general error: {exc}")
        except Exception as exc:
            log.error(f"> unexpected error: {exc}")
        return None
