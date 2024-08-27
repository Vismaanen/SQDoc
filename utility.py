"""
Author		: paradowski.michal@outlook.com
Description	: utility methods for use across script functionalities
Updates:

* 2024-08-02 - v1.0 - creation
"""

# import generic libraries
import os
import logging
from datetime import datetime


class MyUtils:
    """
    Utility methods class.
    """

    def __init__(self):
        """
        Initialize class instance!
        """

    @staticmethod
    def timestamp():
        """
        Returns current timestamp for visual feedback.
        :return: timestamp string
        :rtype: str
        """
        return f'[{datetime.now().strftime("%H:%M:%S")}]'

    @staticmethod
    def get_logger(name):
        os.makedirs(os.path.dirname(name), exist_ok=True)
        # Configure logging
        logging.basicConfig(
            filename=name,
            level=logging.DEBUG,
            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
        )
        logger = logging.getLogger()
        return logger

    @staticmethod
    def get_data(db_conn, query):
        """
        Default function to obtain raw data from db based on a provided query.
        :param db_conn: database connection object
        :param str query: query string
        :type db_conn: pyodbc.connect()
        :return: list of records - result of query
        :rtype: list[Any], optional
        :raise Exception: ``e`` query execution error, returning None
        """
        result_list = []
        try:
            cursor = db_conn.cursor()
            cursor.execute(query)
            for record in cursor.fetchall():
                result_list.append(list(record))
            return result_list
        except Exception as exc:
            return None
