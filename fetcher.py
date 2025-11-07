"""
Description	: module enabling reading details of database and included tables / stored procedures.
"""


import sys
import main
import pyodbc
import logging
from typing import Any


class MyFetcher:
    """
    Class responsible for obtaining database structure details.
    """

    def __init__(self, settings: dict[str, Any], log: logging.Logger):
        """
        Initialize cass instance.
        """
        self.log = log
        self.utils = settings['utils']
        self._db_conn = self._set_connection()
        self.db_config = False
        self.db_tables = False
        self.db_procedures = False
        # obtain data depending on a document content settings
        if settings['doc content']['db configuration']:
            self.db_config = self._get_db_configuration()
        if settings['doc content']['db tables']:
            self.db_tables = self._get_tables()
        if settings['doc content']['db procedures']:
            self.db_procedures = self._get_procedures()

    def _set_connection(self) -> pyodbc.Connection:
        """
        Attempt to connect with database, exit script on error.

        :return: database connection object or exit script on failure
        :rtype: pyodbc.Connection
        """
        self.log.info(f'connecting with database')
        try:
            _db_conn = pyodbc.connect(main.CONN_STRING)
            self.log.info(f'> connection OK')
            return _db_conn
        except Exception as exc:
            self.log.error(f'> database connection failed: {exc}')
            sys.exit(0)

    def _get_db_configuration(self) -> dict[str, Any] | None:
        """
        Attempt to read database properties.

        :return: dictionary of configuration details, optional
        :rtype: dict[str, Any] or None
        :raise Exception: general data properties retrieval exception
        """
        results = {}
        self.log.info("reading database configuration details")
        try:
            for subject in ['Configuration', 'Scoped configuration']:
                results[subject] = self._get_db_options(subject)
            return results
        except Exception as exc:
            self.log.warning(f'> cannot obtain details: {exc}')
        return None

    def _get_tables(self) -> dict[str, Any] | None:
        """
        Attempt to read database structure.

        :return: database tables details dict, optional
        :rtype: dict[str, Any] or None
        :raise Exception: ``exc`` table details retrieval exception - cannot read tables from database
        :raise Exception: ``exd`` table data parsing exception - validate results obtained in a previous query
        """
        # attempt to obtain db details
        self.log.info(f'reading database structure')
        try:
            query = ("SELECT TABLE_CATALOG, TABLE_SCHEMA, TABLE_NAME, TABLE_TYPE "
                     "FROM INFORMATION_SCHEMA.TABLES "
                     "WHERE TABLE_NAME != 'sysdiagrams' "
                     "ORDER BY TABLE_NAME")
            structure = self.utils.get_data(self._db_conn, query, self.log)
            self.log.info('> tables structure info OK')
        except Exception as exc:
            self.log.warning(f'> cannot read tables details: {exc}')
            return None

        # proceed if succeeded
        self.log.info("parsing db tables data")
        results = {}
        # loop tables
        for table in structure:
            try:
                details = {}
                catalog, schema, name, table_type = table
                # get column info
                details['columns'] = self._get_column_details(catalog, schema, name)
                # get keys info
                details['keys'] = self._get_key_details(catalog, schema, name)
                # get extended properties info
                details['extended'] = self._get_table_ep(schema, name)
                # append
                results[table[2]] = details
                self.log.info(f"OK {table[2]}")
            except Exception as exd:
                self.log.warning(f"> cannot read {table[2]} info: {exd}, skipping")
                continue
        # check data volume
        self.log.info(f"collected details of {len(results)} tables")
        if len(results) > 0:
            return results
        else:
            self.log.warning(f"no table data for documentation")
            return None

    def _get_procedures(self) -> dict[str, Any] | None:
        """
        Attempt to obtain basic details about stored procedures.

        :return: dictionary of stored procedure details, optional
        :rtype: dict[str, Any] or None
        :raise Exception: stored procedures info obtaining / parsing exception
        """
        results = {}
        self.log.info(f'reading stored procedures')
        try:
            # get raw properties of stored procedures
            query = (f"select p.name, s.name, cast(p.create_date as varchar(32)), cast(p.modify_date as varchar(32)), "
                     f"cast(m.uses_ansi_nulls as varchar(max)), cast(m.uses_quoted_identifier as varchar(max)), "
                     f"cast(p.is_auto_executed as varchar(max)) "
                     f"from sys.procedures p "
                     f"inner join sys.schemas s on p.schema_id = s.schema_id "
                     f"inner join sys.sql_modules m on p.object_id = m.object_id "
                     f"where p.name not like 'sp%'")
            procedures = self.utils.get_data(self._db_conn, query, self.log)

            # obtain extended properties, return as dict
            if len(procedures) == 0:
                return None
            else:
                for procedure in procedures:
                    details = {'info': [['Created on', procedure[2]],
                                        ['Updated on', procedure[3]],
                                        ['Use ANSI nulls', procedure[4]],
                                        ['Use quoted identifier', procedure[5]],
                                        ['Is auto executed', procedure[6]]],
                               'extended': self._get_procedure_ep(procedure[0], procedure[1])}
                    results[procedure[0]] = details.copy()
            # final data volume validation
            return results if results else None
        # in case of any unexpected exception
        except Exception as exc:
            self.log.warning(f"cannot retrieve stored procedure details: {exc}")
            return None

    # utility methods for obtaining details
    def _get_column_details(self, catalog: str, schema: str, name: str) -> list[Any] | None:
        """
        Attempt to obtain column details: data type, if nullable, character lengths.

        :param str catalog: catalog name
        :param str schema: schema name
        :param str name: table name
        :return: table properties list, optional
        :rtype: list[Any] or None
        """
        query = (f"select column_name, data_type, isnull(cast(character_maximum_length as varchar), 'not set'),"
                 f" is_nullable from information_schema.columns "
                 f"where table_catalog = '{catalog}' and table_schema = '{schema}' and table_name = '{name}'")
        return self.utils.get_data(self._db_conn, query, self.log)

    def _get_key_details(self, catalog: str, schema: str, name: str) -> list[Any] | None:
        """
        Attempt to obtain key details for a given table.

        :param str catalog: catalog name
        :param str schema: schema name
        :param str name: table name
        :return: table keys list, optional
        :rtype: list[Any] or None
        """
        query = (f"select K.table_name, K.column_name, K.constraint_name, T.constraint_type "
                 f"from information_schema.key_column_usage as K "
                 f"join information_schema.table_constraints as T on K.constraint_name = T.constraint_name "
                 f"where K.table_catalog = '{catalog}' and K.table_schema = '{schema}' and K.table_name = '{name}'")
        return self.utils.get_data(self._db_conn, query, self.log)

    def _get_table_ep(self, schema: str, name: str) -> list[Any] | None:
        """
        Attempt to obtain extended properties of a given table.

        :param str schema: database schema string
        :param str name: table name string
        :return: table extended properties list, optional
        :rtype: list[Any] or None
        """
        query = (f"select cast(isnull(p.name, '---') as varchar(max)) as Property, "
                 f"cast(isnull(p.value, '---') as varchar(max)) as Value "
                 f"from sys.extended_properties p "
                 f"inner join sys.tables t on p.major_id = t.object_id "
                 f"inner join sys.schemas s on t.schema_id = s.schema_id "
                 f"where t.name = '{name}' and s.name = '{schema}' and p.minor_id = 0")
        return self.utils.get_data(self._db_conn, query, self.log)

    def _get_db_options(self, subject: str) -> list[Any] | None:
        """
        Attempt to read database options section.

        :param str subject: database setting subject name string
        :return: list of configuration details, optional
        :rtype: list[Any] or None
        """
        if subject not in ['Configuration', 'Scoped configuration']:
            return ['Not configured']
        # set query
        if subject == 'Configuration':
            query = ("select name, "
                     "cast(value as nvarchar(max)) as value, "
                     "cast(value_in_use as nvarchar(max)) as value_in_use "
                     "from sys.configurations")
        else:
            query = ("select name, "
                     "cast(value as nvarchar(max)) as value "
                     "from sys.database_scoped_configurations")
        # execute query
        try:
            return self.utils.get_data(self._db_conn, query, self.log)
        except Exception as exc:
            self.log.warning(f"Cannot read database options: {exc}")
            return ["Not available"]

    def _get_procedure_ep(self, name: str, schema: str) -> list[Any] | None:
        """
        Attempt to read extended properties of a procedure.

        :param str name: object name string
        :param str schema: database schema string
        :return: list of procedure properties, optional
        :rtype: list[Any] or None
        """
        query = (f"select cast(ep.name as varchar(max)), cast(ep.value as varchar(max)) "
                 f"from sys.extended_properties ep "
                 f"join sys.objects o on ep.major_id = o.object_id "
                 f"join sys.schemas s on o.schema_id = s.schema_id "
                 f"where o.type = 'P' and s.name = '{schema}' and o.name = '{name}' "
                 f"order by ep.name;")
        properties = self.utils.get_data(self._db_conn, query, self.log)
        return properties if len(properties) > 0 else None
