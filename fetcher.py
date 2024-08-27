"""
Author		: paradowski.michal@outlook.com
Description	: reads details of database and included tables
Updates:

* 2024-07-26 - v1.0 - creation
"""

# import generic libraries
import sys
import time


class MyFetcher:
    """
    Class responsible for obtaining database structure details.
    """

    def __init__(self, _core):
        """
        Initialize cass instance.
        """
        self._core = _core
        self._db_conn = self._set_connection()
        self.db_config = False
        self.db_tables = False
        self.db_procedures = False
        # obtain data depending on a document content settings
        if _core.doc_content['db_configuration']:
            self.db_config = self._get_db_configration()
        if _core.doc_content['db_tables']:
            self.db_tables = self._get_tables()
        if _core.doc_content['db_procedures']:
            self.db_procedures = self._get_procedures()

    def _set_connection(self):
        """
        Attempt to connect with database.
        :return: database connection object or exit script on failure
        :rtype: pyodbc.connect(), optional
        """
        print(f"{self._core.utils.timestamp()} connecting with database")
        _db_conn = self._core.utils.connect_db()
        if not _db_conn:
            print(f"{self._core.utils.timestamp()} ┗ [ERROR] cannot establish connection")
            self._core.log.warn("cannot establish pyodbc connection with database! exiting")
            sys.exit(0)
        else:
            print(f"{self._core.utils.timestamp()} ┗ [OK] connection set")
            self._core.log.info("connection with database established")
            return _db_conn

    def _get_db_configration(self):
        """
        Attempt to read database properties.
        :return: dictionary of configuration details
        :rtype: dict(str, any)
        """
        results = {}
        self._core.log.info("reading database configuration details")
        print(f"{self._core.utils.timestamp()} reading database configuration details")
        for subject in ['Configuration', 'Scoped configuration']:
            results[subject] = self._get_db_options(subject)
        return results

    def _get_tables(self):
        """
        Attempt to read database structure.
        :return: database details list or exit script on failure
        :rtype: list(any), optional
        """
        # attempt to obtain db details
        print(f"----------\n{self._core.utils.timestamp()} reading database structure")
        try:
            self._core.log.info("reading database details")
            query = ("SELECT TABLE_CATALOG, TABLE_SCHEMA, TABLE_NAME, TABLE_TYPE "
                     "FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME != 'sysdiagrams' ORDER BY TABLE_NAME")
            structure = self._core.utils.get_data(self._db_conn, query)
            print(f"{self._core.utils.timestamp()} ┗ [OK]")
        except Exception as exc:
            print(f"{self._core.utils.timestamp()} ┗ [ERROR] cannot read database details")
            self._core.log.warn(f"cannot read database details: {exc}. exiting")
            time.sleep(5)
            sys.exit(0)

        # proceed if succeeded
        print(f"----------\n{self._core.utils.timestamp()} reading table details")
        self._core.log.info("reading db tables")
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
                print(f"{self._core.utils.timestamp()} ┗ {table[2]}")
                self._core.log.info(f"OK {table[2]}")
            except Exception as ext:
                print(f"{self._core.utils.timestamp()} ┗ [ERROR] {table[2]}")
                self._core.log.warn(f"NOK - cannot read {table[2]} info: {ext}, skipping")
                continue
        # check data volume
        self._core.log.info(f"collected details of {len(results)} tables")
        if len(results) > 0:
            return results
        else:
            self._core.log.warn(f"no data for documentation, exiting")
            sys.exit(0)

    def _get_procedures(self):
        """
        Attempt to obtain basic details about stored procedures.
        :param str catalog: catalog name
        :param str schema: schema name
        :return: list of details per stored procedure
        """
        results = {}
        print(f"----------\n{self._core.utils.timestamp()} reading stored procedures")
        try:
            # get raw properties of stored procedures
            query = (f"select p.name, s.name, cast(p.create_date as varchar(32)), cast(p.modify_date as varchar(32)), "
                     f"cast(m.uses_ansi_nulls as varchar(max)), cast(m.uses_quoted_identifier as varchar(max)), "
                     f"cast(p.is_auto_executed as varchar(max)) "
                     f"from sys.procedures p "
                     f"inner join sys.schemas s on p.schema_id = s.schema_id "
                     f"inner join sys.sql_modules m on p.object_id = m.object_id "
                     f"where p.name not like 'sp%'")
            procedures = self._core.utils.get_data(self._db_conn, query)

            # obtain extended properties, return as dict
            if len(procedures)  == 0:
                return False
            else:
                for procedure in procedures:
                    details = {}
                    details['info'] = [['Created on', procedure[2]],
                                       ['Updated on', procedure[3]],
                                       ['Use ANSI nulls', procedure[4]],
                                       ['Use quoted identifier', procedure[5]],
                                       ['Is auto executed', procedure[6]]]
                    details['extended'] = self._get_procedure_ep(procedure[0], procedure[1])
                    results[procedure[0]] = details.copy()
            return results
        except Exception as exc:
            self._core.log.warn(f"cannot retrieve stored procedure details: {exc}")
            return False

    # utility methods for obtaining details

    def _get_column_details(self, catalog, schema, name):
        """
        Attempt to obtain column details: data type, if nullable, character lengths.
        :param str catalog: catalog name
        :param str schema: schema name
        :param str name: table name
        :return: table properties list
        :rtype: list(any)
        """
        query = (f"select column_name, data_type, isnull(cast(character_maximum_length as varchar), 'not set'),"
                 f" is_nullable from information_schema.columns "
                 f"where table_catalog = '{catalog}' and table_schema = '{schema}' and table_name = '{name}'")
        return self._core.utils.get_data(self._db_conn, query)

    def _get_key_details(self, catalog, schema, name):
        """
        Attempt to obtain key details for a given table.
        :param str catalog: catalog name
        :param str schema: schema name
        :param str name: table name
        :return: table keys list
        :rtype: list(any)
        """
        query = (f"select K.table_name, K.column_name, K.constraint_name, T.constraint_type "
                 f"from information_schema.key_column_usage as K "
                 f"join information_schema.table_constraints as T on K.constraint_name = T.constraint_name "
                 f"where K.table_catalog = '{catalog}' and K.table_schema = '{schema}' and K.table_name = '{name}'")
        return self._core.utils.get_data(self._db_conn, query)

    def _get_table_ep(self, schema, name):
        """
        Attempt to obtain extended properties of a given table.
        :param str schema: schema name
        :param str name: tabl name
        :return: table extended properties list
        :rtype: list(any)
        """
        query = (f"select cast(isnull(p.name, '---') as varchar(max)) as Property, "
                 f"cast(isnull(p.value, '---') as varchar(max)) as Value "
                 f"from sys.extended_properties p "
                 f"inner join sys.tables t on p.major_id = t.object_id "
                 f"inner join sys.schemas s on t.schema_id = s.schema_id "
                 f"where t.name = '{name}' and s.name = '{schema}' and p.minor_id = 0")
        return self._core.utils.get_data(self._db_conn, query)

    def _get_db_options(self, subject):
        """
        Attempt to read database options section.`
        :return: dictionary of configuration details
        :rtype: dict(str, any)
        """
        query = ""
        # set query
        if subject == 'Configuration':
            query = ("select name, "
                     "cast(value as nvarchar(max)) as value, "
                     "cast(value_in_use as nvarchar(max)) as value_in_use "
                     "from sys.configurations")
        if subject == 'Scoped configuration':
            query = ("select name, "
                     "cast(value as nvarchar(max)) as value "
                     "from sys.database_scoped_configurations")
        if not query:
            return ["Not configured"]
        else:
            # execute query
            try:
                print(f"{self._core.utils.timestamp()} ┗ {subject} [OK]")
                data = self._core.utils.get_data(self._db_conn, query)
                return data
            except Exception as exc:
                print(f"{self._core.utils.timestamp()} ┗ {subject} [ERROR]")
                self._core.log.warn(f"Cannot read database options: {exc}")
                return ["Not available"]

    def _get_procedure_ep(self, name, schema):
        """
        Attempt to read extended properties of a procedure
        :return: 
        """
        query = (f"select cast(ep.name as varchar(max)), cast(ep.value as varchar(max)) "
                 f"from sys.extended_properties ep "
                 f"join sys.objects o on ep.major_id = o.object_id "
                 f"join sys.schemas s on o.schema_id = s.schema_id "
                 f"where o.type = 'P' and s.name = '{schema}' and o.name = '{name}' "
                 f"order by ep.name;")
        properties = self._core.utils.get_data(self._db_conn, query)
        if len(properties) > 0:
            return properties
        else:
            return False


def execute(_core):
    """
    Carry service reports creation tasks.
    :raise exc: general unspecified exception.
    """
    try:
        return MyFetcher(_core)
    except Exception as exc:
        _core.log.warn(f"Unspecified data fetcher exception: {exc}")
        print(f"{_core.utils.timestamp()} [ERROR] unspecified data fetch exception, exiting...")
        return
