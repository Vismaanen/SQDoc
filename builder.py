"""
Author		: paradowski.michal@outlook.com
Description	: main executable of SQDoc - MSSQL database documentation script
Updates:

* 2024-08-02 - v1.0 - creation
"""

# import generic libraries
import os
import sys
from datetime import date
from docx import Document
from docx.shared import Inches, Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT


class MyPrinter:
    """
    Class responsible for exporting data into *.docx file.
    """

    def __init__(self, _details, _core):
        """
        Initialize cass instance.
        """
        self._db_name = _core.db_name
        self._log = _core.log
        self._utils = _core.utils
        self._export = _core.export
        self._db_config = _details.db_config
        self._db_tables = _details.db_tables
        self._properties = self._set_properties()
        self._print_document()

    @staticmethod
    def _set_properties():
        """
        Define properties of a document.
        :return: list of properties
        :rtype: list(any)
        """
        return [['Owner', 'ECS'],
                ['Author', 'Michał Paradowski'],
                ['E-mail', 'michal.paradowski@fujitsu.com'],
                ['Version', '1.0'],
                ['Status', 'Final'],
                ['Created on', date.today().strftime("%Y-%m-%d")]]

    def _print_document(self):
        """
        Attempt to export database information to *.docx file.
        """
        print(f"----------\n{self._utils.timestamp()} creating *.docx file")
        doc = Document()
        section = doc.sections[0]
        _logo_path = os.path.join(os.path.dirname(__file__), '.resources', 'doc_logo.png')

        # ---------------------------------------------------------------------------------
        # header and footer
        # ---------------------------------------------------------------------------------
        header = section.header
        header_paragraph = header.paragraphs[0]
        run = header_paragraph.add_run()
        run.add_picture(_logo_path, width=Inches(1.0))
        run.add_text(f" {self._db_name} database technical documentation")

        # Add page numbers to each section's footer
        for section in doc.sections:
            self._add_page_numbers(section)

        # ---------------------------------------------------------------------------------
        # main page
        # ---------------------------------------------------------------------------------
        doc.add_heading(f'{self._db_name} database technical documentation', 0)
        # document details table
        table = doc.add_table(rows=1, cols=2)
        table.style = 'Light List Accent 1'
        theader = table.rows[0].cells
        # header text
        theader[0].text = 'Document properties'
        theader[1].text = ''
        # column dimensions
        theader[0].width = Inches(2)
        theader[1].width = Inches(2)
        for item in self._properties:
            row = table.add_row().cells
            row[0].text = item[0]
            row[1].text = item[1]
            # column dimensions
            row[0].width = Inches(2)
            row[1].width = Inches(2)
        # next page
        doc.add_page_break()

        # ---------------------------------------------------------------------------------
        # table of contents
        # ---------------------------------------------------------------------------------
        doc.add_heading('Table of contents', level=1)
        toc_paragraph = doc.add_paragraph()
        toc_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        self._add_toc(toc_paragraph)
        # next page
        doc.add_page_break()

        # ---------------------------------------------------------------------------------
        # document purpose
        # ---------------------------------------------------------------------------------
        doc.add_heading('1. Document purpose', level=1)
        doc.add_paragraph(f'Purpose of this document is to provide a technical overview on details and structure of '
                          f'{self._db_name} database. Document covers configuration of the database itself and '
                          f'properties of each included table such as:')
        doc.add_paragraph(f'table columns', style='List Bullet')
        doc.add_paragraph(f'column types', style='List Bullet')
        doc.add_paragraph(f'keys', style='List Bullet')
        doc.add_paragraph(f'Details related to table contents and / or volume are not part o this document due to '
                          f'their potentially sensitive nature.')
        # next page
        doc.add_page_break()

        # ---------------------------------------------------------------------------------
        # database details
        # ---------------------------------------------------------------------------------
        doc.add_heading(f'2. {self._db_name} database details', level=1)
        doc.add_paragraph(f'This section covers basic configuration details of database.')
        # for each subject - create content table
        paragraph = 1
        for scope in self._db_config:
            content = self._db_config[scope]
            doc.add_heading(f"2.{paragraph} {scope}", level=2)
            # create table
            # separate formatting for different scope subjects
            if scope == 'Configuration':
                table = doc.add_table(rows=1, cols=3)
                table.autofit = False
                table.style = 'Light List Accent 1'
                theader = table.rows[0].cells
                # header text
                theader[0].text = 'Configuration item'
                theader[1].text = 'Value default'
                theader[2].text = 'Value in use'
                # column dimensions
                theader[0].width = Inches(2)
                theader[1].width = Inches(2)
                theader[2].width = Inches(2)
                for item, default, value in content:
                    row = table.add_row().cells
                    row[0].text = item
                    row[1].text = default
                    row[2].text = value
                    # column dimensions
                    row[0].width = Inches(2)
                    row[1].width = Inches(2)
                    row[2].width = Inches(2)

            if scope == "Scoped configuration":
                table = doc.add_table(rows=1, cols=2)
                table.autofit = False
                table.style = 'Light List Accent 1'
                theader = table.rows[0].cells
                # header text
                theader[0].text = 'Configuration item'
                theader[1].text = 'Value'
                # column dimensions
                theader[0].width = Inches(3)
                theader[1].width = Inches(3)
                for item, value in content:
                    row = table.add_row().cells
                    row[0].text = item
                    row[1].text = value
                    # column dimensions
                    row[0].width = Inches(3)
                    row[1].width = Inches(3)

            # increment for table header
            paragraph += 1
        # next page
        doc.add_page_break()

        # ---------------------------------------------------------------------------------
        # per-table details
        # ---------------------------------------------------------------------------------
        doc.add_heading(f'3. {self._db_name} tables', level=1)
        doc.add_paragraph(f'This section covers basic configuration details of database tables.')
        # for each subject - create content table
        paragraph = 1
        for table in self._db_tables:
            content = self._db_tables[table]
            doc.add_heading(f"3.{paragraph} {table}", level=2)

            # section - keys
            doc.add_heading(f"3.{paragraph}.1 Keys", level=3)
            if len(content['keys']) == 0:
                doc.add_paragraph(f'No keys configured for this table.')
            else:
                # create table
                table = doc.add_table(rows=1, cols=3)
                table.autofit = False
                table.style = 'Light List Accent 1'
                theader = table.rows[0].cells
                # header text
                theader[0].text = 'Key column name'
                theader[1].text = 'Constraint name'
                theader[2].text = 'Constraint type'
                # column dimensions
                theader[0].width = Inches(2)
                theader[1].width = Inches(2)
                theader[2].width = Inches(2)
                for key in content['keys']:
                    row = table.add_row().cells
                    row[0].text = key[1]
                    row[1].text = key[2]
                    row[2].text = key[3]
                    row[0].width = Inches(2)
                    row[1].width = Inches(2)
                    row[2].width = Inches(2)

            # section - extended properties
            doc.add_heading(f"3.{paragraph}.2 Extended properties", level=3)
            if len(content['extended']) == 0:
                doc.add_paragraph(f'No extended properties configured for this table.')
            else:
                # create table
                table = doc.add_table(rows=1, cols=2)
                table.autofit = False
                table.style = 'Light List Accent 1'
                theader = table.rows[0].cells
                # header text
                theader[0].text = 'Property name'
                theader[1].text = 'Property value'
                # column dimensions
                theader[0].width = Inches(2)
                theader[1].width = Inches(4)
                for key in content['extended']:
                    row = table.add_row().cells
                    row[0].text = key[0]
                    row[1].text = key[1]
                    row[0].width = Inches(2)
                    row[1].width = Inches(4)

            # section - columns
            doc.add_heading(f"3.{paragraph}.3 Columns", level=3)
            # create table
            table = doc.add_table(rows=1, cols=4)
            table.autofit = False
            table.style = 'Light List Accent 1'
            theader = table.rows[0].cells
            # header text
            theader[0].text = 'Column name'
            theader[1].text = 'Data type'
            theader[2].text = 'Max length'
            theader[3].text = 'Null-able'
            # column dimensions
            theader[0].width = Inches(2)
            theader[1].width = Inches(2)
            theader[2].width = Inches(1)
            theader[3].width = Inches(1)
            for column in content['columns']:
                row = table.add_row().cells
                row[0].text = column[0]
                row[1].text = column[1]
                row[2].text = column[2]
                row[3].text = column[3]
                row[0].width = Inches(2)
                row[1].width = Inches(2)
                row[2].width = Inches(1)
                row[3].width = Inches(1)
            # increment for table header
            paragraph += 1

        # Save the document
        path = f"{self._export}\\{self._db_name}_documentation.docx"
        print(f"{self._utils.timestamp()} Saving file: {path}")
        doc.save(path)
        print()

    @staticmethod
    def _add_toc(paragraph):
        """
        Build table of content.
        :param paragraph: current page paragraph object
        :type paragraph: doc.add_paragraph()
        :return: modified paragraph with table of content configuration
        :rtype: doc.add_paragraph()
        """
        run = paragraph.add_run()
        char = OxmlElement('w:fldChar')
        char.set(qn('w:fldCharType'), 'begin')
        instr = OxmlElement('w:instrText')
        instr.set(qn('xml:space'), 'preserve')
        instr.text = 'TOC \\o "1-3" \\h \\z \\u'
        # ---
        char_ext_1 = OxmlElement('w:fldChar')
        char_ext_1.set(qn('w:fldCharType'), 'separate')
        char_ext_2 = OxmlElement('w:updateFields')
        char_ext_2.set(qn('w:val'), 'true')
        char_ext_3 = OxmlElement('w:fldChar')
        char_ext_3.set(qn('w:fldCharType'), 'end')
        # ---
        r_element = run._r
        r_element.append(char)
        r_element.append(instr)
        r_element.append(char_ext_1)
        r_element.append(char_ext_2)
        r_element.append(char_ext_3)
        return paragraph

    @staticmethod
    def _add_page_numbers(section):
        """
        Method to provide pages numeration.
        :param section: document section object
        """
        footer = section.footer
        paragraph = footer.paragraphs[0]
        paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        # paragraph for storing numbers
        run = paragraph.add_run()
        run.font.size = Pt(12)
        char = OxmlElement('w:fldChar')
        char.set(qn('w:fldCharType'), 'begin')
        # ---
        instr = OxmlElement('w:instrText')
        instr.set(qn('xml:space'), 'preserve')
        instr.text = 'PAGE'
        # ---
        char_ext = OxmlElement('w:fldChar')
        char_ext.set(qn('w:fldCharType'), 'end')
        # ---
        run._r.append(char)
        run._r.append(instr)
        run._r.append(char_ext)

        footer_text = footer.add_paragraph()
        footer_text.add_run("SQDoc 1.0")
        footer_text.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER


def execute(_details, _core):
    """
    Carry document print task.
    :param _details: MSSQL database details dictionary
    :param _core: SQDoc script object
    :type _details: dict(str, any)
    :type _core: SQDoc()
    :raise exc: Unspecified fil creation exception
    """
    try:
        MyPrinter(_details, _core)
        print(f"{_core.utils.timestamp()} Document saved, all activities finished.")
    except Exception as exc:
        print(f"{_core.utils.timestamp()} Unspecified exception, check log for details. Exiting...")
        _core.log.warn(f"Unspecified *.docx file creation exception: {exc}")
        sys.exit(0)
