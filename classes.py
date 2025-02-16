import pandas
import pandas as pd
from openpyxl import load_workbook
from docx import Document
from docx.opc.exceptions import PackageNotFoundError


import yaml


class Config:
    def __init__(self, filepath):
        try:
            with open(filepath) as file:
                self.config = yaml.safe_load(file)
        except FileNotFoundError as error:
            input(f"{error}: config file 'config.yaml' not found.")
            exit(1)
        else:
            self.column_map = ((self.config['excel']['column_map']['column_title_0'],
                                self.config['excel']['column_map']['column_value_0']),
                               (self.config['excel']['column_map']['column_title_1'],
                                self.config['excel']['column_map']['column_value_1'],)
                               )
            self.row_shift = self.config['excel']['row_shift']
            self.header = self.config['excel']['header']
            self.ignore_color = self.config['excel']['ignore_color']
            self.excel_filepath = self.config['excel']['filepath']
            self.word_template_filepath = self.config['word']['template_path']
            self.word_export_filepath = self.config['word']['export_path']


class ExcelDf:
    def __init__(self, filepath: str, header: int):
        """ loads excel spreadsheet into DF object
        ARGS:
        ------------------------------------------
            filepath: filepath to excel sheet
            header: which row contains standard col names

        ATRRS:
        ----------------------------------------
            self.df = dataframe object built from excel sheet
            """
        try:
            self.df = pd.read_excel(filepath, header=header)
        except FileNotFoundError as error:
            input(f"{error}: Excel file path specified is incorrect.")
            exit(1)

    def compress_df(self, col_map: tuple):
        """ removes unwanted cols and sets the column names in df
         ARGS
         ---------------------------------------
         col_map: tuple with nested tuples, each nested tuple contains
         (the df column index for used col, the string to rename that col)
         """
        self.df = self.df.iloc[:, [tup[1] for tup in col_map]]
        self.df.columns = [tup[0] for tup in col_map]

    def remove_na(self):
        self.df = self.df.dropna()

    def remove_keywords(self, column: str, keyword: str, slice: tuple):
        """ searches df for rows which contain the keyword in specified column by slicing string in each cell
            and comparing the slice to the keyword -> removes all the rows with cell slice matching keyword from df

        ARGS
        ------------------------------------------
        column: string for column name we are searching through
        keyword: keyword we are looking to remove columns by
        slice: the slice into value string that will match keyword

        """
        self.df = self.df[(self.df[column].str[slice[0]:slice[1]] != keyword) &
                          (self.df[column].str[slice[0]:slice[1]] != keyword.title())]

    def remove_rows(self, remove_indexes: list):
        """ removes rows using list of indexes to remove
        ARGS:
        --------------------------------
            remove_indexes: list of indexes to remove from df
        """

        self.df = self.df.drop(remove_indexes)

    def uniques(self, column: str) -> pandas.Series:
        return self.df[column].unique()

    def return_series(self, column: str) -> pandas.Series:
        return self.df[column]


class Worksheet:
    def __init__(self, filepath: str):
        try:
            self.worksheet = load_workbook(filepath).active
        except FileNotFoundError as error:
            input(f"{error}: Excel filepath specified is incorrect.")
            exit(1)

    def gen_drop_indexes_color(self, col: str, argb: str, row_shift: int) -> list:
        """ searches excel spreadsheet by iterating through cells in a specified col and looks for cells with a background
        color matching the one given

        ARGS:
        -----------------------------------
            col: charector for excel column we are searching (goes by excels letter naming convention)
            argb: alpha red green blue value for fill color we wish to identify indexes of
            row_shift: the row value for the first row of the excel sheet (adjusts index to match df row indexes)

        RETURN:
        ----------------------------------
            list of indexes that contain the fill color specified, indexes will match the indexes in dataframe
            """
        return [cell.row - row_shift for cell in self.worksheet[col] if cell.fill.start_color.rgb == argb]

    def check_cell_fill(self, col: str):
        """ prints out the argb values for fill in each cell in a specified col """
        for cell in self.worksheet[col]:
            print(cell.fill.start_color.rgb)


class Data:
    def __init__(self):
        self.mapped_questions = {}

    def map_questions_to_sections(self, sections: list, df: pandas.DataFrame, slice_key: tuple, key: str, value: str):
        """ creates an organized dictionary where each key will be a title for data and the value will be all the
        values for that title

        ARGS
        --------------------------------
        sections: the unique list of values for a col in df (this will function as the title for data
        df: the dataframe object being used to make map
        slice_key: the slice of the sections we are using to title the collections of data
        (example) we are using just the number for a section '1.09' instead of the words for a section
        key: the col in df that will contain the title for a set of data
        value: the col in df that will have values that should be grouped with that title
        """
        for section in sections:
            section_qs = [q for q in df[df[key] == section][value]]
            if section_qs:
                self.mapped_questions[section[slice_key[0]:slice_key[1]]] = section_qs


class Word:
    def __init__(self, filepath: str, mapped_questions: dict):
        """ Word class will carry out all functionality relating to parsing and modifying the existing word template

        ATTR:
        -----------------------------------
        mapped_questions: the dictionary of titles of data mapped to the groups of data belonging to that title
        to_delete: a list of indexes that contain par indexes in document object that should be removed from doc
        to_modify: a list of indexes that contain par indexes in document object that should be replaced with mapped
        data

        """
        try:
            self.doc = Document(filepath)
        except PackageNotFoundError as error:
            input(f"{error}: word template file path specified is incorrect.")
            exit(1)
        self.mapped_questions = mapped_questions
        self.to_delete = []
        self.to_modify = []

    def gen_modify_indexes(self):
        """ generates all par indexes for our to be modified document paragraphs """
        curr_section = None
        for i, par in enumerate(self.doc.paragraphs):
            if self.check_for_question(par):

                if curr_section:
                    self.to_modify.append((curr_section, i))

            if self.check_for_section_head(par):
                identifier = par.text[9:13]
                curr_section = identifier if identifier in self.mapped_questions else False

    def gen_delete_indexes(self):
        """ generates all par indexes for our to be deleted document paragraphs """
        delete = False
        for i, par in enumerate(self.doc.paragraphs):
            if self.check_for_section_head(par):
                if delete:
                    self.to_delete.append((start, i - 1))
                start = i
                identifier = par.text[9:13]
                delete = identifier not in self.mapped_questions

    def modify(self):
        """ goes through the modify indexes and then inserts the new content using the map """
        for tup in self.to_modify[::-1]:
            curr_p = self.doc.paragraphs[tup[1]]
            curr_p.clear()
            curr_p.text = str(f"{self.mapped_questions[tup[0]][0]}\n\n[enter response here]\n\n")

            curr_p._p.get_or_add_pPr().get_or_add_numPr().get_or_add_numId().val = 1
            curr_p._p.get_or_add_pPr().get_or_add_numPr().get_or_add_ilvl().val = 0
            curr_p = curr_p._element

            for q in self.mapped_questions[tup[0]][1:]:
                new_p = self.doc.add_paragraph(f"{q}\n\n[enter response here]\n\n")
                new_p._p.get_or_add_pPr().get_or_add_numPr().get_or_add_numId().val = 1
                new_p._p.get_or_add_pPr().get_or_add_numPr().get_or_add_ilvl().val = 0

                curr_p.addnext(new_p._element)

    def remove(self):
        """ goes through delete indexes and then deletes those paragraphs from document """
        for tup in self.to_delete[::-1]:
            for i in range(tup[1], tup[0] - 1, -1):
                p = self.doc.paragraphs[i]._element
                p.getparent().remove(p)

    def check_modify_indexes(self):
        """ prints out all sections with their current document text """
        for tup in self.to_modify:
            print(tup[0], self.doc.paragraphs[tup[1]].text)
            print("\n-------------------next-------------------")

    def check_delete_indexes(self):
        """ prints out all document paragraph text within delete indexes """
        for tup in self.to_delete:
            for i in range(tup[0], tup[1] + 1):
                print(self.doc.paragraphs[i].text)
            print("\n-------------------next-------------------")

    def check_for_section_head(self, par):
        """ checks for par that represents a new section in word doc
        ARGS:
        -------------------------
            par: the current par we are checking
        RETURN:
        -------------------------
            True if it is a section head
            False if not a section head
            """
        if par.text[14:15] == "–":
            return True
        return False

    def check_for_question(self, par):
        """ checks for par that represent where we should have our new questions inserted at in word doc
        ARGS:
        -------------------------
            par: the current par we are checking
        RETURN:
        -------------------------
            True if it is an insert questions par
            False if not an insert questions par
            """
        if par.text == "Confirm/Submit/Describe…":
            return True
        return False

    def save(self, filepath):
        self.doc.save(filepath)
