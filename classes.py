import pandas as pd
from openpyxl import load_workbook
from docx import Document



class ExcelDf:
    def __init__(self, filepath, header):
        """ loads excel spreadsheet into DF object
        ARGS:
        ------------------------------------------
            filepath: filepath to excel sheet
            header: which row contains standard col names

        ATRRS:
        ----------------------------------------
            self.df = df attr
            self.cols = column names attr
            """
        try:
            self.df = pd.read_excel(filepath, header=header)
        except FileNotFoundError as error:
            print(f"{error}:Excel file path specified is incorrect")

    def compress_df(self, col_map: tuple):
        """ removes unwanted cols and sets the column names in df
         ARGS
         ---------------------------------------
         col_map: tuple with nested tuples for each col index and renaming string needed
         """
        self.df = self.df.iloc[:, [tup[1] for tup in col_map]]
        self.df.columns = [tup[0] for tup in col_map]

    def remove_na(self):
        self.df = self.df.dropna()

    def remove_keywords(self, column, keyword, slice: tuple):
        """ searches df for rows which contain the keyword in specified column value slice and
            removes rows
        ARGS
        ------------------------------------------
        column: string for column name we are searching through
        keyword: keyword we are looking to remove columns by
        slice: the slice into value string that will match keyword

        """
        self.df = self.df[(self.df[column].str[slice[0]:slice[1]] != keyword) &
                          (self.df[column].str[slice[0]:slice[1]] != keyword.title())]

    def remove_rows(self, remove_indexes):
        """ removes rows using list of indexes to remove
        ARGS:
        --------------------------------
            remove_indexes: list of indexes to remove from df
        """

        self.df = self.df.drop(remove_indexes)

    def uniques(self, column):
        return self.df[column].unique()

    def return_series(self, column):
        return self.df[column]


class Worksheet:
    def __init__(self, filepath):
        try:
            self.worksheet = load_workbook(filepath).active
        except FileNotFoundError as error:
            print(f"{error}: Excel filepath specified is incorrect")

    def gen_drop_indexes_color(self, col, argb: str):
        return [cell.row - 3 for cell in self.worksheet[col] if cell.fill.start_color.rgb == argb]

    def check_cell_fill(self, col):
        for cell in self.worksheet[col]:
            print(cell.fill.start_color.rgb)


class Data:
    def __init__(self):
        self.mapped_questions = {}

    def map_questions_to_sections(self, sections, df, slice_key, key, value):
        for section in sections:
            section_qs = [q for q in df[df[key] == section][value]]
            if section_qs:
                self.mapped_questions[section[slice_key[0]:slice_key[1]]] = section_qs


class Word:
    def __init__(self, filepath, mapped_questions):
        try:
            self.doc = Document(filepath)
        except FileNotFoundError as error:
            print(f"{error}: word doc file path specified is not correct")
        self.mapped_questions = mapped_questions
        self.to_delete = []
        self.to_modify = []

    def gen_modify_indexes(self):
        curr_section = None
        for i, par in enumerate(self.doc.paragraphs):
            if self.check_for_question(par):

                if curr_section:
                    self.to_modify.append((curr_section, i))

            if self.check_for_section_head(par):
                identifier = par.text[9:13]
                curr_section = identifier if identifier in self.mapped_questions else False

    def gen_delete_indexes(self):
        delete = False
        for i, par in enumerate(self.doc.paragraphs):
            if self.check_for_section_head(par):
                if delete:
                    self.to_delete.append((start, i - 1))
                start = i
                identifier = par.text[9:13]
                delete = identifier not in self.mapped_questions

    def modify(self):
        for tup in self.to_modify[::-1]:
            curr_p = self.doc.paragraphs[tup[1]]
            curr_p.clear()
            curr_p.text = str(f"\n\n{self.mapped_questions[tup[0]][0]}\n\n[enter response here]\n\n")

            curr_p._p.get_or_add_pPr().get_or_add_numPr().get_or_add_numId().val = 1
            curr_p._p.get_or_add_pPr().get_or_add_numPr().get_or_add_ilvl().val = 0
            curr_p = curr_p._element

            for q in self.mapped_questions[tup[0]][1:]:
                new_p = self.doc.add_paragraph(f"\n\n{q}\n\n[enter response here]\n\n")
                new_p._p.get_or_add_pPr().get_or_add_numPr().get_or_add_numId().val = 1
                new_p._p.get_or_add_pPr().get_or_add_numPr().get_or_add_ilvl().val = 0

                curr_p.addnext(new_p._element)



    def remove(self):
        for tup in self.to_delete[::-1]:
            for i in range(tup[1], tup[0] - 1, -1):
                p = self.doc.paragraphs[i]._element
                p.getparent().remove(p)

    def check_modify_indexes(self):
        for tup in self.to_modify:
            print(tup[0], self.doc.paragraphs[tup[1]].text)
            print("\n-------------------next-------------------")

    def check_delete_indexes(self):
        for tup in self.to_delete:
            for i in range(tup[0], tup[1] + 1):
                print(self.doc.paragraphs[i].text)
            print("\n-------------------next-------------------")

    def check_for_section_head(self, par):
        if par.text[14:15] == "–":
            return True
        return False

    def check_for_question(self, par):
        if par.text == "Confirm/Submit/Describe…":
            return True
        return False

    def save(self):
        self.doc.save("New.docx")
