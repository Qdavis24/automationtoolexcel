from classes import ExcelDf, Data, Worksheet, Word

if __name__ == "__main__":
    section_index = None
    question_index = None
    row_shift = None
    argb = None

    if input("Would you like to enter advanced configuration mode? ('Y' or 'N'): ").upper() == "Y":
        print("\n")
        if input(
                "Would you like to specify the index of the first excel row containing values? ('Y' or 'N'): ").upper() == "Y":
            print("\n")
            row_shift = int(
                input("Please look at the excel sheet and find the value of column A where the corrosponding row"
                      "has real values in it\n"
                      "EXAMPLE : \n"
                      "1 |  | Org ID: \n"
                      "2 |  | Reference \n"
                      "3 |  | 1.01 Guiding Principles \n"
                      "in this case row 3 is this index of the first row containing values \n"
                      "When you have figured out the row index, please input just the digit with no spaces below and press enter.\n"))
            print("\n")
        if input(
                "Would you like to specify index of the section title (Reference) column & question (RFI) column in the excel sheet? ('Y' or 'N'): ").upper() == "Y":
            print("\n")
            section_index = int(input("Please enter the index of the section title (Reference) column...\n"
                                      "Start at the left most column (index 0) (excel col 'A') and count each column to the right until you reach Reference column.\n"
                                      "EXAMPLE OF INDEXES : [0,1,2,3,4,5,6,7,...]\n"
                                      "When you have figured out the section title (Reference) column index, please input just the digit below and press enter.\n"))
            print("\n")

            question_index = int(input("Please enter the index of the question (RFI) column...\n"
                                       "Start at the left most column (index 0) (excel col 'A') and count each column to the right until you reach RFI column.\n"
                                       "EXAMPLE OF INDEXES : [0,1,2,3,4,5,6,7,...]\n"
                                       "When you have figured out the question (RFI) column index please input just the digit below and press enter.\n"))
            print("\n")
        if input(
                "Would you like to specify the ARGB (color code) fill value for RFI cells that should be ignored? ('Y' or 'N'): ").upper() == "Y":
            print("\n")
            argb = input("Please enter the argb value...\n"
                         "argb is alpha red green blue color format.\n"
                         "EXAMPLE : 'FF00B050' \n"
                         "When you have figured out the argb value, please input the full value with no spaces below and press enter.\n")
            print("\n")
    print("\n")
    EXCEL_FILEPATH = input("Please enter the filepath to your excel sheet (include extension .xlsx)\n"
                           "EXAMPLE (C:/Users/my-spreadsheet.xlsx)\n")
    print("\n")
    TEMPLATE_FILEPATH = input("Please enter the filepath to your word template (include extension .docx)\n"
                              "EXAMPLE (C:/Users/my-template.docx)\n")
    print("\n")
    EXPORT_FILEPATH = input("Please enter the filepath and name for you exported word doc (include extension .docx)\n"
                            "EXAMPLE (C:/Users/my-exported-document.docx)\n")
    print("\n")

    COLUMN_MAP = (("section", section_index or 2), ("question", question_index or 10))
    ROW_SHIFT = row_shift or 0
    HEADER = 1
    ARGB = argb or "FF00B050"

    excel = ExcelDf(EXCEL_FILEPATH, header=HEADER)
    ws = Worksheet(EXCEL_FILEPATH)
    data = Data()

    drops = ws.gen_drop_indexes_color("K", ARGB, row_shift=ROW_SHIFT)

    excel.compress_df(COLUMN_MAP)

    excel.remove_rows(drops)

    excel.remove_na()

    excel.remove_keywords(column="question", keyword="see", slice=(0, 3))

    data.map_questions_to_sections(sections=excel.uniques("section"), slice_key=(0, 4), df=excel.df, key="section",
                                   value="question")

    word = Word(TEMPLATE_FILEPATH, mapped_questions=data.mapped_questions)

    word.gen_modify_indexes()
    word.check_modify_indexes()
    word.modify()
    word.gen_delete_indexes()
    word.check_delete_indexes()
    word.remove()

    word.save(EXPORT_FILEPATH)
