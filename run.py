from classes import ExcelDf, Data, Word, Config

if __name__ == "__main__":
    config = Config("./config.yaml")
    excel = ExcelDf(filepath=config.excel_filepath, sheet_name=config.sheet_name,
                    row_shift=config.row_shift, header=config.header)

    excel.compress_df(config.standard_col_index, config.rfi_col_index)
    excel.clean_df(config.ignore_color, config.rfi_col)
    excel.display_dataframe()

    data = Data()
    data.map_questions_to_sections(sections=excel.uniques(excel.STANDARD_COLUMN_NAME), df=excel.df,
                                   rfi_col_name=excel.RFI_COLUMN_NAME,
                                   standard_col_name=excel.STANDARD_COLUMN_NAME)
    data.pretty_print()

    word = Word(filepath=config.word_template_filepath, mapped_questions=data.mapped_questions)
    word.modify()
    word.remove()

    word.save(filepath=config.word_export_filepath)

    input("Successfully generated document... \n Press enter to exit script.\n")
    exit(0)
