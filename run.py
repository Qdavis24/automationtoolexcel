from classes import ExcelDf, Data, Worksheet, Word, Config

if __name__ == "__main__":
    config = Config("./config.yaml")
    excel = ExcelDf(config.excel_filepath, header=config.header)
    ws = Worksheet(config.excel_filepath)
    data = Data()

    drops = ws.gen_drop_indexes_color("K", config.ignore_color, row_shift=config.row_shift)

    excel.compress_df(config.column_map)

    excel.remove_rows(drops)

    excel.remove_na()

    excel.remove_keywords(column="question", keyword="see", slice=(0, 3))

    data.map_questions_to_sections(sections=excel.uniques("section"), slice_key=(0, 4), df=excel.df, key="section",
                                   value="question")

    word = Word(config.word_template_filepath, mapped_questions=data.mapped_questions)

    word.gen_modify_indexes()
    word.modify()
    word.gen_delete_indexes()
    word.remove()

    word.save(config.word_export_filepath)
