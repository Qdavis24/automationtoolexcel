from classes import ExcelDf, Data, Worksheet, Word

COLUMN_MAP = (("section", 2), ("question", 10))
HEADER = 1

excel = ExcelDf("./files/input.xlsx", header=HEADER)
ws = Worksheet("./files/input.xlsx")
data = Data()

drops = ws.gen_drop_indexes_color("K", "FF00B050")

excel.compress_df(COLUMN_MAP)

excel.remove_rows(drops)

excel.remove_na()

excel.remove_keywords(column="question", keyword="see", slice=(0, 3))

data.map_questions_to_sections(sections=excel.uniques("section"), slice_key=(0, 4), df=excel.df, key="section",
                               value="question")

word = Word("./files/template.docx", mapped_questions=data.mapped_questions)
for style in word.doc.styles:
    print(style)
word.gen_modify_indexes()
word.check_modify_indexes()
word.modify()
word.gen_delete_indexes()
word.remove()


word.save()




