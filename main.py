import pandas
from docx import Document
from docx.shared import Inches

try:
    dataframe = pandas.read_excel("./real.xlsx")
    df = dataframe.iloc[:, [2, 8]]
    df.columns = ["section", "question"]

    diff_sections = df["section"]
    diff_sections = diff_sections.dropna()
    diff_sections = diff_sections.unique()

    output = {}

    for section in diff_sections:
        strn = section[:4]
        output[strn] = df[df["section"] == section]["question"].dropna().tolist()
    print(output)



except Exception as e:
    print(e)

else:
    try:
        looking = True
        curr_section = ""
        document = Document("./worddoc.docx")

        for par in document.paragraphs:
            original_lines = par.text.split("\n")

            target = 0
            for i in range(0, len(original_lines)):
                if looking:
                    if original_lines[i] == "Confirm/Submit/Describe…":
                        looking = False
                        original_lines[i] = "\n".join([str(e) + "\n\n[enter response here]\n\n" for e in output[curr_section]])
                        break
                if original_lines[i][9:13] in output:
                    curr_section = original_lines[i][9:13]
                    looking = True

            output_string = "".join(original_lines)
            par.text = output_string
        document.save("new_doc.docx")


    except Exception as e:
        print(e)
