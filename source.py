import docx
from openpyxl import load_workbook
from docx.shared import Pt

workbook = load_workbook(filename="temp.xlsx")
sheet = workbook.active
names = []
surnames = []
schools = []
counter = 0
for value in sheet.iter_cols(min_row=1, max_col=3, values_only=True):
    for i in range(1, len(value)):
        if counter == 0:
            names.append(value[i])
        elif counter == 1:
            surnames.append(value[i])
        else:
            schools.append(value[i])
    counter += 1
document = docx.Document("tester.docx")

# 4 and 5 names
# 6 for school name
curr_student = 0

para_format = document.paragraphs[4].paragraph_format
nameRun = document.paragraphs[4].runs[0]
fontName = nameRun.font
fontName.size = Pt(18)

surname_run = document.paragraphs[5].runs[0]
font_surname = surname_run.font
font_surname.size = Pt(18)

school_run = document.paragraphs[6].runs[0]
font_school = school_run.font
font_school.size = Pt(12)

while (curr_student < len(names)):
    name = names[curr_student];
    name = name.replace(" ", "-")

    surname = surnames[curr_student]
    surname = surname.replace(" ", "-")

    school = schools[curr_student];
    school = school.replace(" ", "-")
    document.paragraphs[4].runs[0].text = names[curr_student];
    document.paragraphs[5].runs[0].text = surnames[curr_student];
    if  schools[curr_student] == "Sabancı Üniversitesi":
        print(document.paragraphs[6].runs[0].text)
        print(document.paragraphs[6].runs[1].text)
        print(document.paragraphs[6].runs[2].text)
    if schools[curr_student].count(' ') == 0:
        document.paragraphs[6].runs[2].text = "";
        document.paragraphs[6].runs[0].text = schools[curr_student];
    else:
        document.paragraphs[6].runs[2].text = "";
        document.paragraphs[6].runs[0].text = schools[curr_student];
    documentName = name + "-" + surname + "-" + school + ".docx";
    curr_student += 1;
    document.save(documentName)
