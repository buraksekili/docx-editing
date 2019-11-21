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

while (curr_student < len(names) - 1):
    name = names[curr_student];
    name = name.replace(" ", "-")

    surname = surnames[curr_student]
    surname = surname.replace(" ", "-")

    school = schools[curr_student];
    school = school.replace(" ", "-")

    document.paragraphs[4].text = names[curr_student];
    document.paragraphs[5].text = surnames[curr_student];
    document.paragraphs[6].text = schools[curr_student];
    documentName = name + "-" + surname + "-" + school + ".docx";
    curr_student += 1;
    document.save(documentName)
