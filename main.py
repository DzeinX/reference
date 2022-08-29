import pandas as pd
from docxtpl import DocxTemplate

xl = pd.read_excel('spravka.xlsx')


for num_st in range(0, len(xl)):

    student = xl[xl.index == num_st]


    doc = DocxTemplate("spravka_shablon.docx")

    context = {'name': ''.join(student.Имя.tolist()),
               'dateOfBirth': str(student.Дата_рождения.tolist()[0]).split()[0],
               'course': student.Курс.tolist()[0],
               'formOfEd': ''.join(student.Форма_обучения.tolist()),
               'edProgram': ''.join(student.Образовательная_программа.tolist()),
               'directionOfTraining': student.Направление_подготовки.tolist()[0],
               'nameOfUniversity': ''.join(student.Институт.tolist()),
               'dateOfStart': str(student.Дата_начала_обучения.tolist()[0]).split()[0],
               'dateOfEnd': str(student.Дата_конца_обучения.tolist()[0]).split()[0],
               'dateOfApplication': str(student.Дата_приказа.tolist()[0]).split()[0],
               'orderNumber': student.Номер_приказа.tolist()[0]
               }

    doc.render(context)
    doc.save(f"{''.join(student.Имя.tolist())}.docx")



