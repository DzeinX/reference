import pandas as pd
from docxtpl import DocxTemplate

xl = pd.read_excel('spravka.xlsx')


print('Введите номер ученика: ', end='')
num_st = int(input())

student = xl[xl.index == num_st]

doc = DocxTemplate("spravka_shablon.docx")

context = {'name': ''.join(student.Имя.tolist()),
           'dateOfBirth': student.Дата_рождения.tolist()[0],
           'course': student.Курс.tolist()[0],
           'formOfEd': ''.join(student.Форма_обучения.tolist()),
           'edProgram': ''.join(student.Образовательная_программа.tolist()),
           'directionOfTraining': student.Направление_подготовки.tolist()[0],
           'nameOfUniversity': ''.join(student.Институт.tolist()),
           'dateOfStart': student.Дата_начала_обучения.tolist()[0],
           'dateOfEnd': student.Дата_конца_обучения.tolist()[0],
           'dateOfApplication': student.Дата_приказа.tolist()[0],
           'orderNumber': student.Номер_приказа.tolist()[0]
           }

doc.render(context)
doc.save("Справка_обучающегося.docx")



