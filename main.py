from docx import Document
from docx.shared import Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH

def show_styles():
    document = Document()
    all_styles = document.styles
    for style in all_styles:
        print(style.name)


def my_styles(document):
    section = document.sections[0]
    section.top_margin    = Cm(2)
    section.right_margin  = Cm(1.5)
    section.bottom_margin = Cm(2)
    section.left_margin   = Cm(3)

    styleNormal = document.styles["Normal"]
    styleNormal.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    styleNormal.paragraph_format.first_line_indent = Cm(1.25)
    styleNormal.paragraph_format.left_indent = 0
    styleNormal.paragraph_format.line_spacing = 1.5
    styleNormal.paragraph_format.right_indent = 0
    styleNormal.paragraph_format.space_after = 0
    styleNormal.paragraph_format.space_before = 0
    

    return document


def introduction(document):
    # document.add_heading('Сделать титульную страницу', level=1)

    # document.add_page_break()

    document.add_heading("Введение", level=1)


    document.add_paragraph("Flex-box - это технология CSS (для создания сайта), которая позволяет легко и гибко управлять расположением элементов на веб-странице. Она позволяет создавать адаптивные макеты, которые легко адаптируются к различным размерам экранов и устройств.")
    document.add_paragraph("Проект предназначен для людей, только начинающих свой путь в frontend разработке. Этот материал поможет, как новичкам начать свое обучение, так и опытным разработчикам повторить ранее изученный материал.")
    document.add_paragraph("Актуальностью проекта для команды является возможность узнать подробнее о различных свойствах Flex-box и улучшить свои навыки по работе с данной технологией. Проект будет нести в себе не только теоретические знания, но и возможность использования их на практике.")
    document.add_paragraph("Цель проекта - создать обучающую игру для людей, которые хотят грамотно и быстро научиться технологии Flex-box.")
    document.add_paragraph("В связи с поставленной целью, необходимо решить следующие задачи:")
    

    document.add_paragraph(style="List Number", text="Выбрать тему проекта;")
    document.add_paragraph(style="List Number", text="Распределиться по группам;")
    document.add_paragraph(style="List Number", text="Определить роли в группе;")
    document.add_paragraph(style="List Number", text="Установить рабочее расписание:")
    document.add_paragraph(style="List Number 2", text="Распределить работу между участниками группы;")
    document.add_paragraph(style="List Number 2", text="Поставить ограничение по времени для выполнения рабочих задач.")
    document.add_paragraph(style="List Number", text="Выбрать источники информации;")
    document.add_paragraph(style="List Number", text="Выбрать инструменты, с помощью которых будет создаваться обучающая игра:")
    document.add_paragraph(style="List Number 2", text="Текстовый редактор;")
    document.add_paragraph(style="List Number 2", text="Фоторедакторы;")
    document.add_paragraph(style="List Number 2", text="Приложения для создания кода;")
    document.add_paragraph(style="List Number 2", text="Общие приложения для работы команды.")
    document.add_paragraph(style="List Number", text="Создать дизайн игры:")
    document.add_paragraph(style="List Number 2", text="Подобрать текстуры и дизайн;")
    document.add_paragraph(style="List Number", text="Написать практическую часть с помощью дополнительных источников информации;")
    document.add_paragraph(style="List Number", text="Реализовать сайт с помощью ранее выполненных задач;")
    document.add_paragraph(style="List Number", text="Ввести конечные правки;")
    document.add_paragraph(style="List Number", text="Защитить аналитический отчет.")


    document.add_paragraph('Объектом исследования в данном проекте является технология Flex-box.')

    document.add_paragraph('Предметом исследования является обучающая игра, основанная на использовании технологии Flex-box.')

    document.add_paragraph('Субъект исследования: проектная команда.')

    document.add_paragraph('Методы исследования, которые используются в проекте: Кабинетный, разведочный, описательный, моделирование и метод экспертных оценок.')


    # document.add_page_break()

    # document.add_heading("Глава 1. Натуральное описание", level=1)



    # document.add_paragraph('ААААААААА')

    # document.add_paragraph('FFFFFFFFF', style='List Number')
    # document.add_paragraph('FFFFFFFFF', style='List Number 2')

    # document.add_paragraph(
    #     'first item in unordered list', style='List Bullet'
    # )
    # document.add_paragraph(
    #     'first item in ordered list', style='List Number'
    # )

    # document.add_picture('./media/some-image.jpg', width=Inches(1))

    # records = (
    #     (3, '101', 'Spam'),
    #     (7, '422', 'Eggs'),
    #     (4, '631', 'Spam, spam, eggs, and spam')
    # )




    return document





def main():
    document = Document()
    document = my_styles(document)
    document = introduction(document)
    document.save('target.docx')



if __name__ == '__main__':
    main()