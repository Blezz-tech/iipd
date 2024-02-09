from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.style import WD_STYLE_TYPE
import os

# Жоско для решения проблемы (см Пояснение_за_код.md)
from docx.oxml.ns import qn



def my_styles(document):
    # Макет
    section = document.sections[0]
    section.top_margin    = Cm(2)
    section.right_margin  = Cm(1.5)
    section.bottom_margin = Cm(2)
    section.left_margin   = Cm(3)

    styles = document.styles

    for style in styles:
        if hasattr(style, 'paragraph_format'):
            print("name:", style.name)
            # print("paragraph: ", hasattr(style, 'paragraph_format'))
            # pf = style.paragraph_format
            # pf.space_after =  Pt(0)
            # pf.space_before =  Pt(0)

        match style.name:
            # Обычный текст
            case "Normal":
                pf = style.paragraph_format
                font = style.font

                pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                pf.first_line_indent = Cm(1.25)
                pf.left_indent = 0
                pf.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
                pf.line_spacing = 1.5
                pf.right_indent = 0
                pf.space_after =  Pt(0)
                pf.space_before =  Pt(0)
                
                font.name = "Times New Roman"
                font.size = Pt(14)


            # Основной текст
            case "Body Text":
                pf = style.paragraph_format

                pf.space_after =  Pt(0)
                pf.space_before =  Pt(0)
            
            
            # Заголовок 1
            case "Heading 1":

                # Жоско для решения проблемы (см Пояснение_за_код.md)
                rFonts = style.element.rPr.rFonts
                rFonts.set(qn("w:asciiTheme"), "Times New Roman")
                rFonts.set(qn("w:eastAsiaTheme"), "Times New Roman")
                rFonts.set(qn("w:hAnsiTheme"), "Times New Roman")    # Из всех 4 помагло именно это, но применяю все 4, чтобы наверняка
                rFonts.set(qn("w:cstheme"), "Times New Roman")
                # Конец Жоскости

                pf = style.paragraph_format
                font = style.font

                pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
                pf.first_line_indent = Cm(1.25)
                pf.left_indent = 0
                pf.line_spacing = 1.5
                pf.right_indent = 0
                pf.space_after = 0
                pf.space_before = 0
                
                font.size = Pt(16)
                font.color.rgb = RGBColor(0,0,0)
                font.name = "Times New Roman" # Не работает, Обращаться к Ожскости


            # Compact
            case "Heading 1":
                pf = style.paragraph_format
                font = style.font

                # pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
                pf.first_line_indent = Cm(1.25)
                pf.left_indent = 0
                pf.line_spacing = 1.5
                pf.right_indent = 0
                pf.space_after = 0
                pf.space_before = 0
                pf.tab_stops.add_tab_stop(Cm(2.25))
                
                font.name = "Times New Roman"
                font.size = Pt(14)
                font.color.rgb = RGBColor(0,0,0)


    # List Number
    # if True:
    #     ListNumber1 = styles["List Number"]

    #     pf = ListNumber1.paragraph_format
    #     font = ListNumber1.font

    #     # pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
    #     pf.first_line_indent = Cm(1.25)
    #     pf.left_indent = 0
    #     pf.line_spacing = 1.5
    #     pf.right_indent = 0
    #     pf.space_after = 0
    #     pf.space_before = 0
    #     pf.tab_stops.add_tab_stop(Cm(2.25))

    #     font.name = "Times New Roman"
    #     font.size = Pt(14)
    #     font.color.rgb = RGBColor(0,0,0)


    # List Number 2
    # if True:
    #     ListNumber2 = styles["List Number 2"]

    #     pf = ListNumber2.paragraph_format
    #     font = ListNumber2.font

    #     # pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
    #     pf.first_line_indent = Cm(1.75)
    #     pf.left_indent = 0
    #     pf.line_spacing = 1.5
    #     pf.right_indent = 0
    #     pf.space_after = 0
    #     pf.space_before = 0
    #     pf.tab_stops.add_tab_stop(Cm(2.75))

    #     font.name = "Times New Roman"
    #     font.size = Pt(14)
    #     font.color.rgb = RGBColor(0,0,0)


    # Image Caption


    return document


def main():
    os.system("nu generate.nu")
    document = Document('target/source.docx')

    document = my_styles(document)
    document.save('target/Аналитический_отчет.docx')



if __name__ == '__main__':
    main()