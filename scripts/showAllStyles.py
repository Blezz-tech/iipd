from docx import Document


document = Document()
all_styles = document.styles
for style in all_styles:
    print(style.name)