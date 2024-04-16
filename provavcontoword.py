from docxtpl import DocxTemplate

doc = DocxTemplate('Prova_Template.docx')

# Prepare the data to replace the bookmark 
context = {
    'currency': "EURs"
}


print(context)

# Update the document content
doc.render(context)

# Save the changes
doc.save('updated_prova.docx') 
