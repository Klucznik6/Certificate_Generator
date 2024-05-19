import os
##### scieżka szablonu = templatePath
##### ścieżka pliku xlsx = xlsxPath
##### ścieżka szablonów = templatesFolderPath
programPath = os.getcwd()
templatesFolderPath = programPath + '\\templates'
pdfLookup = programPath + '\\lookup\\szablon.png'
popplerPath = programPath + '\\poppler-23.07.0\\Library\\bin'

#####
senderEmail = ''
senderPassword = ''
receiverEmail = ''