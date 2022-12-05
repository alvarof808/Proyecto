# doc.paragraphs[0].text[0]
# y.add_section(" d")
# print(doc.txt)

"""
for paragraph in doc.paragraphs:
    if 'Esto' in paragraph.text:
        print (paragraph.text)

        paragraph.text[0].replace('Esto', 'Unoo',10)
        print(paragraph.text[0])
 """
 
 # print(getText('1.docx'))
""" print(getText('2.docx'))
y = doc.paragraphs[0].text
print(type(doc.paragraphs[0].text))
y2.replace("Esto", "hola1")
y2 = str(y)
print(y2) """

""" # Funcion para obtener las primeras palabras de las paginas paginas
def getPaginas(pdfReader):
    lpagina =[]
    lpaginas= []
    IniPag = []
    for x in range(pdfReader.numPages):
        paginas = pdfReader.getPage(x)
        pagina = str(paginas.extractText())
        lpagina = pagina.split()
        lpaginas.append(lpagina)
        for m in lpagina:
            if m == " " or m == "":
                lpagina.remove(m)  
        for d in range(3):
            IniPag.append(lpagina[d])
            #print(lpagina[d])
    # Funcion para separar lista IniPag en espacios de 3
    #x = 3
    #final_list = lambda IniPag, x: [IniPag[i:i+x] for i in range(0, len(IniPag), x)]
    #output=final_list(IniPag, x)
    print(lpaginas[1])
    return lpaginas """
