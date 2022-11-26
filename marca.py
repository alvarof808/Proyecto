import docx
import random
from random import sample
import re
from docx import Document


# Funcion para modificar el archivos .docx
def docx_replace_regex(doc_obj, regex , replace):
    for p in doc_obj.paragraphs:
        if regex.search(p.text):
            inline = p.runs
            # Loop added to work with runs (strings with same style)
            for i in range(len(inline)):
                if regex.search(inline[i].text):
                    text = regex.sub(replace, inline[i].text)
                    inline[i].text = text
                    
# Funcion de prueba                    
def docx_replace_regex2(doc_obj, regex , replace):
    #for p in doc_obj.paragraphs:
    #p = doc_obj.paragraphs[9]
    p = doc_obj
    if regex.search(p.text):
        inline = p.runs
        # Loop added to work with runs (strings with same style)
        for i in range(len(inline)):
            if regex.search(inline[i].text):
                text = regex.sub(replace, inline[i].text)
                inline[i].text = text
    

# Funcion para obtener todo el texto del documento .docx
def getText(filename):
    doc = docx.Document(filename)
    fullText = []
    for para in doc.paragraphs:
        fullText.append(para.text)
    return '\n'.join(fullText)

# Nro 1
# Funcion para obtener los parrafos independientes 
def getParrafo(filename):
    doc = docx.Document(filename)
    parrafo = []
    x = len(doc.paragraphs)
    for y in range(0, x):
        parrafo.append(doc.paragraphs[y].text)
    return parrafo  
  
# Nro 2
# Funcion para separar las palabras del parrafo
def getPalabras(listado):
    lista = []
    for parrafo in listado:
        lista.append(parrafo.split(sep=' '))
    return lista

# Nro 3
# Funcion para seleccionar todas las palabras con mas de 5 letras
def getMayoresCinco(lista):
    listaM = []
    for parrafo in lista:
        palabrasM = []
        for x in parrafo:
            if len(x) >= 5:
                palabrasM.append(x)
        listaM.append(palabrasM)
    return listaM

# Nro 4
# Funcion para eliminar Repetidos de una lista
def getUnicos(listado):
    listas = []
    for lista in listado:
        nueva_lista = []
        for elemento in lista:
# Este if verifica que el listado no tenga los siguientes carácteres ( ó ) ó ""
            if elemento.find(")") == -1 and elemento.find("(") == -1 and elemento.find('"')== -1:
            # Verificamos que no existan repetidos en lista
                if not elemento in nueva_lista:
                    nueva_lista.append(elemento)
        listas.append(nueva_lista)
    return listas


# Nro 5
# Funcion para selecionar las palabras
def getSeleccion(listas):
    lista_seleccion = []
    for listado in listas:
        seleccion = []
        if len(listado) >= 6:
            valor = sorted(sample(range(0 , len(listado)-1),3))
            for i in range(3):
                seleccion.append(listado[valor[i]])
        lista_seleccion.append(seleccion)
    return lista_seleccion

# Nro 6
# Funcion para modificar docx con valores seleccionados
def getMarca(lista_palabras, doc):
    for z in range(len(doc.paragraphs)):
        if len(lista_palabras[z]) == 3:
            for i in range(3):
                regex1 = re.compile(lista_palabras[z][i])
                nuevo = lista_palabras[z][i] + ' '
                docx_replace_regex2(doc.paragraphs[z], regex1 , nuevo)
    doc.save('7.docx')
                

doc = docx.Document('1.docx')
w  = getParrafo('1.docx') 
p = getPalabras(w)
#print(p)
l = getMayoresCinco(p)
u = getUnicos(l)     
o = getSeleccion(u)  
print(o) 

#print(getMayoresCinco(p))
#print(getSeleccion(l))

#l = getMayoresCinco(p)    
#filename = '1.docx'
#doc = Document(filename)
getMarca(o,doc)

""" 

print(regex1)
print(p)
print(l)
print(q)
print(len(l))
print(len(q))
print(o)

"""



