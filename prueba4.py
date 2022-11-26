import zipfile
from bs4 import BeautifulSoup
from xml.etree.ElementTree import parse, Element

doc_zip = zipfile.ZipFile("1.docx")

doc_xml = doc_zip.read("word/document.xml")



soup_xml = BeautifulSoup(doc_xml, "xml")
pretty_xml = soup_xml.prettify()

z =pretty_xml.replace('Esto','aquio', 1)
doc_xml.replace('Esto','aquio', 1)

print(type(pretty_xml))

hola = "hlalesejlñ"

hola+= '1'
print(hola)
print(z)