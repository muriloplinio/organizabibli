from openpyxl import load_workbook
import unicodedata


# Function definition is here
def TextToBibliArray(biblitext):
    
    Bibliarray=[]
    a=0
    ini_temp=0
    d=[]
    for c in biblitext:
        a = a+1
        if unicodedata.category(c) == "Lu":
           if ini_temp==0: 
               ini_temp=a
        if c== "," and ini_temp>0:
            d.append(ini_temp-1)
            ini_temp=0
        if unicodedata.category(c) != "Lu" and c!= ",":
            ini_temp=0
    d.append(len(biblitext))
    i = 1
    while i <= len(d):
        if i<len(d):
           Bibliarray.append(biblitext[d[i-1]:d[i]])
        i += 1
    return Bibliarray

#String exemple
s = "ANDRADE, Maria Margarida de. Introdução à metodologia do trabalho científico: elaboração de trabalhos na graduação, 10ª. São Paulo Atlas 2012 1 recurso online ISBN 9788522478392. Disponível em https://integrada.minhabiblioteca.com.br/books/9788522478392.  CALIJURI, Maria do Carmo. Engenharia Ambiental: conceitos, tecnologias e gestão. 2. Rio de Janeiro GEN LTC 2019 1 recurso online ISBN 9788595157446.  Disponível em https://integrada.minhabiblioteca.com.br/books/9788595157446.  HALLIDAY, David. Fundamentos de física, v.1 mecânica. 10. São Paulo LTC 2016 1 recurso online ISBN 9788521632054. Disponível em https://integrada.minhabiblioteca.com.br/reader/books/9788521632054. BEER, Ferdinand. Mecânica vetorial para engenheiros: Estática, v. 1. 11. Porto Alegre AMGH 2019 1 recurso online ISBN 9788580556209.  Disponível em https://integrada.minhabiblioteca.com.br/books/9788580556209.   MIRANDA, Shirley Aparecida de. Diversidade e ações afirmativas combatendo as desigualdades sociais. São Paulo Autêntica 2010 1 recurso online ISBN 9788582178157.Disponível em https://integrada.minhabiblioteca.com.br/books/9788582178157."
s2 = "ANDRADE, Maria Margariologia,  ISBN 9788522478392. CALIJURI, Maria GEN LTC 2019 1 recurso online ISBN 9788595157446."
caminho = 'base.xlsx'
arquivo_excel = load_workbook(caminho)
planilha1 = arquivo_excel.active
max_linha = planilha1.max_row
max_coluna = planilha1.max_column
max_linha = 6
max_coluna = 4
disccoll = 1
basicbiblicoll = 4
compbiblicoll = 5
#for i in range(1, max_linha + 1):
#    print(planilha1.cell(row=i, column=basicbiblicoll).value, end=" \n  ")
linha = 2
disciplina = planilha1.cell(row=linha, column=disccoll).value
basicbiblitxt = planilha1.cell(row=linha, column=basicbiblicoll).value
basicbiblitxt = basicbiblitxt.replace("\n","")
print(basicbiblitxt)
compbiblitxt = planilha1.cell(row=linha, column=compbiblicoll).value
compbiblitxt = compbiblitxt.replace("\n","")


print("Disciplina: ", disciplina, "\n")
Bibliarray2 = TextToBibliArray(basicbiblitxt)
i=0
while i < len(Bibliarray2):
    print(i,"-",Bibliarray2[i])
    i += 1