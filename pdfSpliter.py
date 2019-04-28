#program bioracy nazwe pliku ze schowka, przyjmujacy zakresy stron ktore ma wydzielic, zakres np 1-4, 7-8, zapisujacy poszczegolne zakresy jako osobne pliki i wysylajacy je na skrzynke
#zalezna od tego co wpisze, z mojej wlasnej pracowej skrzynki, dodawanie zakresow az q nie zostanie wprowadzone, dzialania na kazdym zakresie po kolei
#ostatnie strony: numPages - x i do konca
#czy sa strony ktore chcesz umiescic we wszystkich plikach? y/n jezeli y to podaje strone od ktorej wszystkie dalsze maja byc zalaczone
#regex do obcinanania znakow specjalnych: re.sub('[^A-Za-z0-9]+', '', string)
#patrzy ile ma inputow w liscie i tyle razy robi fora z indeksem wartosci z listy ile jest inputow
import PyPDF2
import os
import pyperclip
import re
import win32com.client as win32


os.chdir("C:\\Users\\RG053306\\Desktop\\Just in case")

country = input("What country invoice is issued to? UK, CA, US, PL, IE? ")

if country in ("UK", "US", "CA", "PL", "IE"):
    mailAddress = f"{country}APInvoices@jacobs.com"
else:
    print("You have entered something wrong.")

items1=[]
items2=[]
i=0
while True: #dodawanie numerow stron do list az uzytkownik nie wpisze q co przerywa petle
    i +=1
    item1 = input("Enter first page of the range %d: "%i)
    if item1 =="q":
        break
    item1 = int(item1)
    items1.append(item1)
    item2 = input("Enter last page of the range %d: "%i)
    if item2 =="q":
        break
    item2 = int(item2)  
    items2.append(item2)
#dodac zakresy w jakis sposob, gui
backups = int(input("How many pages needs to be added to each new file? Emails, backups, etc.: "))
originalFileName = pyperclip.paste()
workingfileName = originalFileName + ".pdf"
pdfFile = open(workingfileName, "rb")
strippedFileName = re.sub(r"[^a-zA-Z0-9]","",originalFileName) #regex usuwajacy znaki specjalne - do sprawdzenia
# print(inputs)
x = 0
for file in items1:
    # print(items1.index(file))
    # for x in range(0, len(items1)):

    reader = PyPDF2.PdfFileReader(pdfFile)
    writer = PyPDF2.PdfFileWriter()
    for pageNum in range(items1[items1.index(file)]-1, items2[items1.index(file)]):
        page = reader.getPage(pageNum)
        writer.addPage(page)
        x +=1

    for lastPage in range(reader.numPages - backups, reader.numPages):
        writer.addPage(reader.getPage(lastPage))
    

    newFileName = strippedFileName + str(items1.index(file)) + ".pdf"
    output = open(newFileName, "wb")
    writer.write(output)
    output.close()
    print(newFileName)
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = mailAddress
    mail.Subject = originalFileName
    mail.Body = ''
    attachment  = f"C:\\Users\\RG053306\\Desktop\\Just in case\\{newFileName}"
    mail.Attachments.Add(attachment)
    mail.Send()
pdfFile.close()
input()