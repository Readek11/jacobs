import openpyxl
import os
import re
import win32com.client as win32
os.chdir("C:\\Users\\RG053306\\Desktop\\dawdsda")
wb = openpyxl.load_workbook("LO_CH2M-UK.xlsx")
ws90 = wb["Sheet2"]
ws16 = wb["Sheet4"]


mailRegex = re.compile(r''' 
#name.surname@jacobs.com
[a-zA-Z0-9_\.]+ #name.surname
@               #@
[a-zA-Z0-9_\.]+  #jacobs.com
''', re.VERBOSE)
personRegex = re.compile(r'''
#surname, Name/(moze ale nie musi)(3 litery)

([a-zA-Z ]+ #surname
,)
([a-zA-Z ]+) #name
(/
[a-zA-Z ])? #kod biura
''', re.VERBOSE)
invoiceRegex = re.compile(r'''
([a-zA-Z0-9_\-/\. ]+) #invoice number and gibbersh 
(\d+\.\d\d) #invoice amount
''', re.VERBOSE)

#people90.append({"name": cell.value, "invoices90": [], "invoices16" = [], "email": []})
people = []
for row in ws90.iter_rows(min_row=1, max_col=2):
    for cell in row:
        if personRegex.search(str(cell.value)):
            people.append({"name": cell.value, "invoices90": [], "invoices16": [], "email": ""})
        if invoiceRegex.search(str(cell.value)):
            (people[-1])["invoices90"].append(cell.value)
        if mailRegex.search(str(cell.value)):
            (people[-1])["email"] = cell.value

for row in ws16.iter_rows(min_row=1, max_col=2):
    for cell in row:
        if personRegex.search(str(cell.value)):
            personExist = False
            for index, person in enumerate(people):
                if person["name"] == cell.value:
                    personExist = True
                    personID = index
                    break
            if personExist == False:
                people.append({"name": cell.value, "invoices90": [], "invoices16": [], "email": ""})
                personID = len(people) - 1
        if invoiceRegex.search(str(cell.value)):
            people[personID]["invoices16"].append(cell.value)
        if mailRegex.search(str(cell.value)):
            people[personID]["email"] = cell.value

newLine = "\n"
print(people[17])

for person in people: #po pipulach iterowac
    invoices90 = [] #to zdecydowanie mozna lepiej przeiterowac
    for invoice in person["invoices90"]:
        invoices90.append(invoice +"\n")
    invoices16 = []
    for invoice in person["invoices16"]:
        invoices16.append(invoice +"\n")


    if len(invoices90) > 0 and len(invoices16) > 0: #more then one old and more then one younger
        message = (
                f"""Dear {personRegex.search(str(person['name'])).group(2).strip()}\n\n"""
                f"""I am contacting you because you have {(len(invoices90) + len(invoices16))} invoices waiting for your action in Liquid Office.\nThose invoices wait to be approved already for more than 90 days. The oldest invoices in your inbox is listed below.\n\n"""
                f"""Please login to the Liquid Office system and take the appropriate action to clear those items as soon as possible. If you are not able to do that, or believe that the information below is incorrect, please let us know about that.\nOver 90 days:\n\n"""
                f"""{''.join(invoices90)}"""
                f"""{person["email"]}"""
                f"""\nOver 16 days:\n\n"""
                f"""{''.join(invoices16)}"""
                f"""\n\nIf you are on the Jacobs intranet the URL is http://liquidoffice.jacobs.com\n"""
                f"""If you are not on the Jacobs intranet the URL is http://connect.jacobs.com and use the Liquid Office link there.\n\n"""
                f"""If you have any technical issue preventing you from taking action on this invoice please contact jidsliquidoffice@jacobs.com.\n\n""" 
                f"""Regards,\n\nRadoslaw Gasior\nJacobs\nAccounting Professional | Accounts Payable\nradoslaw.gasior@jacobs.com""")   

    elif len(invoices90) == 1 and len(invoices16) == 0: #one old and no younger
        message = (

                f"""Dear {personRegex.search(str(person['name'])).group(2).strip()}\n\n"""
                f"""I am contacting you because you have 1 invoices waiting for your action in Liquid Office.\nThis invoice waits to be approved already for more than 90 days. The oldest invoice in your inbox is listed below.\n\n"""
                f"""Please login to the Liquid Office system and take the appropriate action to clear those items as soon as possible. If you are not able to do that, or believe that the information below is incorrect, please let us know about that.\nOver 90 days:\n\n"""
                f"""{''.join(invoices90)}"""
                f"""{person["email"]}"""
                f"""\n\nIf you are on the Jacobs intranet the URL is http://liquidoffice.jacobs.com\n"""
                f"""If you are not on the Jacobs intranet the URL is http://connect.jacobs.com and use the Liquid Office link there.\n\n"""
                f"""If you have any technical issue preventing you from taking action on this invoice please contact jidsliquidoffice@jacobs.com.\n\n""" 
                f"""Regards,\n\nRadoslaw Gasior\nJacobs\nAccounting Professional | Accounts Payable\nradoslaw.gasior@jacobs.com""")   

    elif len(invoices90) == 0 and len(invoices16) == 1: #no old and one younger
        message = (

                f"""Dear {personRegex.search(str(person['name'])).group(2).strip()}\n\n"""
                f"""I am contacting you because you have 1 invoices waiting for your action in Liquid Office.\nThis invoice waits to be approved already for more than 16 days. The oldest invoice in your inbox is listed below.\n\n"""
                f"""Please login to the Liquid Office system and take the appropriate action to clear this item as soon as possible. If you are not able to do that, or believe that the information below is incorrect, please let us know about that.\nOver 16 days:\n\n"""
                f"""{''.join(invoices16)}"""
                f"""{person["email"]}"""
                f"""\n\nIf you are on the Jacobs intranet the URL is http://liquidoffice.jacobs.com\n"""
                f"""If you are not on the Jacobs intranet the URL is http://connect.jacobs.com and use the Liquid Office link there.\n\n"""
                f"""If you have any technical issue preventing you from taking action on this invoice please contact jidsliquidoffice@jacobs.com.\n\n""" 
                f"""Regards,\n\nRadoslaw Gasior\nJacobs\nAccounting Professional | Accounts Payable\nradoslaw.gasior@jacobs.com""")   

    elif len(invoices90) > 1 and len(invoices16) == 0: #more then one old and no younger
        message = (

                f"""Dear {personRegex.search(str(person['name'])).group(2).strip()}\n\n"""
                f"""I am contacting you because you have {len(invoices90)} invoices waiting for your action in Liquid Office.\nThose invoices wait to be approved already for more than 90 days. The oldest invoices in your inbox is listed below.\n\n"""
                f"""Please login to the Liquid Office system and take the appropriate action to clear those items as soon as possible. If you are not able to do that, or believe that the information below is incorrect, please let us know about that.\nOver 90 days:\n\n"""
                f"""{''.join(invoices90)}"""
                f"""{person["email"]}"""
                f"""\n\nIf you are on the Jacobs intranet the URL is http://liquidoffice.jacobs.com\n"""
                f"""If you are not on the Jacobs intranet the URL is http://connect.jacobs.com and use the Liquid Office link there.\n\n"""
                f"""If you have any technical issue preventing you from taking action on this invoice please contact jidsliquidoffice@jacobs.com.\n\n""" 
                f"""Regards,\n\nRadoslaw Gasior\nJacobs\nAccounting Professional | Accounts Payable\nradoslaw.gasior@jacobs.com""")   

    elif len(invoices90) == 0 and len(invoices16) > 1: #more then one young and no older
        message = (

                f"""Dear {personRegex.search(str(person['name'])).group(2).strip()}\n\n"""
                f"""I am contacting you because you have {len(invoices16)} invoices waiting for your action in Liquid Office.\nThose invoices wait to be approved already for more than 16 days. The oldest invoices in your inbox is listed below.\n\n"""
                f"""Please login to the Liquid Office system and take the appropriate action to clear those items as soon as possible. If you are not able to do that, or believe that the information below is incorrect, please let us know about that.\nOver 16 days:\n\n"""
                f"""{''.join(invoices16)}"""
                f"""{person["email"]}"""
                f"""\n\nIf you are on the Jacobs intranet the URL is http://liquidoffice.jacobs.com\n"""
                f"""If you are not on the Jacobs intranet the URL is http://connect.jacobs.com and use the Liquid Office link there.\n\n"""
                f"""If you have any technical issue preventing you from taking action on this invoice please contact jidsliquidoffice@jacobs.com.\n\n""" 
                f"""Regards,\n\nRadoslaw Gasior\nJacobs\nAccounting Professional | Accounts Payable\nradoslaw.gasior@jacobs.com""")   


    print(message)

# else:
#     message16 = (
#         f"""Dear {personRegex.search(str(person['name'])).group(2).strip()}\n\n"""
#         f"""I am contacting you because you have {len(invoices16)} invoices waiting for your action in Liquid Office.\nThose invoices wait to be approved already for more than 16 days. The oldest invoices in your inbox are listed below.\n\n"""
#         f"""Please login to the Liquid Office system and take the appropriate action to clear those items as soon as possible. If you are not able to do that, or believe that the information below is incorrect, please let us know about that.\nOver 16 days:\n\n"""
#         f"""{''.join(invoices16)}"""
#         f"""\n\nIf you are on the Jacobs intranet the URL is http://liquidoffice.jacobs.com\n"""
#         f"""If you are not on the Jacobs intranet the URL is http://connect.jacobs.com and use the Liquid Office link there.\n\n"""
#         f"""If you have any technical issue preventing you from taking action on those invoices please contact jidsliquidoffice@jacobs.com.\n\n""" 
#         f"""Regards,\n\nRadoslaw Gasior\nJacobs\nAccounting Professional | Accounts Payable\nradoslaw.gasior@jacobs.com""")



# if (len(invoices90) + len(invoices16)) == 1:
#     message90 = (
#         f"""Dear {personRegex.search(str(person['name'])).group(2).strip()}\n\n"""
#         f"""I am contacting you because you have {(len(invoices90) + len(invoices16))} invoice waiting for your action in Liquid Office.\nThis invoice waits to be approved already for more than 90 days. The oldest invoice in your inbox is listed below.\n\n"""
#         f"""Please login to the Liquid Office system and take the appropriate action to clear this item as soon as possible. If you are not able to do that, or believe that the information below is incorrect, please let us know about that.\nOver 90 days:\n\n"""
#         f"""{''.join(invoices90)}"""
#         f"""{''.join(invoices16)}"""
#         f"""\n\nIf you are on the Jacobs intranet the URL is http://liquidoffice.jacobs.com\n"""
#         f"""If you are not on the Jacobs intranet the URL is http://connect.jacobs.com and use the Liquid Office link there.\n\n"""
#         f"""If you have any technical issue preventing you from taking action on this invoice please contact jidsliquidoffice@jacobs.com.\n\n""" 
#         f"""Regards,\n\nRadoslaw Gasior\nJacobs\nAccounting Professional | Accounts Payable\nradoslaw.gasior@jacobs.com""")   

# else:
#     message90 = (
#         f"""Dear {personRegex.search(str(person['name'])).group(2).strip()}\n\n"""
#         f"""I am contacting you because you have {(len(invoices90) + len(invoices16))} invoices waiting for your action in Liquid Office.\nThose invoices wait to be approved already for more than 90 days. The oldest invoices in your inbox are listed below.\n\n"""
#         f"""Please login to the Liquid Office system and take the appropriate action to clear those items as soon as possible. If you are not able to do that, or believe that the information below is incorrect, please let us know about that.\nOver 90 days:\n\n"""
#         f"""{''.join(invoices90)}"""

#         f"""{''.join(invoices16)}"""
#         f"""\n\nIf you are on the Jacobs intranet the URL is http://liquidoffice.jacobs.com\n"""
#         f"""If you are not on the Jacobs intranet the URL is http://connect.jacobs.com and use the Liquid Office link there.\n\n"""
#         f"""If you have any technical issue preventing you from taking action on those invoices please contact jidsliquidoffice@jacobs.com.\n\n""" 
#         f"""Regards,\n\nRadoslaw Gasior\nJacobs\nAccounting Professional | Accounts Payable\nradoslaw.gasior@jacobs.com""")
#     print(message90)
    # outlook = win32.Dispatch("outlook.application")
    # mail = outlook.CreateItem(0)
    # mail.To = "radoslaw.gasior@jacobs.com"
    # mail.Subject = "ACTION REQUIRED: Outstanding invoices in Liquid Office"
    # mail.Body = message
    # mail.Send()

# for person16 in people:
#     invoices16 = [] #to zdecydowanie mozna lepiej przeiterowac
#     for invoice in person16["invoices16"]:
#         invoices16.append(invoice +"\n")


#     print(message16)

    # outlook = win32.Dispatch("outlook.application")
    # mail = outlook.CreateItem(0)
    # mail.To = "radoslaw.gasior@jacobs.com"
    # mail.Subject = "ACTION REQUIRED: Outstanding invoices in Liquid Office"
    # mail.Body = message
    # mail.Send()

    # print(f"Dear {personRegex.search(str(person['name'])).group(2).strip()}{newLine}{newLine}You have {' '.join(invoices)}{newLine}{newLine}")


input()

#jezeli A1 spelnia kryteria to dodaj A1 jako dictionary key, nastepne wartosci
#w kolumnie A jako wartosci tego klucza (jako lista?) az kolejna wartosc nie spelni personRegex
#wtedy sprawdz czy C1 jest mailem ktory spelnia warunki i jezeli tak to dodaj
#go do listy.
#Maile beda przygotowywane jako Dear {wyciagniete imie z regex}, masz x faktur
#w inboxie, zajmij sie nimi, lista faktur powyzej 90 dni i lista faktur do 90 dni

#if jezeli kolejka to osobna lista!!!!!!!!!!!


#for num in range (1,131): # liczba wierszy
#



#pusta liste people w forze if z regexem od osoby, jezeli spelni warunek ze to osoba
#{name: czlowiek, faktury: [faktura1, faktura2], mail: mail@mail.com}
# jezeli znajdujesz goscia to tworzysz slownik, jezeli znajdujesz fakture to dodaj listy

        # if personRegex.search(str(cell.value)):
            

    # people.append([(personRegex.search(str(cell.value)).group()) for cell in row if personRegex.search(str(cell.value))])
    # people.append([(personRegex.search(str(cell.value)).group()) else (invoiceRegex.search(str(cell.value)).group()) for cell in row])?????????


        # if personRegex.search(str(cell.value)):
        #     print(cell.value)
            # print(personRegex.search(str(cell.value)).group(2).strip()) #zwraca imie danej osoby

#wyciaganie faktur
# for row in ws.iter_rows(min_row=1, min_col=1, max_col=1, max_row=131):
#     for cell in row:
#         if invoiceRegex.search(str(cell.value)):
#             # invoices.append(cell.value)
#             print(cell.value)   #zwraca fakture
#         # else:
#         #     print(cell.value) #zwraca człowieka lub kolejke
# #wyciaganie maili
# # mails = []
# for row in ws.iter_rows(min_row=1, max_col=3, max_row=131):
#     for cell in row:
#         if mailRegex.search(str(cell.value)):
#             print(cell.value)
# #             mails.append(mailRegex.search(str(cell.value)).group())
# print(mails)

#
