import requests, openpyxl
from bs4 import BeautifulSoup
from time import sleep
import smtplib
from email.message import EmailMessage

Email_Address = "MY EMAIL HERE"
Email_Password = "MY PASSWORD HERE"

wb = openpyxl.load_workbook('Links.xlsx')               #This connected with the excel file
sheet = wb.active

#This function will add a new item to the database
def add_item():
    link = input("Enter the link of the item: ")
    expected_price = int(input("Add the expected price below which you would like to receive a notification mail: "))

    sheet.append((link,expected_price))                 #Adding link and price to the database(excel)
    wb.save("Links.xlsx")
    print("Item added. \n")

#This function will check prices of all the items in the database and will send a mail to the user if the price is below than the user's expected price
def check_price():

    for i in range(3,sheet.max_row+1):                              #Looping through all the items
        link = sheet.cell(row=i,column=1).value
        expected_price = sheet.cell(row=i,column=2).value

        current_price = 0
        headers = {"User-Agent": "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.163 Safari/537.36"}
        try:
            page = requests.get(link, headers = headers)
            soup = BeautifulSoup(page.content, 'html.parser')
            title= (soup.find(id="productTitle").get_text()).strip()                                                #Getting title from the amazon link
            current_price = int((soup.find(id="priceblock_ourprice").get_text())[2:].split('.')[0].replace(",","")) #Getting price and converting it to integer after removing , and rs symbol
        except:
            print("\nEither the item doesn't exists or it isn't available right now\n")    
        
        if(current_price == 0):
            pass
        elif(int(current_price) < expected_price):
            send_mail(title, current_price)
        else:
            print(f"\nThe price of {title} is still above your expected value at {current_price}\n")
        
        if(i != sheet.max_row):
            for i in range(30,0,-10):                                                           #Waiting for 30 seconds before checking price for next item to avoid continuous load on amazon
                print(f"Waiting for {i} seconds before checking for next item...")
                sleep(10)

def send_mail(title, current_price):
    print(f"YIPEE!! Price of {title} is reduced to {current_price} now. We have sent a mail to the user as well.\n")
    msg = EmailMessage()
    msg['Subject'] = "Price Decreased for your amazon item"
    msg['From'] = Email_Address
    msg['To'] = Email_Address
    msg.set_content(f"Hi,\nWe are glad to inform you that the price for {title} has now been reduced to {current_price}.")

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(Email_Address, Email_Password)
        smtp.send_message(msg)
        
#This method will delete an item from the database
def del_item():
    link = input("Enter the link of the item you want to delete: ")
    del_row=0                                                   #Variable to store the row to be deleted
    for i in range(2,sheet.max_row+1):
        if(sheet.cell(row=i,column=1).value==link):             #Searching if the entry is present in the data or not
            del_row = i
            break

    if del_row==0:                                                 
        return("No such entry found")                              
    else:
        sheet.delete_rows(del_row)
        wb.save("Links.xlsx")
        print("Entry deleted!\n\n")






def email_update():
    email = input("Enter the new email address: ")
    sheet['A1'] = email
    wb.save("Links.xlsx")
    print("Email address updated\n\n")





def main():
    print('''
                                            _______             _             
     /\                                    |__   __|           | |            
    /  \   _ __ ___   __ _ _______  _ __      | |_ __ __ _  ___| | _____ _ __ 
   / /\ \ | '_ ` _ \ / _` |_  / _ \| '_ \     | | '__/ _` |/ __| |/ / _ \ '__|
  / ____ \| | | | | | (_| |/ / (_) | | | |    | | | | (_| | (__|   <  __/ |   
 /_/    \_\_| |_| |_|\__,_/___\___/|_| |_|    |_|_|  \__,_|\___|_|\_\___|_|   
                                                                              
''')                                                                              

    
    if(sheet['A1'].value is None) :                                          #Checking if an email address of the user already exists 
        email = input("Enter the email address where you would like to receive the notifiaction. (This will be asked just once): ")
        sheet['A1'] = email
        wb.save("Links.xlsx")
        print("Email address added\n")
    
    while(True):
        print("What would you like to do?\n1.Add a new item\n2.Check for prices\n3.Delete an item\n4.Update email address\n9.To quit\n\n")

        choice = input("Pick an option: ")

        if choice == '1':
            add_item()
        elif choice =='2':
            check_price()
        elif choice =='3':
            del_item()
        elif choice == '4':
            email_update()
        else:
            print("Existing...")
            break


if __name__ == '__main__':
    main()