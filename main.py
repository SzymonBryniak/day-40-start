import pygsheets
import requests
from datetime import datetime
import smtplib


gc = pygsheets.authorize(client_secret='client_secret.json', scopes=['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'])


def get_user_details():
     users = []
     while True:
      first_name = input('Enter first name: ')
      last_name = input('Enter last name: ')
      email = input('Enter email address: ')
      users.append([first_name, last_name, email])
      question = input('Would you like to add more users? Y/N')
      if question == 'Y':
          continue
      else:
          print(users)
          return list(users)
      
def edit_pygsheet():
        users = get_user_details()
        print(users)
        header = ['First Name', 'Last Name', 'Email']
        try:
            sh2 = gc.open('users')
            wk1 = sh2[0]
            wk1.update_row(1, header)
        except pygsheets.exceptions.SpreadsheetNotFound:
            sh2 = gc.create('users')
            sh2 = gc.open('users')
            wk1 = sh2[0]
            wk1.update_row(1, header)

        wk1 = sh2[0]  # Open first worksheet of spreadsheet
        # Or
        # wks = sh.sheet1 # sheet1 is name of first worksheet
        print(wk1, sh2.url, wk1.get_value('A1'))
        # Append the new row at the end of the worksheet
        print(users)
        for i in range(0, len(users)):
            wk1.append_table(values=users[i])
        data = wk1.get_all_records()
        return data

def send_email_to_users():
    data = edit_pygsheet()
    print(data)
    user = 'szymonbryniakproject@gmail.com'
    password = 'psgw ndzo nnhm nylg'
    with smtplib.SMTP('smtp.gmail.com') as connection:
        connection.starttls()
        connection.login(user=user, password=password)
        connection.sendmail(msg=data, from_addr=user, to_addrs='oneplusszymonbryniak@gmail.com')

send_email_to_users()
