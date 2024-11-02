import pygsheets
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
          return list(users)
      
def get_data():
    try:
        spreadsheet_flights = gc.open('Flights')
        worksheet = spreadsheet_flights[0]
        data = worksheet.get_all_records()
    except pygsheets.exceptions.SpreadsheetNotFound:
        spreadsheet_flights = gc.create('Flights')
        worksheet = spreadsheet_flights[0]
        data = 'no flights'

    data_to_send = format_data(data)
    return data_to_send
    
def format_data(flight_data):
    formatted_data = "Flight Details:\n\n"
    for flight in flight_data:
        formatted_data += (
            f"From: {flight['FROM']} ({flight['CODE']})\n"
            f"To: {flight['TO']}\n"
            f"Departure: {flight['Departure']}\n"
            f"Return: {flight['Return']}\n"
            f"Price: ${flight['Price']}\n"
            "-----------------------------------\n"
        )
    return formatted_data


def edit_pygsheet():
        users = get_user_details()
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
    user_details = edit_pygsheet()
    print(user_details)
    data = get_data()
    user = 'szymonbryniakproject@gmail.com'
    password = 'psgw ndzo nnhm nylg'
    for i in user_details:
        with smtplib.SMTP('smtp.gmail.com') as connection:
            connection.starttls()
            connection.login(user=user, password=password)
            connection.sendmail(msg=data.encode('utf-8'), from_addr=user, to_addrs=i['Email'])

send_email_to_users()
