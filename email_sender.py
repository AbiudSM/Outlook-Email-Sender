import win32com.client as win32
import os
import datetime, time

CURRENT_PATH = os.getcwd()

months = ['meses','Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre']

# ? list [day, month, year]
now_time = datetime.datetime.now().strftime('%d-%m-%Y').split('-')

# ? date string in format '15-01-22' same as xlsx file
today = now_time[0] + '-' + now_time[1] + '-' + now_time[2][2:]

# ? string of the current month in spanish
current_month = months[int(now_time[1])]

# Email Parameters
contacts = 'example@gmail.com; second_example@live.com.mx'
subject = f'Plan de trabajo día {now_time[0]} de {current_month} del {now_time[2]} - Roberto Abiud Sánchez Montoya'

# Get the xlsx file name if it's from today's date
file_name = False
files_type = ('.xlsx','.xls')
for file in os.listdir(CURRENT_PATH):
    if file.startswith(today) and file.endswith(files_type):
        file_name = file


# Open outlook tool waiting 1 second for it to respond
outlook = win32.Dispatch('outlook.application')
time.sleep(1)    
mail = outlook.CreateItem(0)

# Set email parameters
mail.To = contacts
mail.Subject = subject

if file_name:

    # ? xlsx file path
    attachment  = CURRENT_PATH + '\\' + file_name

    mail.Attachments.Add(attachment)
    mail.Send()
    
    # Move xlsx file into month folder once sent
    time.sleep(1)
    month_path = CURRENT_PATH + '\\' + current_month
    if os.path.exists(month_path):
        os.replace(attachment, month_path + '\\' + file_name)
    else:
        os.mkdir(month_path)
        os.replace(attachment, month_path + '\\' + file_name)

# Display outlook window with email parameters if xlsx file was not found, to prevent wrong emails
else:
    mail.Display()