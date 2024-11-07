from excel_handler import ExcelHandler
from gmail_automation import GmailAutomation

driver_path = '/home/mak/Documents/chromedriver-linux64/chromedriver'
excel_file = 'credentials.xlsx'       

excel = ExcelHandler(excel_file)
gmail = GmailAutomation(driver_path)

credentials = excel.read_credentials()
email_details = excel.read_email_details()

for idx, creds in enumerate(credentials, start=2):  
    success = gmail.login(creds['login_id'], creds['password'])
    
    if success:
        status, message = "Pass", "Login Successful"
        gmail.send_email(email_details['to'], email_details['subject'], email_details['body'], email_details['attachment_path'])
    else:
        status, message = "Fail", "Login Failed"
    
    excel.update_status(idx, status, message)

    if success:
        gmail.logout_and_close()
