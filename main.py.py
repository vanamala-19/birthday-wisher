import pandas as pd
import datetime
import smtplib
import os
os.chdir(r"C:\\Users\\vanam\\OneDrive\\Desktop\\Python_Projects\\birthday_wisher")
# os.mkdir("testing") 

# Enter your authentication details
GMAIL_ID = 'user_email@gmail.com'
GMAIL_PSWD = 'user_password'

# funtion which sends the email
def sendEmail(to, sub, msg):
    print(f"Email to {to} sent with subject: {sub} and message {msg}" )
    s = smtplib.SMTP('smtp.gmail.com', 587)
    s.starttls()
    s.login(GMAIL_ID, GMAIL_PSWD)
    s.sendmail(GMAIL_ID, to, f"Subject: {sub}\n\n{msg}")
    s.quit()
    

if __name__ == "__main__":

    # to read the excel
    df = pd.read_excel("data.xlsx")

    # to get the current year and date
    today = datetime.datetime.now().strftime("%d-%m")
    yearNow = datetime.datetime.now().strftime("%Y")

    # index to stoge the year and  of wished person
    writeInd = []

    # to check wheather any one has a birthday
    for index, item in df.iterrows():
        bday = item['Birthday']
        if(today == bday) and yearNow not in str(item['Year']):

            if(item['Email'] != "not given"):
                sendEmail(item['Email'], "Happy Birthday", item['Dialogue']) 
            sendEmail("user_email@gmail.com", "Oye Its Birthday Of", item['Name'])
            writeInd.append(index)

    # for writing year in excel and to stop msg spam
    for i in writeInd:
        yr = df.loc[i, 'Year']
        df.loc[i, 'Year'] = str(yr) + ', ' + str(yearNow)

    df.to_excel('data.xlsx', index=False)   
