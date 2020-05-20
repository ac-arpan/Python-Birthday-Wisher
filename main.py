import pandas as pd
import datetime
import smtplib
import os

os.chdir(r"C:\Users\DELL\Desktop\Python\Mini Proj\New folder")
os.mkdir("Test")

# Enter the authentication details
GMAIL_ID = 'arpanchowdhury050@gmail.com'
GMAIL_PSWD = 'google123#'

def sendEmail(to,sub,msg):
    # print(f'Email to {to} send with subject : {sub} and message {msg}')
    s = smtplib.SMTP('smtp.gmail.com',587)
    s.starttls()
    s.login(GMAIL_ID,GMAIL_PSWD)

    s.sendmail(GMAIL_ID,to,f'{sub} \n\n {msg}')
    s.quit()

if __name__ == "__main__":
    # read the csv in the dataframe df
    df = pd.read_excel('data.xlsx')

    # get today's date in date-month format
    today = datetime.datetime.now().strftime("%d-%m")

    #get current year
    yearNow = datetime.datetime.now().strftime("%Y")

    #this array is used to append the wished year to the csv
    writeInd = []

    for index, item in df.iterrows():
        bday = item['Birthday'].strftime("%d-%m")
        
        if today == bday and yearNow not in str(item['Year']):
            sendEmail(item['Email'],"Happy Birthday to You",item['Dialogue'])
            writeInd.append(index)

    if len(writeInd) != 0:
        for i in writeInd:
            #getting the specific [row,column] value from the dataframe
            yr = df.loc[i,'Year']

            #adding the wished year 
            df.loc[i,'Year'] = str(yr) + ',' + str(yearNow)

        #writing the data frame to the excel sheet
        # print(df)
        df.to_excel('data.xlsx',index = False)
        

