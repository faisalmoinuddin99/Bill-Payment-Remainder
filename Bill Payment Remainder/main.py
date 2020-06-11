import pandas as pd
import datetime
import smtplib

GMAIL_ID = 'faisal25marcg@gmail.com'
GMAIL_PWD = 'messifacebook'

def sendEmail(to,sub,msg):
    print(f"Email to: {to}, sent with subject: {sub}, Message: {msg}")
    s = smtplib.SMTP('smtp.gmail.com',587)
    s.starttls()
    s.login(GMAIL_ID, GMAIL_PWD)
    s.sendmail(GMAIL_ID, to, f"Subject: {sub}\n\n {msg}")
    s.quit()

df = pd.read_excel("MyData.xlsx")
# print(df)
today = datetime.datetime.now().strftime("%d-%m")
# print(today)
yearNow = datetime.datetime.now().strftime("%Y")
monthNow = datetime.datetime.now().strftime("%m")
# print(monthNow)

writeIndex = []
for index,item in df.iterrows():
    # print(index,item["DUEDATE"])
    dueDate = item["DUEDATE"].strftime("%d-%m")
    # print(dueDate)
    if today == dueDate and monthNow not in str(item["MONTH"]) and yearNow not in str(item["YEAR"]):
        sendEmail(item["EMAIL"],item["DATAITEM"],item["MESSAGE"])
        writeIndex.append(index)
for i in writeIndex:
    mon = df.loc[i,"MONTH"]
    yr = df.loc[i,"YEAR"]
    df[i,"MONTH"] = f"{mon},{monthNow}"
    df[i,"YEAR"] = f"{yr},{yearNow}"
df.to_excel('MyData.xlsx',index=False)


