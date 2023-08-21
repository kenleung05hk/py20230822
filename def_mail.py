import win32com.client as win32
from datetime import datetime, timedelta ,timezone
import math

timedata = datetime.now()
DDDD = timedata+timedelta(hours=1)
MMMM= math.ceil(int(DDDD.strftime("%M"))/15)*15
DDDD=DDDD.strftime("%H")

if MMMM == 60:
    MMMM = 0
    DDDD = int(DDDD)+1

def send_outlook_mail(to,attach_path="",client_name="",
                      currency=""):

    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.Subject = f'期货账户-{client_name} 货币兑换通知'
    mail.To = to
    try:
        mail.CC = "dealing@htfc.com.hk"
    except:
        ...
    mail.Display(False)
    mail.HTMLBody = f"""
     <font size="3">尊敬的客户-{client_name},</font>
    <br><br><br>
    阁下在本公司的期货账户-{client_name}由于货币結欠，需要兑换{currency}。本公司为了能够有效控制风险, 特此通知,  若客户未能于当天{DDDD}时{MMMM}分前通知本司为其进行换汇, 本司将有权根据客户协议书中的货币条款直接为客户进行该结算货币的兑换, 而不作另外通知。
    <br><br><br>
    谢谢客户的支持与合作。
    <br><br>
    如有其他问题，请与交易室联系。
    """ + mail.HTMLBody
    try:
        mail.Attachments.Add(str(attach_path))
    except:
        ...


if __name__ == "__main__":
    timedata = datetime.now()
    YYYY = timedata.strftime("%Y")
    MMMM = timedata.strftime("%m")
    DDDD = timedata.strftime("%d")
    YYYYMMMM = YYYY + '-' + MMMM
    numbertime = timedata.strftime('%Y%m%d')

    client_name="陈治国(A000695)"
    currency="HKFE_CNH 666.80"
    body_text = f"""
     <font size="3">尊敬的客户-{client_name},</font>
    <br><br><br>
    阁下在本公司的期货账户-{client_name}由于货币結欠，需要兑换{currency}。本公司为了能够有效控制风险, 特此通知,  若客户未能于当天13时00分前通知本司为其进行换汇, 本司将有权根据客户协议书中的货币条款直接为客户进行该结算货币的兑换, 而不作另外通知。
    <br><br>
    谢谢客户的支持与合作。
    <br>
    如有其他问题，请与交易室联系。
    """

    subject = f'期货账户-{client_name} 货币兑换通知'
    to = "am-ops@clsa.com" + ";" + "am-inv@clsa.com"
    cc = "dealing@htfc.com.hk"

    send_outlook_mail(body_text=body_text, subject=subject, to=to, cc=cc)