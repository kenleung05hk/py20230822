from datetime import datetime, timedelta

timedata=datetime.now() - timedelta(1)

insidetime=timedata.strftime('%d-%m-%Y')
outsidetime=timedata.strftime('%d %b %y')
ft1=timedata.strftime("%Y%m\%Y%m%d")
ft2=timedata.strftime("%d %b")
numbertime=timedata.strftime('%Y%m%d')

YYYY,MMMM,DDDD = [timedata.strftime("%Y"),timedata.strftime("%m"),timedata.strftime("%d")]


print (insidetime)
print (outsidetime)
print(ft1)
print(ft2)
print (numbertime)

print(YYYY)
print(MMMM)
print(DDDD)