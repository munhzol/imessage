import applescript
import xlrd 


loc = ("Jurnal.xlsx") 

# To open Workbook 
wb = xlrd.open_workbook(loc) 
sheet = wb.sheet_by_index(0) 
  
# For row 0 and column 0 

def sendSms(sms,number):
	script = f"""tell application "Messages"
		send "{sms}" to buddy "{number}" of service "E:munhzol@icloud.com"
	end tell"""

	applescript.run(script)



sms = """Сайн байна уу
Коронавирусын (COVID-19) дэгдэлт өндөр байгаатай холбогдуулан хичээл тодорхойгүй хугацаагаар хойшилж буй тул онлайн сургалт явагдахаар болж байна. Эцэг эхчүүд та бүхнийг манай Facebook группт нэгдэж мэдээлэл авахыг урьж байна. 
Утас : 571 643 4534
Facebook group хаяг : <https://www.facebook.com/groups/349808615793817/> https://www.facebook.com/groups/349808615793817/"""

for i in range(2,102):
	if sheet.cell_value(i, 6) != '':
		number = str(int(sheet.cell_value(i, 6)))
		print(number)
		sendSms(sms,number)

		





