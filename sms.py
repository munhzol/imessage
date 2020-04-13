import applescript
import xlrd 


loc = ("excel.xlsx") 

# To open Workbook 
wb = xlrd.open_workbook(loc) 
sheet = wb.sheet_by_index(0) 
  
# For row 0 and column 0 

def sendSms(sms,number):
	script = f"""tell application "Messages"
		send "{sms}" to buddy "{number}" of service "E:xxxx@icloud.com"
	end tell"""

	applescript.run(script)



sms = """test tesxt test test test spam"""

for i in range(2,102):
	if sheet.cell_value(i, 6) != '':
		number = str(int(sheet.cell_value(i, 6)))
		print(number)
		sendSms(sms,number)

		





