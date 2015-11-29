from selenium import webdriver
import xlsxwriter

usn=[] #list to contain all usn's
no_of_students= 207

for i in range(1, no_of_students):
	if(i<10):
		string= '1RV13CS00'+str(i)
	elif(i>9 and i<100):
		string= '1RV13CS0'+str(i)
	else:
		string= '1RV13CS'+str(i)
	usn.append(string)

browser = webdriver.PhantomJS()

#URL of the main domain
url = 'http://www.rvce.edu.in/results/'
browser.get(url)

#Interacting with webpage via xpath
browser.find_element_by_xpath('//*[@id="ld"]/option[contains(text(), "Computer Science Engineering")]').click()
browser.find_element_by_xpath('//*[@id="resultview"]/td[4]/label/input').clear()
browser.find_element_by_xpath('//*[@id="resultview"]/td[4]/label/input').send_keys('4')
#Hard Coded for 4th Semester CSE

browser.find_element_by_name('Submit').click()

#Excel Workbook Name
workbook = xlsxwriter.Workbook('4thSem.xlsx')
worksheet = workbook.add_worksheet()
worksheet.set_column('A:A', 25)


for i in range(no_of_students-1):
	browser.find_element_by_name('usn').clear()
	browser.find_element_by_name('usn').send_keys(usn[i])

	browser.find_element_by_name('get_result').click()


	name=browser.find_element_by_xpath('/html/body/form/center/div[1]/font').text
	sgpa=browser.find_element_by_xpath('//*[@id="dataTablenew"]/tbody/tr[10]/td[2]/strong').text

	worksheet.write(i, 0, name)
	worksheet.write(i, 1, sgpa)
	


workbook.close()