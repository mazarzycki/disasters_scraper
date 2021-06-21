from selenium import webdriver
import xlsxwriter
from datetime import datetime

now = (datetime.now()).strftime("%d-%m-%Y_%H-%M")

PATH = "C:\Program Files (x86)\chromedriver.exe"
driver = webdriver.Chrome(PATH)

workbook = xlsxwriter.Workbook("RSOE_" + now + ".xlsx")

worksheet = workbook.add_worksheet("EventList") 

#Open the website
driver.get("https://rsoe-edis.org/eventList")

#Take events list
articles = driver.find_elements_by_tag_name("tr")
row = 0
col = 0

# Loop through event list and unhide all event cards
event_cards = driver.find_elements_by_class_name("event-card")
for card in event_cards:
    driver.execute_script("arguments[0].removeAttribute(\"style\")", card)

for article in articles:
        
        
        header = article.find_element_by_class_name("title")
        date = article.find_element_by_class_name("eventDate")
        location = article.find_element_by_class_name("location")
        link = article.find_element_by_tag_name("a")  
        worksheet.write(row, col, header.text)
        worksheet.write(row, col + 1, date.text)
        worksheet.write(row, col + 2, location.text)
        worksheet.write(row, col + 3, link.get_attribute("href"))   

        print(header.text)

        row += 1      
workbook.close()      

driver.close()

print("Your file has been successfully created!")
