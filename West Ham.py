from selenium import webdriver
from bs4 import BeautifulSoup
import os
from openpyxl import *
import pandas as pd
from selenium.webdriver.support import *






source_file = load_workbook('data.xlsx')
# In this line here we open the file in which our output will be submitted.

sheet = source_file["Sheet1"]
# here we are telling that we want our output in sheet1 of the xlsx file.

column=2
# this value and the row value will determine the cell value where the output will start writting in excel file.





options = webdriver.ChromeOptions()
# here we are defining the options or features which comes with the selenium.

options.add_argument("start-maximized")
# here we definig our first option which is we are maximizing the chrome browser window.

options.add_argument("--log-level=3")
# here we are telling the selenium to show only main warnings in our terminal when we run the code.

driver = webdriver.Chrome(options=options)

#here we are putting our option as a command or function in selenium's chrome class.

#driver is the main class which operates on as we define our options.






driver.get('https://www.whoscored.com/Teams/29/Show/England-West-Ham')
#here we are giving the website url which we want to scrape. now selenium will open it in chrome automatically as we know selenium is a headless browser.

page=driver.page_source
#here we are definig a variable which gets all the source of page like html content in text format.

soup=BeautifulSoup(str(page),'html.parser')
# here we are using giving out page source to beautiful soup in a string format so it can parse it using the html.parser or lxml parse and then we store it in the soup variable.





table=soup.find('table',{"id":'top-player-stats-summary-grid'})
# here we are looking for all the table tags in our html source which we build using beautifulsoup. we use find fuction as a singular to only finding the first or singal table of our html.



trs=table.find_all('tr')
# here we are looking for all the tr tags in our html or table variable that is why we are using the find_all() function. so we can find all the tr tags.
row=65
# this number is very important to represent the cell values of xlsx file

for tr in trs:
        # here we are looping over the tr tags.

        row=65

        tds=tr.find_all('td')
        # here we are finding all the td tags in every tr tag.


        for td in tds:
                #here we are looping over the td tags 
                print(td.text)
                #printing tags containig td





                sheet[chr(row)+str(column)]=td.text
                # here we are using the chr() to convert the two numerical values to character or a unicode value for cell like first number 65 will represent A and the second values reprsent the 2 they combined means A2 cell in excel sheet. 
                # sheet is defining that we are using the values unicode character as a cell value of sheet.

                row=row+1
                # by adding the 1 to every time after second loop end means it shift the cell and then the output value goes into second cell so it doesn't get overwritten.



        column=column+1
        # same as above happening here we are shifting the column value so we don't overwrite the values in same column.

source_file.save('data.xlsx')
# and here as you can guess we are saving the file after all the data is written.

