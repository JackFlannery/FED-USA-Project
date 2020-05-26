# this program will login into the ACD inventory system, navigate to the inventory text file, and then make a new Excel file
# with the inventory information on it

# importing the "webdriver" module, which allows Selenium to work with any of the web drivers that are installed and
# in the executable PATH

from selenium import webdriver

# a function that makes it easier to find elements "By" a certain classification

from selenium.webdriver.common.by import By

# importing "os" and "time" module for later use

import os

import time

import dotenv

# importing the pandas library, which will allow text and csv files to be moved and altered.
# writing "as pd" means that the pandas library can be referred to as "pd" instead of "pandas". Saving letters :)

import pandas as pd

# naming a variable called options in which the options preferences will be stored for the web driver

# MAKE SURE BEFORE RUNNING THIS PROGRAM THAT CHROME DRIVER IS INSTALLED AND IN PATH WITH PYTHON!
# Chrome Driver Version 71 through 75 are to be utilized

options = webdriver.ChromeOptions()

# the preferences variable is storing the preferences that are to be changed (Download location)

preferences = {'download.default_directory': r'C:\Users\jackf\OneDrive - The Pennsylvania State University\Careers\FED USA'}

# this next code adds the preferences to the options variable

options.add_experimental_option("prefs", preferences)

# activating the web driver for Chrome along with the options

driver = webdriver.Chrome(options=options)

# creating a function that will do most of the work

def site_login():

    dotenv.load_dotenv()

    # Logging into the popup window. The username and password of my employer were placed in environment files in the same directory of the python file. 

    driver.get('ftp://' + os.getenv('USER') + ':' + os.getenv('PASS')+ '@ftp.acdd.com')

    # moving to the inventory page, where the "acd.txt" will be selected and downloaded to the desired folder.
    # the XPATH of the "acd.txt" file is used below
    driver.get('ftp://' + os.getenv('USER') + ':' + os.getenv('PASS')+ '@ftp.acdd.com/Inventory')
    driver.find_element(By.XPATH, '//*[@id="tbody"]/tr[1]/td[1]/a').click()

    # time.sleep() will keep the web driver window open long enough for the necessary file to be downloaded

    time.sleep(10)

    # after the allotted time, the driver window will close

    driver.quit()

# this function will fetch the date and time, and rename the downloaded file accordingly

def rename_file():

    # setting the time to a 12 hour clock, with the minute and seconds

    time_now = time.strftime("%I:%M:%S")

    # removing the colons from the time, as a file name does not accept colons

    time_no_colons = time_now[0:2] + '-' + time_now[3:5]

    # fetching the month, day and year

    date_now = time.strftime("%m/%d/%y")

    # removing the slashes from the date, as a file name does not accept slashes

    date_no_slash = date_now[0:2] + "-" + date_now[3:5] + "-" + date_now[6:8]

    # creating the name of the new file, with ACD Inventory followed by the date and the time

    file_name_change = 'ACD Inventory ' + str(date_no_slash) + ' ' + str(time_no_colons) + ".txt"

    # test, ignore -- print(file_name_change)

    # tying the old file to its directory and name

    old_file = os.path.join(r'C:\Users\jackf\OneDrive - The Pennsylvania State University\Careers\FED USA', 'acd.txt')

    # tying the new file to its directory and name

    new_file = os.path.join(r'C:\Users\jackf\OneDrive - The Pennsylvania State University\Careers\FED USA', file_name_change)

    # using the os module to rename the file

    os.rename(old_file, new_file)

    # pandas.read_csv will read a csv or text file, and collect the information of that file into a variable,
    # in this case "df". It uses a tab separator to divide up the columns

    df = pd.read_csv(new_file, sep="\t")

    # "variable".to_excel transfers the data stored in that variable to a given Excel file and sheet. index=False just
    # reorients how the data is lined up with its headers in the Excel spreadsheet

    df.to_excel(r'C:\Users\jackf\Desktop\output.xlsx', 'Sheet1', index=False)


# executing our functions

site_login()

time.sleep(10)

rename_file()