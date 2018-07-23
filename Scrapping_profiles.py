from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
import openpyxl


# Login Details Linked in
print('Linked in details, enter valid data for the script to work')
lnEmail = input('Enter your Linkedin Email Address:')
lnPassword = input('Enter your Linkedin password:')

# Login Details Gmail
print('Enter Gmail Details, enter valid data for the script to work')
gmailEmail = input('Enter your Gmail Email Address:')
gmailPassword = input('Enter Gmail your password:')

# List of email-address to fetch the data
print('Enter the email address to fetch data , please use SPACE to separate between two email-id ')
print('Example::: abc@gmail.com xyz@gmail.com lmn@yahoo.com uba@rediffmail.com')
testProfiles = input().split()

# Setting up the browser
op = webdriver.ChromeOptions()
op.add_extension('e.crx')
op.add_argument("--start-maximized")
path = r'C:\Users\ONEST\AppData\Local\Programs\Python\Python36-32\selenium\webdriver\chromedriver_win32\chromedriver.exe'
browser = webdriver.Chrome(path, chrome_options=op)

# Linkedin Login
url = 'https://www.linkedin.com/'
browser.get(url)

# Email and Password details for linked in profile
browser.find_element_by_class_name('login-email').send_keys(lnEmail)
browser.find_element_by_class_name('login-password').send_keys(lnPassword)
browser.find_element_by_css_selector('#login-submit').click()
print('Logged in to LinkedIn !')

# Gmail Login
time.sleep(3)
browser.execute_script("window.open('https://accounts.google.com/signin/v2/identifier?continue=https%3A%2F%2Fmail."
                       "google.com%2Fmail%2F&service=mail&sacu=1&rip=1&flowName=GlifWebSignIn&flowEntry=ServiceLogin')")

browser.switch_to_window(browser.window_handles[1])

# Writing the email-address to the input and clicking the next button -----
time.sleep(5)
browser.find_element_by_css_selector('#identifierId').send_keys(gmailEmail)
browser.find_element_by_css_selector('#identifierNext').click()
time.sleep(10)

# The Password Section and clicking the next button
browser.find_element_by_name('password').send_keys(gmailPassword)
time.sleep(5)
browser.find_element_by_css_selector('#passwordNext').click()
print('You are now Logged In to Gmail!')
time.sleep(12)

# Creating a Excel File
wb = openpyxl.Workbook()

# Entries in row one
sheet = wb.active
sheet['A1'] = "Name"
sheet['B1'] = "Designation"
sheet['C1'] = "Company / Student"
sheet['D1'] = "Location"
sheet['E1'] = "Account_link"
i = 2 # trace the value of the row

# Iterate over all the id's in the list testProfiles and fetching data
for val in testProfiles:
    # The compose button
    browser.find_element_by_class_name('z0').click()
    time.sleep(3)

    # Send to -----
    rec = browser.find_element_by_name('to')
    rec.send_keys(val)
    rec.send_keys(Keys.TAB)
    time.sleep(4)

    # Fetching details from the linked-in-navigator
    print('Fetching details !' + val)

    # Name
    try:
        name = browser.find_element_by_id('li-profile-name').text
        sheet[f'A{i}'] = name
    except:
        sheet[f'A{i}'] = "No details !"

    # Designation
    try:
        designation = browser.find_element_by_class_name('li-user-title-company').text
        sheet[f'B{i}'] = designation
    except:
        sheet[f'B{i}'] = "No details!"

    # Company
    try:
        company = browser.find_element_by_class_name('li-user-title').text
        sheet[f'C{i}'] = company
    except:
        sheet[f'C{i}'] = "No details!"

    # Location
    try:
        location = browser.find_element_by_class_name('li-user-location').text
        sheet[f'D{i}'] = location
    except:
        sheet[f'D{i}'] = "No details"

    # Account_link
    try:
        acc_link = browser.find_element_by_css_selector(
            '#li-header > div.li-user-profile-name > div > a').get_attribute('href')
        sheet[f'E{i}'] = acc_link
    except:
        sheet[f'E{i}'] = "No details!"

    time.sleep(2)
    browser.find_element_by_class_name('Ha').click()
    i += 1
    time.sleep(1)

browser.quit()
wb.save('profile_details.xlsx')
print('Done !')
# ----- End -----
