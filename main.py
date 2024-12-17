#!/usr/bin/env python
# coding: utf-8

####Import relevant packages####

from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException
from selenium import webdriver
from dotenv import load_dotenv

import configparser

import pandas as pd
import os
import time

# Assign values from config file

config = configparser.ConfigParser()
config.read("config.ini")  # Make sure config.ini exists
admit_term = config["DEFAULT"]["admit_term"]
program_filter = config["DEFAULT"]["program_filter"]
excel_file = config["DEFAULT"]["excel_file"]

print(f"Default Admit Term: {admit_term}")
print(f"Default Program Filter: {program_filter}")
print(f"Default Excel File: {excel_file}")

###User Inputs##

admit_term = input("Enter the admit term (e.g., 1258): ").strip()
program_filter = input("Enter the program filter (e.g., 'cl', 'sp', 'cd'): ").strip()
excel_file = input("Enter the Excel file name (e.g., SelectionWorksheet.xlsx): ").strip()


####Set the Default Download Directory####

#Save the original working directory
base_dir = os.path.dirname(os.path.abspath(__file__))

# Set the default download location to a new "Applicants" folder
download_directory = os.path.join(base_dir, "Applicants")

# Create the folder if it doesn't exist
if not os.path.exists(download_directory):
    os.makedirs(download_directory)

# Set the download directory
os.chdir(download_directory)
print(f"Download directory set to: {download_directory}")


# Configure Chrome preferences to automatically download PDFs to that directory
prefs = {
    "download.default_directory": download_directory,
    "download.prompt_for_download": False,
    "plugins.always_open_pdf_externally": True
}

chrome_options = Options()
chrome_options.add_experimental_option("prefs", prefs)


####Create a Driver for Chrome####

# Set the path to your chromedriver
driver_path = os.path.join(base_dir, "chromedriver.exe")
if not os.path.exists(driver_path):
    print("ERROR: ChromeDriver not found in the script directory.")
    print(f"Expected path: {driver_path}")
    print("Please download the appropriate ChromeDriver version and place it in the script folder.")
    exit(1)  # Exit the script with an error status code
    
# Create a Service object
service = Service(driver_path)

# Start the WebDriver using the Service object and chrome options
driver = webdriver.Chrome(service=service, options=chrome_options)


###Securely Provide Your login info###

#In your working folder, create a new text file (plain text)
#Add your username and password to the text file in this format:
#USERID=youruserid
#PASSWORD=yourpassword

##They should be each on their own line with no quotes and no spaces.

#rename the text file to .env (no file extension).

#Load the .env file
load_dotenv()

# Access variables using os.getenv
userid_value = os.getenv("USERID")
password_value = os.getenv("PASSWORD")

if not userid_value or not password_value:
    raise ValueError("Missing USERID or PASSWORD in .env file.")

####Log into Western's student portal####

# Navigate to the student portal
portal_url = "https://student.uwo.ca/"
driver.get(portal_url)

# Wait up to 10 seconds for the username field to appear
username = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.ID, "userid"))
)

password = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.ID, "pwd"))
)

# Log in to the portal
username.send_keys(userid_value)
password.send_keys(password_value)

# Submit the login form
password.send_keys(Keys.RETURN)

# Wait for the login to complete
time.sleep(5)

print("Logged in successfully!")



#####Download every applicant's materials in a loop#####

def process_applicant(last_name):
    dept_link = ("https://student.uwo.ca/psp/heprdweb/EMPLOYEE/SA/c/"
        "WSA_AAWS_MENU.WSA_SGPS_APPL_INFO.GBL?&cmd=uninav&Rnode=LOCAL_NODE&"
        "uninavpath=Root%7bPORTAL_ROOT_OBJECT%7d.Student%20Admissions%7bHCAD_"
        "STUDENT_ADMISSIONS%7d.Application%20Transaction%20Mgmt%7bHCAD_APPL_TRANS_MGMT%7d&"
        "PORTALPARAM_PTCNAV=WSA_SGPS_APPL_INFO_GBL&EOPP.SCNode=SA&EOPP.SCPortal=EMPLOYEE&"
        "EOPP.SCName=HCAD_STUDENT_ADMISSIONS&EOPP.SCLabel=Student%20Admissions&"
        "EOPP.SCPTfname=HCAD_STUDENT_ADMISSIONS&FolderPath=PORTAL_ROOT_OBJECT.HCAD_"
        "STUDENT_ADMISSIONS.HCAD_APPL_TRANS_MGMT.WSA_SGPS_APPL_INFO_GBL&IsFolder=false")

    #Go to the "Department Assessment" search page.
    
    try:    
        driver.get(dept_link)
        driver.switch_to.default_content()

    # Switch to the relevant iframe
        WebDriverWait(driver, 10).until(
            EC.frame_to_be_available_and_switch_to_it((By.ID, "ptifrmtgtframe"))
    )
    
    ###Search for applications###

    #Ensure the search page is loaded
        last_name_field = WebDriverWait(driver, 10).until(
         EC.presence_of_element_located((By.ID, "WSA_APP_SCR_VW_LAST_NAME_SRCH"))
    )
        admit_term_field = driver.find_element(By.ID, "WSA_APP_SCR_VW_ADMIT_TERM")
        acad_plan_field = driver.find_element(By.ID, "WSA_APP_SCR_VW_ACAD_PLAN")
        acad_plan_condition_dropdown = driver.find_element(By.ID, "WSA_APP_SCR_VW_ACAD_PLAN$op")  # Dropdown for condition

    #Clear the fields before entering new data
        last_name_field.clear()
        admit_term_field.clear()    
        acad_plan_field.clear()

    # Change the dropdown for 'acad plan' to "contains"
        select = Select(acad_plan_condition_dropdown)
        select.select_by_visible_text("contains")  # Select "contains" from the dropdown
    
    # Enter the current application cycle, program filter,
    #and the applicant's last name
        admit_term_field.send_keys(admit_term)
        last_name_field.send_keys(last_name)
        acad_plan_field.send_keys(program_filter)
    
    # click the search button:
        search_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "#ICSearch"))
    )
        search_button.click()

    # Wait for the results table to load
        results_table = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "PTSRCHRESULTS"))
    )

    # Check if results are empty
        rows = results_table.find_elements(By.TAG_NAME, "tr")
        if len(rows) <= 1:  # No results found
            print(f"Applicant '{last_name}' not found in the system.")
            return  # Exit the function gracefully

    # If one or more applicants found, select the first result
        if len(rows) > 1: 
            # Locate the first clickable element
            first_result = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//*[@id='SEARCH_RESULT1']"))
            )
            first_result.click()
            
            print(f"Successfully started processing for: {last_name}")

    except TimeoutException:
        print(f"Timeout: Applicant '{last_name}' could not be found in the system.")
    except Exception as e:
        print(f"An error occurred while processing '{last_name}': {e}")

            
            
            
    # click the 'view all' button:
        view_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "WSA_DERIVED_AP2_VIEW_DETAIL_BTN"))
        )
        view_button.click()
    
   #Wait for download to complete
        WebDriverWait(driver, 30).until(
                lambda d: os.path.exists(os.path.join(download_directory, "merged.pdf"))
            )

    #Rename the downloaded PDF to the student's last name
        downloaded_file = os.path.join(download_directory, "merged.pdf")
        renamed_file = os.path.join(download_directory, f"{last_name}.pdf")
        if os.path.exists(downloaded_file):
            os.rename(downloaded_file, renamed_file)
            print(f"Renamed file to: {renamed_file}")
        else:
            print(f"File for {last_name} not found.") 
 
    except TimeoutException as e:
        print(f"Timeout error while processing {last_name}: {e}")
    except Exception as e:
        print(f"Error processing {last_name}: {e}")

#Read the Excel file to get the list of last names###

# Ensure the Excel file is read from the original directory
excel_file = os.path.join(base_dir, "SelectionWorksheet.xlsx")

try:
    df = pd.read_excel(excel_file, engine='openpyxl')
except Exception as e:
    print(f"ERROR: Unable to read the Excel file '{excel_file}'.\nDetails: {e}")
    sys.exit(1)

# Validate that the required column exists
if 'Last Name' not in df.columns:
    raise ValueError("The Excel file must contain a 'Last Name' column.")
    
# Remove spaces from 'Last Name' column
df['Last Name'] = df['Last Name'].str.replace(" ", "", regex=False)

print("Spaces removed from 'Last Name' column.")

# Process each applicant
for index, row in df.iterrows():
    last_name = row['Last Name']
    print(f"Processing applicant: {last_name}")
    
    try:
        process_applicant(last_name)
    except Exception as e:
         print(f"Error processing {last_name}: {e}")

print("All applicants processed.")

