# WesternGradScraper

This script automates the process of logging into Western University's student portal, 
searching for graduate school applications, and downloading applicant materials 
based on a list of last names provided in an Excel file.

---

## **Features**
- Creates an "Applicants" directory in your project folder.
- Logs into PeopleSoft using your credentials provided in a .env file.
- Automatically searches for applications based on last names and program filters.
- Downloads and renames applicant materials (PDF files) by last name.

---

## **Setup Instructions**

### 1. Prerequisites
Ensure the following are installed on your system:
- Python 3.8+  
- Google Chrome (check your version at `chrome://settings/help`)

Install the required Python libraries:
```bash
pip install pandas openpyxl selenium python-dotenv
```

---
### 2. Folder Setup
Organize your project folder like this:

```plaintext
WesternGradScraper/
│
├── main.py           		# Your Python script
├── SelectionWorksheet.xlsx  	# Excel file with applicant last names
├── .env                	# Environment file with login credentials
├── chromedriver.exe    	# ChromeDriver for Selenium
└── requirements.txt    	# Dependencies list
```

---

### 3. Prepare the `.env` File
In the project folder, create a `.env` file to securely store your credentials:
```plaintext
USERID=your_username
PASSWORD=your_password
```

---

### 4. Prepare the excel file.
In the project folder, include an excel file with a list of the applicants who you would like to evaluate. 

## **Example Excel File Format**
The script reads applicant last names from an Excel file (`Selection Worksheet.xlsx`). Ensure the file has a "Last Name" column:

| Last Name  |
|------------|
| Smith      |
| Johnson    |
| Williams   |


### 5. Download ChromeDriver
1. Check your Chrome version at `chrome://settings/help`.
2. Download the matching ChromeDriver from ChromeDriver Downloads(https://chromedriver.chromium.org/downloads).
3. Place the `chromedriver.exe` file in your project folder.


### 6. Update Configuration File (Optional)
The script will prompt you to change these default settings. Alternatively, you can update `config.ini` with your preferred settings:

```ini
[DEFAULT]
admit_term = 1258
program_filter = cl
excel_file = SelectionWorksheet.xlsx
```


## **How to Use WesternGradScraper**

1. **Run the Script**  
   Open a terminal, navigate to the project folder, and run:
   ```bash
   python main.py
   ```

2. **User Inputs**
   The script will prompt you for:
   - Admit term
   - Program filter (e.g., 'cl', 'sp', 'cd', 'io')
   - Excel file name
   Leave these blank and press the enter key to use the default options from the configuration file.


3. **What Happens**:
   - An Applicants directory is generated inside your project folder.
   - The script logs into the Western University portal using your credentials.
   - The script searches for each applicant based on last names and program filters.
   - Applicant materials are downloaded, renamed, and saved within the Applicants directory.


## **Dependencies**
The script requires the following libraries:
- `pandas` – For Excel file handling.
- `openpyxl` – Excel file engine.
- `selenium` – For web scraping and automation.
- `python-dotenv` – For secure environment variable management.

Install dependencies using:
```bash
pip install -r requirements.txt
```

---

## **Expected Output**
The script outputs messages to the terminal, such as:
```
Logged in successfully!
Successfully started processing for: Smith
Renamed file to: Smith.pdf
Successfully started processing for: Johnson
...
All applicants processed.
```

The downloaded PDFs are renamed and saved to the Applicants folder.

---

## **Troubleshooting**
- Ensure ChromeDriver matches your Chrome version.
- Verify that your Excel file contains the correct column (`Last Name`).
---

## **Author**
Created by **Samantha Joel**.  