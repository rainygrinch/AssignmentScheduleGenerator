### Assignment Schedule Generator - README ###

#### VERSION RELEASE NOTES ####

**VERSION 1.1**
- Asks the user to specify the input file to be used, instead of specifying a file path.
- Prompts the user to select a destination folder for saving the output Word files.
- Seeks a `.csv` file instead of a `.xlsx` file.
- Ignores the first line of the CSV export from Bullhorn.
- Placeholders updated to use Bullhorn defaults by modifying placeholders in the `template.docx` file.

**VERSION 1.2**
- Added a count for the total number of attempted document generations.
- Added a count for successful document generations.
- Displays a list of unsuccessful document generations by Placement ID.
- Added a welcome screen, disclaimer, and instructions.

**VERSION 1.3**
- Implemented a process timer to measure execution time.
- Removed Static Template File Directory (allowing the user to select it themselves).

**VERSION 1.4**
- Completely refactored for GUI interface.
- Added a progress bar and real-time document count.
- Fixed issue where successful document generation count was always 1 less than total.
- Corrected issue where failed Placement IDs did not display.

**VERSION 1.5**
- Added program information to the main window, equivalent to the title screen in version 1.3 and earlier.
- Bug fix: "Generate Docs" button now properly displays and does not get pushed below window height.
- Simplified title information.
- Increased window size to accommodate all content.
- Added NOTICE PERIOD to TEMPLATE and CSV input file.
- Added PROJECT NAME to TEMPLATE and CSV input file.

### VERSION 1.6 ###

# Set PAYRATE to 2 decimal places instead of 1

---

#### DESCRIPTION ####
This program generates assignment schedules by replacing placeholders in a Word template with data from a CSV file. The user selects the necessary files and output location through a graphical interface.

---

#### USAGE INSTRUCTIONS ####
1. **Run the Program**: Execute the script.
2. **Select Input File**: Choose the CSV file containing placement data.
3. **Select Template File**: Choose the Word document template with placeholders.
4. **Select Output Folder**: Choose where the generated assignment schedules should be saved.
5. **Processing**: The program will replace placeholders in the template with data from the CSV and generate assignment schedules.
6. **Completion Summary**:
   - Displays the total number of attempted document generations.
   - Displays the number of successfully generated documents.
   - Lists Placement IDs for failed generations.
   - Shows the total execution time.

---

#### REQUIREMENTS ####


- Python 3.x
- Required libraries: `pandas`, `python-docx`, `tkinter`
- CSV input file containing placement data (with Bullhorn-compatible placeholders).
- Word template file (`.docx`) with placeholders for replacement.

---

#### DISCLAIMER ####
This software is a prototype developed for IDPP Consulting Ltd. No liability is accepted for malfunctions or data loss. Use at your own risk.

---

#### AUTHOR ####
**Created by:** Peter Grint (RainyGrinch)  
**For use by:** IDPP Consulting Ltd.
**petergrint@idpp.com**

---

#### IMPORTANT NOTE ####
Make sure that whichever headers you use in the CSV file are used as placeholders in the Word document template file, surrounded by double curly braces `{{ }}`.
Please note you do not need Python or the libraries to run the .exe file, but you will  need the CSV Input File and Word Template File
This programs SKIP THE SECOND ROW OF THE EXCEL INPUT FILE - if your CSV file has the column titles/headers on row 1, leave row 2 empty, and start your
data set from row 3!