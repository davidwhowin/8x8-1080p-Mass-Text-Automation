# 8x8-1080p-Mass-Text-Automation
## 8x8 Click Automation

### Overview

This program is crafted to streamline the delivery of promotional messages to a broad customer base, as cataloged in an Excel spreadsheet. It emerges as a strategic solution to circumvent financial constraints that preclude the use of API-driven approaches. Given the extensive number of recipients targeted by this campaign, employing APIs was deemed financially untenable due to the cumulative cost of the requisite calls. This automation technique is ingeniously devised to facilitate mass messaging endeavors efficiently, circumventing the considerable expenses typically associated with API consumption.

### Requirements

- Python 3.x
- pandas
- pyperclip
- pyautogui
- keyboard

# Program Installation and Usage Guide

## Installation

1. **Python Installation:** Ensure Python 3.x is installed on your system.

2. **Library Installation:** Install the required Python libraries using pip. Open your terminal or command prompt and enter the following command:

    ```bash
    pip install pandas pyperclip pyautogui keyboard
    ```

3. **Prepare Excel File:** Ensure the Excel file containing customer data is prepared and accessible.

## Setup

1. **Phone Number Configuration:** Modify the `phone_number` variable in the script to the phone number you intend to use for sending messages.

2. **Image Path:** Update the `image_path` variable with the full path to the promotional image you wish to send.

3. **File Path and Sheet Name:** Adjust the `file_path` and `sheet_name` variables to match the location and name of your Excel spreadsheet.

4. **Excel File Columns:** Ensure your Excel file includes a "Cell Phone" column for phone numbers and a "Checked" column to monitor processed entries.

## Operation

1. Open your terminal or command prompt.

2. Navigate to the directory containing the script.

3. Execute the script with Python by entering the following command:

    ```bash
    python your_script_name.py
    ```

    Replace `your_script_name.py` with the actual name of your Python script.

4. The program will start processing rows from the Excel file, bypassing entries marked as 'read' in the 'Checked' column, verifying for 'STOP' messages, and sending promotional messages to numbers not marked with 'STOP'.

5. To manually stop the program, press 'q'. This script includes keyboard interrupts for emergency stops.

## Caution

- The program automates mouse and keyboard actions. To prevent unintended actions, do not use the mouse and keyboard while the script is running.

- Ensure the screen layout (including resolution and window positions) matches the expected layout coded in the script, as `pyautogui` relies on screen coordinates for mouse actions.

- Regularly update the Excel file to accurately reflect the current customer data and consent status.

## Troubleshooting

- If the program fails to find or interact with UI elements, check the screen resolution and the coordinates used in `pyautogui` functions.

- Verify that all required libraries are installed and updated.

- Ensure the Excel file path and sheet name are correctly specified in the script.
