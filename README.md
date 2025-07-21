# Automated Payment Slip Processor for Google Sheets

This Python script automates the process of tracking payments by linking payment slip files from a Google Drive folder to their corresponding entries in a Google Sheet. It is designed to reduce manual data entry and streamline bookkeeping for rental properties or similar services.

## Overview

The script performs the following core tasks:
1.  **Scans a Google Drive Folder:** It lists all files (assumed to be payment slips) in a designated "source" folder.
2.  **Parses Filenames:** It uses regular expressions to extract a unique identifier (e.g., a "room number" like `101/1`) from each filename.
3.  **Reads a Google Sheet:** It fetches data from a specified Google Sheet, including a column of room numbers and designated columns for payment verification.
4.  **Matches and Updates:** It matches the room number from a file to a row in the sheet. If an empty payment cell is found for that row, it updates the cell with a `HYPERLINK` formula pointing directly to the payment slip file.
5.  **Organizes Files:** After processing a file, it moves it to an "updated" folder. If a file's name cannot be parsed or matched, it's moved to a "failed" folder for manual review.
6.  **Handles Overflow:** The script can be configured to check a second spreadsheet if the payment columns for a specific room in the first sheet are already full.

---

## Prerequisites & Setup

Before running the script, you need to configure your Google Cloud, Drive, and Sheets environment.

### 1. Google Cloud Project & Service Account
1.  Create a new project in the [Google Cloud Console](https://console.cloud.google.com/).
2.  Enable the **Google Drive API** and **Google Sheets API** for your project.
3.  Create a **Service Account**:
    - Go to `IAM & Admin` > `Service Accounts`.
    - Click `+ CREATE SERVICE ACCOUNT`, give it a name (e.g., `sheets-automation-bot`), and click `CREATE AND CONTINUE`.
    - Grant it the `Editor` role for simplicity, or more restricted roles if you prefer.
    - Click `CONTINUE`, then `DONE`.
4.  Create a JSON key for the service account:
    - Click on the email of the newly created service account.
    - Go to the **KEYS** tab.
    - Click `ADD KEY` > `Create new key`.
    - Select **JSON** and click `CREATE`.
    - A JSON file will be downloaded. **Rename this file to `service_account.json`** and place it in the root directory of this project.

### 2. Google Drive & Sheets Configuration
1.  **Create Google Drive Folders:** In your Google Drive, create three folders:
    - A source folder for new payment slips (e.g., "New Monthly Slips").
    - An archive folder for successfully processed slips (e.g., "Updated Slips Archive").
    - A folder for slips that failed processing (e.g., "Failed Slips - Manual Check").
2.  **Get Folder & Sheet IDs:** An ID is the unique string in the URL of a Google Drive folder or Google Sheet.
    - Folder URL: `https://drive.google.com/drive/folders/THIS_IS_THE_FOLDER_ID`
    - Sheet URL: `https://docs.google.com/spreadsheets/d/THIS_IS_THE_SPREADSHEET_ID/edit`
3.  **Share Access with Service Account:**
    - Open your `service_account.json` file and find the `client_email` address (e.g., `your-bot@your-project.iam.gserviceaccount.com`).
    - **You must share all three Google Drive folders AND the Google Sheet with this email address**, granting it **Editor** permissions. This is a critical step.

---

## Installation

1.  Clone the repository:
    ```bash
    git clone https://github.com/your-username/your-repository.git
    cd your-repository
    ```
2.  It is highly recommended to use a Python virtual environment:
    ```bash
    python -m venv venv
    source venv/bin/activate  # On Windows: venv\Scripts\activate
    ```
3.  Install the required packages using the provided `requirements.txt` file:
    ```bash
    pip install -r requirements.txt
    ```

---

## Configuration

The script is configured using an environment file.

1.  Create a file named `.env.nextmonth` in the root directory.
2.  Add the following key-value pairs, replacing the placeholder values with your actual IDs obtained during setup.

    ```env
    # The long ID from the URL of your Google Sheet
    SPREADSHEET_ID="1a2b3c4d5e6f7g8h9i0j_YOUR_SPREADSHEET_ID"

    # The ID from the URL of the folder containing NEW payment slips
    FOLDER_ID="1ABCDE_SOURCE_FOLDER_ID"

    # The ID from the URL of the folder where processed slips will be MOVED
    UPDATED_FOLDER_ID="1FGHIJ_UPDATED_FOLDER_ID"

    # The ID from the URL of the folder for slips that could not be matched
    FAILED_FOLDER_ID="1KLMNO_FAILED_FOLDER_ID"
    ```

3.  Inside the script itself, you can adjust static parameters like `sheet_name1`, `last_row`, and the payment columns (`payment_column_start`, `payment_column_end`).

---

## Usage

1.  Place your new payment slip files (e.g., PDFs, images) into the source Google Drive folder specified by `FOLDER_ID`.
2.  **Important:** Ensure filenames contain the room number in a parsable format, such as `Payment Slip (101-1).pdf` or `John Doe 102/3.jpg`. The script looks for patterns like `(101-1)` or `102/3`.
3.  Run the script from your terminal:
    ```bash
    python your_script_name.py
    ```
4.  The script will print a summary of its actions to the console, including which rooms were updated, which were full, and which failed to match. Check the "failed" and "updated" folders in your Google Drive to verify the results.