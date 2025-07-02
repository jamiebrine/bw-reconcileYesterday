# MoD Sales Data Exporter

This Python script connects to a SQL Server database to retrieve new sales and lot details since the last run, generates a CSV file in-memory, and emails it to a specified recipient. It also tracks the last processed date to avoid duplicate data in future runs.

---

## Features

- Connects securely to SQL Server using credentials stored in a `.env` file
- Retrieves sales data filtered by the last processed date stored in `previous.txt`
- Generates a CSV attachment in-memory for efficient email delivery
- Sends an email with the CSV attached if new records exist; otherwise sends a notification
- Logs successes and failures in `logs.txt`
- Updates the tracking file with the latest processed date

---

## Setup

1. **Clone or download the files** `MoDAutomation.py` and `requirements.txt` into a new directory

2. **Create a `.env` file** in the same directory with the following variables:
   ```env
   SQL_SERVER=your_sql_server_name
   SQL_DATABASE=your_database_name
   SQL_UID=your_db_username
   SQL_PWD=your_db_password

   SMTP_SERVER=your_smtp_server
   SMTP_PORT=587
   SMTP_USERNAME=your_smtp_username
   SMTP_PASSWORD=your_smtp_password
   SMTP_RECIPIENT=recipient_email_address
3. **Ensure `previous.txt` exists** with a valid date (e.g., `01 Jan 2000`) to serve as the starting point for data retrieval.

4. **Install required packages** by running `pip install -r requirements.txt`

---

## Usage

Run the script manually with `python MoDAutomation.py`, or schedule it to run as often as required

---

## Notes
- Ensure the machine running this script has network access to the SQL Server and SMTP server.

- Handle sensitive credentials securely and never commit `.env` or passwords to version control.

---