import pyodbc
import csv
import smtplib
from email.message import EmailMessage
import io
from datetime import datetime, timedelta
from dotenv import load_dotenv
import os
import sys


def getDateStrings():
    """
    Gets the dates for which the query should be run,
    by calculating the previous working day.

    Returns:
        yesterdayString (str): Last working day in YYYY/MM/DD format
        todayString (str): Current working day in YYYY/MM/DD format
    """

    # Gets the day that it is currently (Monday = 0, Sunday = 6)
    today = datetime.today()
    weekday = today.weekday()

    if weekday == 0:
        # Monday → go back to Friday
        lastWorkingDay = today - timedelta(days=3)
    else:
        # Tuesday–Friday → go back one day
        lastWorkingDay = today - timedelta(days=1)

    # Return dates in the format that SQL expects
    yesterdayString = lastWorkingDay.strftime('%Y/%m/%d')
    todayString = today.strftime('%Y/%m/%d')
    print(yesterdayString)
    print(todayString)

    return yesterdayString, todayString


def getData(generalQuery, bankTransferQuery, yesterdayString, todayString):
    """
    Executes SQL queries to retrieve payment data for the previous working day.

    Args:
        generalQuery (str): SQL query for non-bank-transfer payments.
        bankTransferQuery (str): SQL query for bank transfer payments.
        yesterdayString (str): Last working day in 'YYYY/MM/DD' format.
        todayString (str): Current day in 'YYYY/MM/DD' format.

    Returns:
        list: Combined list of rows from both queries.
              Each row is a tuple containing (str amount, str payment type).
    """

    # Initialise DB connection, ensuring all necessary credentials exist
    server = os.getenv("SQL_SERVER")
    database = os.getenv("SQL_DATABASE")
    uid = os.getenv("SQL_UID")
    pwd = os.getenv("SQL_PWD")

    if not all([server, database, uid, pwd]):
        logErrorAndExit(ValueError("Missing one or more SQL connection environment variables"))

    # Define connection string
    conn_str = (
        f"DRIVER=ODBC Driver 17 for SQL Server;"
        f"SERVER=tcp:{server};"
        f"DATABASE={database};"
        f"UID={uid};"
        f"PWD={pwd};"
    )

    # Connect to database and execute queries, combining their results
    try:
        with pyodbc.connect(conn_str) as conn:
            with conn.cursor() as cursor:
                cursor.execute(generalQuery, yesterdayString, yesterdayString)
                generalRows = cursor.fetchall()

                cursor.execute(bankTransferQuery, todayString, todayString)
                bankTransferRows = cursor.fetchall()

                rows = generalRows + bankTransferRows

    # Log any errors
    except Exception as e:
        logErrorAndExit(e)

    return rows


def calculateTotals(rows):
    """
    Aggregates and negates payment totals by type.

    Args:
        rows (list): List of tuples containing (amount, payment type) as strings.

    Returns:
        tuple: A tuple of 5 floats representing the negated totals for:
               (cardTotal, sageTotal, cashTotal, chequeTotal, bankTotal)
    """

    cardTotal = 0
    sageTotal = 0
    cashTotal = 0
    chequeTotal = 0
    bankTotal = 0

    # Gets the type of each transaction and subtracts it from the correct total
    # (this leads to all totals being the negative of what BIDS says, which is
    # how the data is inputted into the spreadsheet)
    for row in rows:
        amount = float(row[0].replace(',',''))
        type = row[1].lower().strip()

        match type:
            case 'credit card' | 'debit card': cardTotal -= amount
            case 'sage pay': sageTotal -= amount
            case 'cash': cashTotal -= amount
            case 'cheque': chequeTotal -= amount
            case 'bank transfer': bankTotal -= amount
            case _ : pass

    # Return each total rounded to 2 decimal places
    return (round(cardTotal, 2), round(sageTotal, 2), round(cashTotal, 2), round(chequeTotal, 2), round(bankTotal, 2))


def dumpToCSV(totals):
    """
    Converts payment totals into a UTF-8 encoded CSV in memory.

    Args:
        totals (tuple): Tuple of 5 floats for Card, SagePay, Cash, Cheque, Bank Transfer.

    Returns:
        io.BytesIO: In-memory bytes buffer containing the CSV file content.
    """
    headings = ('Card', 'Sagepay', 'Cash', 'Cheque', 'Bank Transfer')

    # Initialise bytes object to be sent as email attachment
    content = io.BytesIO()
    textWrapper = io.TextIOWrapper(content, encoding='utf-8', newline='')
    writer = csv.writer(textWrapper)

    # Write to file
    writer.writerow(headings)
    writer.writerow(totals)

    # Flush and rewind to the beginning
    textWrapper.flush()
    textWrapper.detach()
    content.seek(0)

    # Return CSV content
    return content


def sendEmail(content):
    """
    Sends an email with the CSV report attached via SMTP.

    Args:
        content (io.BytesIO): In-memory buffer containing CSV data.

    Raises:
        Logs and exits the script if email sending fails.
    """

    # SMTP configuration (gets credentials from .env file so as not to hard code them in the script)
    smtpServer = os.getenv('SMTP_SERVER')
    smtpPort = int(os.getenv('SMTP_PORT'))
    username = os.getenv('SMTP_USERNAME')
    password = os.getenv('SMTP_PASSWORD')

    # Create the email message
    msg = EmailMessage()
    msg['Subject'] = f'Yesterday\'s numbers (you choose the subject)'
    msg['From'] = username
    msg['To'] = os.getenv('SMTP_RECIPIENT')

    # Set message content
    msg.set_content('The attachment contains the results of the RepGen reports used to complete the bank balances summary spreadsheet.')

    # Attach CSV
    msg.add_attachment(
        content.read(),
        maintype='text',
        subtype='csv',
        filename='yesterday.csv'
    )

    # Send the email
    try:
        with smtplib.SMTP(smtpServer, smtpPort) as server:
            server.starttls()
            server.login(username, password)
            server.send_message(msg)

            with open("logs.txt", "a") as logs:
                logs.write(f'Successful run: {datetime.now()}\n')

    except Exception as e: logErrorAndExit(e)


def logErrorAndExit(e):
    """
    Logs an error message and exits the script.

    Args:
        e (Exception): The exception or error message to log.
    """

    with open("logs.txt", "a") as logs:
        logs.write(f'{e}\n')
    sys.exit(1)


def main():
    """
    Main script execution flow:
        1. Loads environment variables.
        2. Defines SQL queries.
        3. Calculates relevant dates.
        4. Retrieves and processes data.
        5. Converts data to CSV.
        6. Sends the email with CSV attached.
    """

    # Load securely stored credentials for DB and SMTP access
    load_dotenv()
   
    # Define queries that extract relevant data
    generalQuery = '''
    SELECT
        FORMAT(tblClientsLedger.gross, 'N2') AS [Gross],
        tblClientsLedger.paymentType         AS [Type]

    FROM
        tblClientsLedger

    WHERE
        ISNULL(tblClientsLedger.tranDate, '') BETWEEN ? AND ?
        AND ISNULL(tblClientsLedger.paymentType, '') != 'Bank Transfer'
        AND ISNULL(tblClientsLedger.cashOffice, '') != 'vendor statements'
    '''

    bankTransferQuery = '''
    SELECT
        FORMAT(tblClientsLedger.gross, 'N2') AS [Gross],
        tblClientsLedger.paymentType         AS [Type]

    FROM
        tblClientsLedger

    WHERE
        ISNULL(tblClientsLedger.tranDate, '') < ?
        AND ISNULL(tblClientsLedger.postingDate, '') BETWEEN ? AND '2099/12/31'
        AND ISNULL(tblClientsLedger.cashOffice, '') != 'vendor statements'
    '''

    # Main process
    yesterdayString, todayString = getDateStrings()
    rows = getData(generalQuery, bankTransferQuery, yesterdayString, todayString)
    totals = calculateTotals(rows)
    content = dumpToCSV(totals)
    sendEmail(content)

# Run program
if __name__ == "__main__":
    main()