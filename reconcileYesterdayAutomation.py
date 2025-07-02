import pyodbc
import csv
import smtplib
from email.message import EmailMessage
import io
from datetime import datetime, timedelta
from dotenv import load_dotenv
import os
import sys


def getLastWorkingDate():
    """
    Gets the date for which the query should be run,
    by calculating the previous working day.

    Returns:
        dateString (str): Last working day in YYYY/MM/DD format
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

    # Return date in the format that SQL expects
    dateString = lastWorkingDay.strftime('%Y/%m/%d')
    print(dateString)
    return dateString


def getData(query, dateString):
    """

    """

    # Initialise DB connection, ensuring all necessary credentials exist
    pyodbc.pooling = False
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

    # Connect to database and execute query
    try:
        with pyodbc.connect(conn_str, timeout=5) as conn:
            with conn.cursor() as cursor:
                cursor.execute(query, dateString, dateString)
                rows = cursor.fetchall()

    # Log any errors
    except Exception as e:
        logErrorAndExit(e)

    return rows

def calculateTotals(rows):
    """
    Takes each of yesterdays payments and sums them by type.
    Flips each of these values negative for reasons unbeknown to me
    but I was told that that's how they are entered into the spreadsheet.

    Args:
        rows (list of (str,str) tuples): each individual payment from the last working
            day, where the first is a string representation of the amount, and the 
            second is the type of payment
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

    print(cardTotal, sageTotal, cashTotal, chequeTotal, bankTotal)
    return (cardTotal, sageTotal, cashTotal, chequeTotal, bankTotal)


def dumpToCSV(totals):
    """
    Converts given data into a CSV format stored in an in-memory bytes buffer.

    Writes the column headings and rows to a UTF-8 encoded CSV file in memory,
    suitable for use as an email attachment or other in-memory processing.

    Args:
        totals (tuple of 5 floats): The totals to be written to the CSV.

    Returns:
        io.BytesIO: A bytes buffer containing the CSV data, ready for reading.
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
    Constructs and sends an email via SMTP with a CSV attachment containing yesterday's numbers.
    Logs the success or failure of the email send operation to "logs.txt".

    Args:
        content (io.BytesIO): In-memory bytes buffer containing the CSV data to attach.
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
    msg.set_content('Anything you want in the email body?\nThese numbers should be accurate (i cross checked with BIDS) but this is still in the final testing stage so please let me know if you would like anything else added/changed.')

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
    Logs error message in file logs.txt, then exits the script.

    Args:
        e (Exception): Error message.
    """

    with open("logs.txt", "a") as logs:
        logs.write(f'{e}\n')
    sys.exit(1)


def main():
    """
.
    """

    # Load securely stored credentials for DB and SMTP access
    load_dotenv()
   
    # Define query that extracts relevant data
    query = '''
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
    dateString = getLastWorkingDate()
    rows = getData(query, dateString)
    totals = calculateTotals(rows)
    content = dumpToCSV(totals)
    sendEmail(content)

# Run program
if __name__ == "__main__":
    main()