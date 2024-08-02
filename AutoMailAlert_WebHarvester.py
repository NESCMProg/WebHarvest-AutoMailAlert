import os
import pandas as pd
from datetime import datetime, timedelta
import smtplib
from email.mime.text import MIMEText
import configparser

# Function to send email with CC recipients
def send_email(recipient, cc_recipients, subject, body, smtp_server, smtp_port, smtp_username, smtp_password):
    msg = MIMEText(body, 'html')
    msg['Subject'] = subject
    msg['From'] = smtp_username
    msg['To'] = recipient
    if cc_recipients:
        msg['Cc'] = ', '.join(cc_recipients)

    print(f"Sending email to {recipient} with CC to {cc_recipients}...")
    try:
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(smtp_username, smtp_password)
            recipients = [recipient] + cc_recipients
            server.sendmail(smtp_username, recipients, msg.as_string())
        print(f"Email sent to {recipient} with CC to {cc_recipients}")
    except Exception as e:
        print(f"Failed to send email to {recipient} with CC to {cc_recipients}: {e}")

# Function to get the latest date from the network path
def get_latest_date_from_path(path):
    print(f"Checking files in the directory: {path}")
    try:
        if not os.path.exists(path):
            print(f"Directory does not exist: {path}")
            return None

        files = os.listdir(path)
        dates = []
        for file in files:
            try:
                date = datetime.strptime(file, '%m-%d-%y').date()
                dates.append(date)
            except ValueError:
                print(f"Skipping file with invalid date format: {file}")
        
        if not dates:
            print(f"No valid date files found in directory: {path}")
            return None

        latest_date = max(dates)
        print(f"Latest date found in directory: {latest_date}")
        return latest_date
    except Exception as e:
        print(f"Error reading directory {path}: {e}")
        return None

# Read config file
config = configparser.ConfigParser()
config.read('harvester.config')

# Get paths and SMTP server details from config file
file_path = config.get('Paths', 'excel_file_path')
smtp_server = config.get('SMTP', 'server')
smtp_port = config.getint('SMTP', 'port')
smtp_username = config.get('SMTP', 'username')
smtp_password = config.get('SMTP', 'password')

print(f"Reading Excel file: {file_path}")
df = pd.read_excel(file_path)

# Explicitly cast 'Last_Delivery_Date' column to datetime
df['Last_Delivery_Date'] = pd.to_datetime(df['Last_Delivery_Date'], errors='coerce')

# Check for 'Data Delivery Location' header and update Last_Delivery_Date
if 'Data Delivery Location' in df.columns:
    for index, row in df.iterrows():
        data_delivery_location = row['Data Delivery Location']
        latest_date = get_latest_date_from_path(data_delivery_location)
        if latest_date:
            # Convert latest_date to string format before assignment
            df.at[index, 'Last_Delivery_Date'] = latest_date.strftime('%Y-%m-%d')

# Explicitly cast 'Last_Delivery_Date' column to datetime again after updating it
df['Last_Delivery_Date'] = pd.to_datetime(df['Last_Delivery_Date'], errors='coerce')

# Save the updated DataFrame back to the Excel file
print(f"Saving updated Last_Delivery_Date back to the Excel file: {file_path}")
df.to_excel(file_path, index=False)

# Re-read the Excel file after updating Last_Delivery_Date
df = pd.read_excel(file_path)

# Explicitly cast 'ExpectedNextDate' column to datetime
df['ExpectedNextDate'] = pd.to_datetime(df['ExpectedNextDate'], errors='coerce')

# Today's date
today = datetime.today().date()
print(f"Today's date: {today}")

# Dictionary to map frequencies to the number of days to add
frequency_mapping = {
    'daily': 1,
    'alternate days': 2,
    'weekly': 7,
    'bi-weekly': 14,
    'monthly': 30,
    'quarterly': 90
}

# Process each row in the DataFrame
for index, row in df.iterrows():
    print(f"\nProcessing row {index + 1}...")
    # Parse Last_Delivery_Date
    last_delivery_date = pd.to_datetime(row['Last_Delivery_Date']).date()
    print(f"Last Delivery Date: {last_delivery_date}")
    
    # Get the NES_Actual_Frequency and convert it to lower case
    frequency = row['NES_Actual_Frequency'].strip().lower()
    if frequency in frequency_mapping:
        next_delivery_date = last_delivery_date + timedelta(days=frequency_mapping[frequency])
        print(f"Next Delivery Date: {next_delivery_date}")
        
        # Save the Next Delivery Date to ExpectedNextDate column
        df.at[index, 'ExpectedNextDate'] = next_delivery_date.strftime('%Y-%m-%d')
        
        # Compare with today's date
        if next_delivery_date < today:
            print(f"Next delivery date {next_delivery_date} is behind today's date {today}. Setting IsNeedEmailAlert to 1.")
            # Set IsNeedEmailAlert to 1
            df.at[index, 'IsNeedEmailAlert'] = 1
        else:
            print(f"Next delivery date {next_delivery_date} is not behind today's date {today}. Setting IsNeedEmailAlert to 0.")
            df.at[index, 'IsNeedEmailAlert'] = 0

# Explicitly cast 'ExpectedNextDate' column to datetime again after updating it
df['ExpectedNextDate'] = pd.to_datetime(df['ExpectedNextDate'], errors='coerce')

# Save the updated DataFrame back to the Excel file
print(f"Saving updated DataFrame back to the Excel file: {file_path}")
df.to_excel(file_path, index=False)

# Re-read the Excel file after updating all columns
df = pd.read_excel(file_path)

# Process each row for sending emails
for index, row in df.iterrows():
    if df.at[index, 'IsNeedEmailAlert'] == 1:
        prodcode = row.get('Prodcode(s)', 'N/A')
        publisher_name = row.get('Publisher Name', 'N/A')
        cc_recipients = row['CC_mails'].split(';') if isinstance(row['CC_mails'], str) else []

        # Create the email body with the updated row details in a table format
        body = f"""
        <html>
        <body>
            <p>This is a reminder that Web harvesting {publisher_name} Publisher - {prodcode} Product code is need to start the downloads for next delivery on time.</p>
            <table border="1">
                <tr>
                    <th>Column Name</th>
                    <th>Value</th>
                </tr>
        """
        for col_name in df.columns:
            body += f"""
                <tr>
                    <td>{col_name}</td>
                    <td>{row[col_name]}</td>
                </tr>
            """
        body += """
            </table>
        </body>
        </html>
        """
        
        for _ in range(int(row['RepeatedCount'])):
            send_email(
                recipient=row['Reciver_mails'],
                cc_recipients=cc_recipients,
                subject="Delivery Alert",
                body=body,
                smtp_server=smtp_server,
                smtp_port=smtp_port,
                smtp_username=smtp_username,
                smtp_password=smtp_password
            )

print("Process completed successfully.")
