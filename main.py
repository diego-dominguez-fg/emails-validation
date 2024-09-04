import pandas as pd
from validate_email_address import validate_email
from datetime import datetime

# Load the Excel file
file_path = 'NewUsersAugustW5.xlsx'
xls = pd.ExcelFile(file_path)

# Read the data
df = pd.read_excel(xls, 'Hoja1')

# Define known valid email servers and company domain
valid_email_servers = ["gmail.com", "yahoo.com", "outlook.com", "hotmail.com", "aol.com", "live.com.mx","hotmail.es", "outlook.es", "yahoo.com.mx"]
company_domain = "fragua.com.mx"
invalid_patterns = ["@fg", "@mail.com", "notinee", "notiene"]

# Initialize lists for categorization
valid_emails = []
suspicious_emails = []
invalid_emails = []
company_emails = []

# Process each email
for index, row in df.iterrows():
    email = row['LOGONID']
    domain = email.split('@')[-1] if '@' in email else ''

    # Check if it's a company email
    if domain == company_domain:
        company_emails.append(row)
    # Check if it's a known invalid pattern
    elif any(pattern in email for pattern in invalid_patterns):
        invalid_emails.append(row)
    # Validate the email
    elif validate_email(email):
        if domain in valid_email_servers:
            valid_emails.append(row)
        else:
            suspicious_emails.append(row)
    else:
        invalid_emails.append(row)

# Convert lists to DataFrames
valid_emails_df = pd.DataFrame(valid_emails)
suspicious_emails_df = pd.DataFrame(suspicious_emails)
invalid_emails_df = pd.DataFrame(invalid_emails)
company_emails_df = pd.DataFrame(company_emails)

# Get the current date in YYYY-MM-DD format
current_date = datetime.now().strftime('%Y-%m-%d')

# Create a new Excel file with categorized emails
output_path = f'{current_date}.xlsx'

with pd.ExcelWriter(output_path) as writer:
    valid_emails_df.to_excel(writer, sheet_name='ValidEmails', index=False)
    suspicious_emails_df.to_excel(writer, sheet_name='SuspiciousEmails', index=False)
    invalid_emails_df.to_excel(writer, sheet_name='InvalidEmails', index=False)
    company_emails_df.to_excel(writer, sheet_name='CompanyEmails', index=False)

print(f"Email categorization completed and saved to '{output_path}'")