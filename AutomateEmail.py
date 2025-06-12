import win32com.client as win32
import pandas as pd
from openpyxl import load_workbook

# ----------------------------------------Reading the Excel file ---------------------------------------------------

def read_excel(file_path):
    df = pd.read_excel(file_path)
    return df

# -----------------------------------------Send Email to Candidates-------------------------------------------------

def Send_email_update_excel():
    # calling the function to read the Excel file
    file_path = "failed.xlsx"
    df = read_excel(file_path)

    # Ensure the "Email Status" column exists, if not, create it
    if "Email Status" not in df.columns:
        df["Email Status"] = None  # Initialize with None if the column does not exist

    # email status list
    Email_status = []

    # Access Outlook Application
    outlook = win32.Dispatch("Outlook.Application")

    cc_emails = [
        "Aditi.Bhatt@cyient.com",
        "Gouramma.Ramachandra@cyient.com"
    ]
    # Join the email addresses with semicolon for CC
    cc_emails_str = ";".join(cc_emails)

    # Loop through each row to send emails
    for index, row in df.iterrows():
        Name = str(row["Candidate Name"]).strip()
        EmailID = str(row["Email"]).strip()
        Position = str(row["Position"]).strip()

        # Skip if the "Email Status" is "Email Sent"
        if row["Email Status"] == "Email Sent":
            print(f"Skipping {Name} ({EmailID}) - Email already sent.")
            continue

        # Validate email
        if pd.isnull(EmailID) or "@" not in EmailID:
            print(f"Skipping {Name} due to invalid email: {EmailID}")
            df.at[index, "Email Status"] = "Failed - Invalid Email"  # Update status in dataframe
            continue

        try:
            # Create a new email item
            mail = outlook.CreateItem(0)
            mail.To = EmailID
            # mail.SenderName = "POD"
            mail.CC = cc_emails_str
            mail.Subject = f"Thank You for Your Application – Update on Your Application for {Position}"

            # Format email body properly
            mail.Body = (
                f"Dear {Name},\n\n"
                "I hope this message finds you well.\n\n"
                "Thank you for your interest in our Internship Program/Position at Cyient. We were genuinely impressed by your "
                "skills and enthusiasm throughout the selection process. Unfortunately, after careful consideration, we have decided to move forward "
                "with another candidate for this particular role.\n\n"
                "This decision was a tough one, as we received many strong applications. That said, we recognize the potential "
                "you bring, and we’ve kept your resume on file. If a position aligned with "
                "your skills and experience opens up in the future, we will certainly consider you for the opportunity.\n\n"
                "We encourage you to stay connected with us on LinkedIn or through our website, where we regularly post "
                "new openings. Your time and effort during the application process are truly appreciated, and we wish you "
                "all the best in your continued career journey.\n\n"
                "Thank you once again for considering Cyient. We look forward to hopefully crossing paths again "
                "in the future!\n\n\n"
                "Best regards,\n"
                "Hiring Team\n"
                "CYIENT LIMITED\n"
            )

            # Send the email
            mail.Send()
            print(f"Email sent successfully to {Name} ({EmailID})")
            df.at[index, "Email Status"] = "Email Sent"  # Update status in dataframe

        except Exception as e:
            print(f"Failed to send email to {Name} ({EmailID}): {e}")
            df.at[index, "Email Status"] = f"Failed - {str(e)}"  # Update status with the failure message
    
    # Save the updated dataframe back to the Excel file
    try:
        df.to_excel(file_path, index=False)
        print("Excel file updated with email statuses")
    except Exception as e:
        print(f"Failed to update Excel file: {e}")

# Call the function
Send_email_update_excel()
