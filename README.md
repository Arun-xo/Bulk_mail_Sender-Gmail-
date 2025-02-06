# Bulk Email Sender - User Guide

## Overview
The Bulk Email Sender application is a Python-based tool that allows users to send bulk emails using multiple sender accounts. The application supports HTML email templates, automatic email tracking, and a pause/resume feature for better control over email campaigns.

---

## How to Use the Application

### 1. **Prepare the Excel File**
The application reads recipient details from an Excel file. The required format is as follows:

| Column Name       | Description                                             |
|------------------|---------------------------------------------------------|
| `from_email`     | Sender's Gmail address                                  |
| `password`       | App password for the sender's Gmail account            |
| `sal`            | Salutation (e.g., Mr., Ms., Dr.)                        |
| `signature`      | Signature name to be used in the email                 |
| `to_email`       | Recipient's email address                              |
| `subject`        | Subject of the email                                   |
| `html_file`      | Full path to the HTML email template file              |

**Example:**

| from_email         | password  | sal  | signature  | to_email       | subject       | html_file                            |
|--------------------|----------|------|------------|---------------|--------------|-------------------------------------|
| sender1@gmail.com | abcd1234 | Mr.  | John Doe   | user1@abc.com | Business Deal | C:\Users\user\Desktop\email1.html  |
| sender2@gmail.com | efgh5678 | Ms.  | Jane Smith | user2@xyz.com | New Offer     | C:\Users\user\Desktop\email2.html  |

> **Note:** Ensure the email sender accounts have app passwords enabled for SMTP authentication.

---

### 2. **Launching the Application**
1. **Run the application:** If using the Python script, execute:
   ```bash
   python Bulk_Mail_Sender.py
   ```
   If you have an executable version, simply double-click the `.exe` file.

2. **Load the Excel File:** Click on the "Load Excel File" button and select the prepared Excel file.

3. **Set Email Sending Delay:** Enter the time delay (in seconds) between each email to avoid being flagged as spam. Recommended delay:
   - 5-10 seconds for normal sending
   - 15-20 seconds for high-volume sending

4. **Start Sending Emails:** Click on the "Send Emails" button to begin the process.

5. **Pause/Resume Sending:** The app allows you to pause and resume the process at any time.

6. **Monitor Progress:** The progress bar updates in real-time, displaying sent and pending emails.

7. **Check Logs and Reports:**
   - The log section provides real-time updates on sent emails.
   - If any emails fail to send, a failure report (`failed_emails_report.xlsx`) is generated.

---

## Additional Features
- **Automatic Network Monitoring:** If the network disconnects, the process is paused and resumes automatically once connectivity is restored.
- **Multiple Sender Accounts:** The app rotates between multiple sender email IDs to distribute sending volume.
- **HTML Email Support:** The application allows sending emails using pre-designed HTML templates.
- **Custom Signature Per Sender:** The signature varies based on the sender’s information in the Excel file.

---

## Troubleshooting
### 1. **Email Not Sending?**
- Ensure that the sender's Gmail account allows app passwords.
- Verify that the password entered in the Excel file is correct.
- Check your internet connection.

### 2. **HTML File Not Found Error?**
- Ensure the `html_file` column contains the correct full file path.
- The file should exist in the specified location.

### 3. **Gmail Sending Limits?**
- Each Gmail account has a daily sending limit (typically 500 for free accounts, 2,000 for Google Workspace accounts).
- If you are sending a large volume of emails, use multiple sender accounts and adjust the sending delay.

---

## Conclusion
This Bulk Email Sender is a powerful tool for automating email campaigns. By properly setting up your Excel file and adjusting sending intervals, you can maximize email deliverability and efficiency. If you have any issues, refer to the troubleshooting section for solutions.

