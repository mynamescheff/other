import os
import re
import win32com.client

# Outlook credentials
your_email = 'your_email@example.com'
shared_mailbox_email = 'shared_mailbox@example.com'

# Outlook category to filter emails
category_to_download = 'YourCategory'

# Get the absolute path of the script
script_path = os.path.dirname(os.path.abspath(__file__))

# Path to save attachments
attachment_save_path = os.path.join(script_path, 'outlook/excel')

# Path to save Outlook messages (.msg)
msg_save_path = os.path.join(script_path, 'outlook/msg')

# Create directories if they don't exist
os.makedirs(attachment_save_path, exist_ok=True)
os.makedirs(msg_save_path, exist_ok=True)

def download_attachments_and_save_as_msg(outlook, category):
    namespace = outlook.GetNamespace("MAPI")

    # Resolve the shared mailbox recipient
    recipient = namespace.CreateRecipient(shared_mailbox_email)
    recipient.Resolve()

    if recipient.Resolved:
        shared_mailbox = namespace.GetSharedDefaultFolder(recipient, 6)  # 6 corresponds to the Inbox folder

        # Get the total number of emails in the specified category
        total_emails = shared_mailbox.Items.Restrict(f"[Categories] = '{category}'").Count
        print(f"Total emails in category '{category}': {total_emails}")

        # Initialize a counter for saved emails
        saved_emails = 0

        for item in shared_mailbox.Items.Restrict(f"[Categories] = '{category}'"):
            try:
                # Mark the email as read
                item.UnRead = False

                # Check if the email has attachments
                if item.Attachments.Count > 0:
                    # Process each attachment
                    for attachment in item.Attachments:
                        # Check if the attachment has .xlsx extension
                        if attachment.FileName.lower().endswith('.xlsx'):
                            # Extract new filename from subject
                            new_filename = extract_filename_from_subject(item.Subject)

                            # Replace characters that might interfere with file paths
                            new_filename = re.sub(r'[\/:*?"<>|]', ' ', new_filename)

                            # Check for allowed characters in the filename
                            new_filename = re.sub(r'[^\w\s\-\–]', '', new_filename)

                            # If ";" is not present in the subject or filename, print as invalid
                            if ';' not in new_filename and ';' not in item.Subject:
                                print(f"Invalid filename: {new_filename}")
                                continue

                            # Save attachment with sanitized filename
                            attachment_path = os.path.join(attachment_save_path, f"{new_filename}.xlsx")
                            try:
                                attachment.SaveAsFile(attachment_path)
                                saved_emails += 1

                                # Save the entire email with " approval" suffix
                                approval_msg_path = os.path.join(msg_save_path, f"{new_filename} approval.msg")
                                try:
                                    item.SaveAs(approval_msg_path)
                                except Exception as msg_error:
                                    print(f"Error saving approval message: {msg_error}")
                                    print(f"Approval message path: {approval_msg_path}")
                            except Exception as attachment_error:
                                print(f"Error saving attachment: {attachment_error}")
                                print(f"Attachment path: {attachment_path}")
                        else:
                            # Print a message for emails with attachments other than .xlsx
                            print(f"Email with subject '{item.Subject}' has attachments other than .xlsx")

            except Exception as e:
                print(f"Error processing email with subject: {item.Subject}")
                print(f"Error details: {e}")

        # Print the number of saved emails
        print(f"Saved emails: {saved_emails}")
    else:
        print(f"Could not resolve the recipient: {shared_mailbox_email}")

def extract_filename_from_subject(subject):
    # Extract filename after the first ";" character in the subject
    match = re.search(r';\s*(.*)', subject)
    if match:
        return match.group(1)
    else:
        # If no ";" found, use the entire subject
        return subject

if __name__ == "__main__":
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    namespace.Logon(your_email)

    download_attachments_and_save_as_msg(outlook, category_to_download)

    print("Attachments downloaded and emails saved as .msg files.")
