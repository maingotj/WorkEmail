import win32com.client as win32

def create_outlook_email():
    # Connect to Outlook
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)  # 0 = olMailItem

    # Predefined parts
    subject_prefix = "Project Update: "
    body_text = (
        "Hi there,\n\n"
        "This is the standard body text.\n\n"
        "Best,\n"
        "Jonathan"
    )

    # Ask user for inputs
    recipient = input("Enter recipient email: ")
    additional_subject = input("Enter additional subject info: ")

    # Fill in email fields
    mail.To = recipient
    mail.Subject = subject_prefix + additional_subject
    mail.Body = body_text

    # Open the email so user can still edit/send it
    mail.Display()

if __name__ == "__main__":
    create_outlook_email()
