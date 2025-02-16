import requests
import time

# Microsoft Graph API authentication details
CLIENT_ID = "29c8707e-876a-4833-8c31-4cad33a8ac0b"
CLIENT_SECRET = "wn28Q~0N-DlfWfyiSQfY.GqFLAfN4g1EuuNkhcHy"
TENANT_ID = "751d98b4-f8be-4510-9ccc-97bdb4e50d02"

# List of email accounts to process
EMAIL_ACCOUNTS = [
    "brijesh@xleadoutreach.com",
    "mahendra@xleadsconsulting.com",
    "lakhendra@xleadsconsulting.com",
    "xgrowthtech@xleadsconsulting.com",
    "audit@xleadoutreach.com",
    "warmup@xleadoutreach.com"
]

# Get Access Token
def get_access_token():
    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    data = {
        "grant_type": "client_credentials",
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default"
    }
    response = requests.post(url, data=data)
    return response.json().get("access_token")

# Get Folder ID by Name
def get_folder_id(email, folder_name):
    access_token = get_access_token()
    headers = {"Authorization": f"Bearer {access_token}"}
    url = f"https://graph.microsoft.com/v1.0/users/{email}/mailFolders"

    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        for folder in response.json().get("value", []):
            if folder["displayName"].lower() == folder_name.lower():
                return folder["id"]
    return None

# Fetch unread emails from a folder (with pagination)
def get_all_unread_emails(email, folder_name):
    folder_id = get_folder_id(email, folder_name)
    if not folder_id:
        return []

    access_token = get_access_token()
    headers = {"Authorization": f"Bearer {access_token}"}
    email_list = []
    next_link = f"https://graph.microsoft.com/v1.0/users/{email}/mailFolders/{folder_id}/messages?$filter=isRead eq false&$top=100"

    while next_link:
        response = requests.get(next_link, headers=headers)
        if response.status_code != 200:
            return email_list

        data = response.json()
        email_list.extend(data.get("value", []))
        next_link = data.get("@odata.nextLink")

    return email_list

# Move emails in bulk
def move_emails_bulk(email):
    email_list = get_all_unread_emails(email, "junk email")
    email_ids = [email["id"] for email in email_list]

    if not email_ids:
        print(f"✅ No emails in Junk to move for {email}.")
        return

    inbox_id = get_folder_id(email, "Inbox")
    if not inbox_id:
        return

    access_token = get_access_token()
    headers = {"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}

    def chunk_list(lst, chunk_size):
        for i in range(0, len(lst), chunk_size):
            yield lst[i:i + chunk_size]

    for batch in chunk_list(email_ids, 50):
        for email_id in batch:
            url = f"https://graph.microsoft.com/v1.0/users/{email}/messages/{email_id}/move"
            body = {"destinationId": inbox_id}
            requests.post(url, headers=headers, json=body)

    print(f"✅ Moved {len(email_ids)} emails to Inbox for {email}.")

    # Recursively move emails if any are left
    remaining_junk_emails = get_all_unread_emails(email, "junk email")
    if remaining_junk_emails:
        move_emails_bulk(email)

# Display unread emails in Inbox
def display_unread_inbox_emails(email):
    unread_emails = get_all_unread_emails(email, "Inbox")
    if not unread_emails:
        print(f"✅ No unread emails in Inbox for {email}.")
        return []

    print(f"\n📌 Unread Emails in Inbox for {email} (Before Reading):")
    for email_data in unread_emails[:50]:  # Display only first 50
        subject = email_data.get("subject", "No Subject")
        sender = email_data.get("from", {}).get("emailAddress", {}).get("address", "Unknown Sender")
        print(f"📩 {subject} - From: {sender}")

    return unread_emails

# Read unread emails in bulk & Mark them as Read
def read_unread_emails_inbox(email):
    unread_emails = display_unread_inbox_emails(email)  # Display before reading
    if not unread_emails:
        return

    access_token = get_access_token()
    headers = {"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}

    email_ids = [email["id"] for email in unread_emails]

    def chunk_list(lst, chunk_size):
        for i in range(0, len(lst), chunk_size):
            yield lst[i:i + chunk_size]

    for batch in chunk_list(email_ids, 50):
        for email_id in batch:
            url = f"https://graph.microsoft.com/v1.0/users/{email}/messages/{email_id}"
            response = requests.get(url, headers=headers)
            if response.status_code == 200:
                email_data = response.json()
                subject = email_data.get("subject", "No Subject")
                sender = email_data.get("from", {}).get("emailAddress", {}).get("address", "Unknown Sender")
                print(f"📖 Read: {subject} - From: {sender}")

        # **Mark these emails as Read**
        for email_id in batch:
            mark_as_read_url = f"https://graph.microsoft.com/v1.0/users/{email}/messages/{email_id}"
            update_body = {"isRead": True}
            requests.patch(mark_as_read_url, headers=headers, json=update_body)

    print(f"✅ Read and marked {len(email_ids)} emails as read for {email}.")

# Run script for multiple accounts
for email in EMAIL_ACCOUNTS:
    print(f"\n🚀 Processing emails for {email}...\n")
    move_emails_bulk(email)
    read_unread_emails_inbox(email)
