import tkinter as tk
from tkinter import simpledialog
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import pickle
import os.path
import requests
from googleapiclient.http import MediaIoBaseDownload
import io
import os

# Google API setup (same as the previous example)

# Your template spreadsheet and folder IDs
TEMPLATE_SPREADSHEET_ID = 'your_template_spreadsheet_id'
FOLDER_ID = 'your_folder_id_where_invoices_are_stored'

def copy_invoice_template(service, title):
    """Copy the invoice template and rename the new file."""
    copied_file = {'name': title}
    return service.files().copy(
        fileId=TEMPLATE_SPREADSHEET_ID, body=copied_file, fields='id').execute()

def update_invoice(service, spreadsheet_id, invoice_number, data):
    """Update the new invoice with the provided data."""
    # Assuming 'Sheet1' and specific cells, adjust as needed
    range_name = 'Sheet1!B2'  # Example: Cell where invoice number goes
    value_input_option = 'USER_ENTERED'
    body = {'values': [[invoice_number]]}
    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id, range=range_name,
        valueInputOption=value_input_option, body=body).execute()

    # Update other fields similarly...



# ... (previous code remains the same)

def download_invoice_as_pdf(drive_service, spreadsheet_id, destination_folder):
    """Download the given spreadsheet as a PDF."""
    request = drive_service.files().export_media(fileId=spreadsheet_id, mimeType='application/pdf')
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while done is False:
        status, done = downloader.next_chunk()
    fh.seek(0)

    with open(os.path.join(destination_folder, 'Invoice.pdf'), 'wb') as f:
        f.write(fh.read())
    print("Invoice downloaded as PDF.")

def main():
    creds = None
    # ... (authentication part is the same as previous example)

    service = build('sheets', 'v4', credentials=creds)
    drive_service = build('drive', 'v3', credentials=creds)

    # Simple GUI to input data
    root = tk.Tk()
    root.withdraw()  # Hides the main window
    client_name = simpledialog.askstring("Input", "Enter client name:", parent=root)
    invoice_date = simpledialog.askstring("Input", "Enter invoice date:", parent=root)

    # Generate invoice title and number (customize as needed)
    invoice_number = 1001  # Replace with logic to get the last used invoice number
    invoice_title = f"Invoice_{invoice_number}_{client_name}"

    # Copy and rename invoice template
    new_invoice = copy_invoice_template(drive_service, invoice_title)
    new_invoice_id = new_invoice['id']

    # Update the new invoice
    update_invoice(service, new_invoice_id, invoice_number, {"ClientName": client_name, "Date": invoice_date})

    # Increment and store the new invoice number...
    download_folder = os.path.expanduser('~/Downloads')  # Path to your downloads folder
    download_invoice_as_pdf(drive_service, new_invoice_id, download_folder)

if __name__ == '__main__':
    main()
