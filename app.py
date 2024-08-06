from flask import Flask, render_template, request, session, send_file
from pywhatkit import sendwhats_image,sendwhatmsg_instantly
import csv
import io
import secrets
import pandas as pd
import time

import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

##Create WSGI Application app is instance of class Flask
app=Flask(__name__)
app.secret_key = secrets.token_hex(16)


##Decorator for telling which method of webpage i am going to access
@app.route('/')
def trial():
    return render_template('index.html')

@app.route('/upload_csv', methods=['POST'])
def upload_csv():
    # Check if the POST request has the file part
    csv_file = request.files['csv_file']
    print(csv_file)

    # If user does not select file, browser also
    # submit an empty part without filename
    if csv_file.filename == '':
        return 'No selected file'

    if csv_file:
        # Read the CSV file and process its data
        stream = io.StringIO(csv_file.stream.read().decode("UTF8"), newline=None) 
        csv_input = csv.reader(stream)
        session['csv_data'] = list(csv_input)
        return 'CSV file uploaded successfully'  # Return a valid response
    return 'Invalid request'  # Handle other cases, such as unexpected file formats 

@app.route('/send_message', methods=['POST'])
def send_message():
    message = request.form['message']
    attachment = request.files['attachment']
    # Your existing backend code to send the message
    # For simplicity, I'm commenting out the actual sending part
    csv_data=session['csv_data']# Find the index of the column containing WhatsApp numbers
    whatsapp_index = None
    for idx, header in enumerate(csv_data[0]):
        if 'whatsapp_numbers' in header.lower():
            whatsapp_index = idx
            break
    print(whatsapp_index)
    # Fetch all WhatsApp numbers if the index is found
    if whatsapp_index is not None:
        whatsapp_numbers = [row[whatsapp_index] for row in csv_data[1:]]
    else:
        # Handle the case where 'whatsapp_numbers' column is not found
        whatsapp_numbers = []
    print(whatsapp_numbers)
   
    #cleaned_numbers = [number.strip('"') for number in whatsapp_numbers]

    for number in whatsapp_numbers:
        try:
            # Use pywhatkit to send the message
            if attachment and attachment.filename != '':
                sendwhats_image("+91"+str(number), attachment)
                time.sleep(30)
                sendwhatmsg_instantly("+91"+str(number), message)
                time.sleep(10)
            else:
                sendwhatmsg_instantly("+91"+str(number), message)
                time.sleep(10)
            
            print(f"Message sent to {number} successfully!")
        except Exception as e:
            print(f"Error occurred while sending message to {number}: {str(e)}")

    return "Message sent successfully!"

# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = "1fkFI4_uaildO1QeogRPZcMExHQ5rB3bDAJyD6XuNraw"

@app.route('/main', methods=['GET'])
def main():
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file(
          "credentials.json", SCOPES
        )
        creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open("token.json", "w") as token:
      token.write(creds.to_json())

    try:
        service = build("sheets", "v4", credentials=creds)
        
        def get_sheet_range(service):
            spreadsheet = service.spreadsheets().get(spreadsheetId=SAMPLE_SPREADSHEET_ID).execute()
            sheet_properties = spreadsheet['sheets'][0]['properties']  # Assuming you are using the first sheet
            sheet_title = sheet_properties['title']
            sheet_range = f"{sheet_title}!A1:ZZ"  # Adjust the range as needed
            return sheet_range

        service = build('sheets', 'v4', credentials=creds)
        SAMPLE_RANGE_NAME = get_sheet_range(service)

        # Call the Sheets API
        sheet = service.spreadsheets()
        result = (
            sheet.values()
            .get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=SAMPLE_RANGE_NAME)
            .execute()
        )
        values = result.get("values", [])

        if not values:
            print("No data found.")
            return
        print(values)
        
        def filter_interested(values):
            # Find the index of the column with the header 'Are you interested in our college ?'
            headers = values[0]  # Assuming the headers are in the first row
            interested_col_index = headers.index('Are you interested in our college ?')
            
    
            # Filter out entries where 'Are you interested in our college ?' is 'Yes'
            interested_entries = [entry for entry in values[1:] if entry[interested_col_index] == 'Yes']
            filtered_data = [headers] + interested_entries
            print(filtered_data)
            return filtered_data
        
        df = pd.DataFrame(filter_interested(values))
        df.to_csv('output_file.csv', index=False,header=False)
        if not os.path.exists('output_file.csv'):
            return 'CSV data not found'
        
        return send_file('output_file.csv', as_attachment=True)
    except HttpError as err:
        print(err)
  
@app.route("/new")      
def testing():
    return render_template("new.html")
    
        
if __name__=='__main__':
    app.run(debug=True)