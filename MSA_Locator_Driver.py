import requests
import tkinter as tk
from tkinter import ttk, END
from html2text import html2text

import Constants

# Replace with your own client ID, client secret, and tenant ID
client_id = Constants.MSA_REPORT_BTN_CLIENT_ID
client_secret = Constants.MSA_REPORT_BTN_API
tenant_id = Constants.MSA_REPORT_BTN_TENANT_ID

# Function to retrieve the user profile and email details
def retrieve_email_details():
    # Get the user and email Graph IDs from the input fields
    user_graph_id = user_entry.get()
    email_graph_id = email_entry.get()

    # remove any whitespaces from beginning and end of user_graph_id and email_graph_id
    user_graph_id = user_graph_id.strip()
    email_graph_id = email_graph_id.strip()

    # Microsoft Graph API endpoints
    graph_endpoint = 'https://graph.microsoft.com/v1.0'
    user_endpoint = f'{graph_endpoint}/users/{user_graph_id}'

    # Get an access token using client credentials flow
    token_url = f'https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token'
    data = {
        'grant_type': 'client_credentials',
        'client_id': client_id,
        'client_secret': client_secret,
        'scope': 'https://graph.microsoft.com/.default'
    }
    response = requests.post(token_url, data=data)
    access_token = response.json()['access_token']

    # Set the authorization header with the access token
    headers = {
        'Authorization': 'Bearer ' + access_token,
        'Content-Type': 'application/json'
    }

    # Make a request to the Microsoft Graph API to retrieve the user profile
    response = requests.get(user_endpoint, headers=headers)

    # Check if the request was successful
    if response.status_code == 200:
        user_profile = response.json()
        display_name_label.config(text=f"Display Name: {user_profile['displayName']}")
        email_address_label.config(text=f"Email Address: {user_profile['mail']}")

        # Get the specific email
        email_endpoint = f"{user_endpoint}/messages/{email_graph_id}"
        response = requests.get(email_endpoint, headers=headers)

        if email_graph_id == "":
            error_label.config(text='Error: No Email Graph ID Was Entered')

            # clear email details from previous entry
            subject_label.config(text=f"Subject: ")
            from_label.config(state="normal")
            from_label.delete(0, END)
            from_label.config(state="readonly")
            received_label.config(text=f"Received: ")
            body_text.delete('1.0', tk.END)

            # clear email graph id entry field
            email_entry.delete(0, END)

        elif response.status_code == 200:
            # clear email details from previous entry
            subject_label.config(text=f"Subject: ")
            from_label.config(state="normal")
            from_label.delete(0, END)
            from_label.config(state="readonly")
            received_label.config(text=f"Received: ")
            body_text.delete('1.0', tk.END)

            # clear email graph id entry field
            email_entry.delete(0, END)

            error_label.config(text='')
            email = response.json()
            subject_label.config(text=f"Subject: {email['subject']}")
            from_label.config(state="normal")
            from_label.insert(0, f"From: {email['from']['emailAddress']['address']}")
            from_label.config(state="readonly")
            received_label.config(text=f"Received: {email['receivedDateTime']}")

            # Parse and display the email body
            email_body = html2text(email['body']['content'])
            body_text.delete('1.0', tk.END)
            body_text.insert(tk.END, email_body)

            # clear email graph id entry field
            email_entry.delete(0, END)

        else:
            error_label.config(text='Error: Unable to retrieve email details.')

            # clear email details from previous entry
            subject_label.config(text=f"Subject: ")
            from_label.config(state="normal")
            from_label.delete(0, END)
            from_label.config(state="readonly")
            received_label.config(text=f"Received: ")
            body_text.delete('1.0', tk.END)

            # clear email graph id entry field
            email_entry.delete(0, END)
    else:
        error_label.config(text='Error: Unable to retrieve user profile.')

        # clear email details from previous entry
        subject_label.config(text=f"Subject: ")
        from_label.config(state="normal")
        from_label.delete(0, END)
        from_label.config(state="readonly")
        received_label.config(text=f"Received: ")
        body_text.delete('1.0', tk.END)

# Create the Tkinter window
window = tk.Tk()
window.title('Microsoft Graph API Email Retrieval')
window.geometry('700x800')

# Create a canvas with a scrollbar
canvas = tk.Canvas(window)
canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
scrollbar = ttk.Scrollbar(window, orient=tk.VERTICAL, command=canvas.yview)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
canvas.configure(yscrollcommand=scrollbar.set)
canvas.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))

# Create a frame inside the canvas
frame = tk.Frame(canvas)
canvas.create_window((0, 0), window=frame, anchor=tk.NW)

# Create labels and entry fields for user and email Graph IDs
user_label = tk.Label(frame, text='User Email Address:', font=('Arial', 14))
user_label.pack()
user_entry = tk.Entry(frame, font=('Arial', 10), width= 35, justify= 'center')
print(type(user_entry))
user_entry.pack()

email_label = tk.Label(frame, text='Email Graph ID:', font=('Arial', 14))
email_label.pack()
email_entry = tk.Entry(frame, font=('Arial', 10), width= 50)
email_entry.pack()

# Create a button to retrieve email details
retrieve_button = tk.Button(frame, text='Retrieve Email Details', command=retrieve_email_details, font=('Arial', 14))
retrieve_button.pack()

# Create labels to display email details
display_name_label = tk.Label(frame, text='', font=('Arial', 14))
display_name_label.pack()
email_address_label = tk.Label(frame, text='', font=('Arial', 14))
email_address_label.pack()
subject_label = tk.Label(frame, text='', font=('Arial', 14))
subject_label.pack()
from_label = tk.Entry(frame, text='', font=('Arial', 14), width= 40, state="readonly", justify="center")
from_label.pack()
received_label = tk.Label(frame, text='', font=('Arial', 14))
received_label.pack()

# Create a text widget for the email body
body_label = tk.Label(frame, text='Body:', font=('Arial', 14))
body_label.pack()
body_canvas = tk.Canvas(frame)
body_canvas.pack(fill=tk.BOTH, expand=True)
body_text = tk.Text(body_canvas, wrap=tk.WORD, font=('Arial', 10))
body_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
body_scrollbar = ttk.Scrollbar(body_canvas, orient=tk.VERTICAL, command=body_text.yview)
body_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
body_text.configure(yscrollcommand=body_scrollbar.set)
body_canvas.configure(yscrollcommand=body_scrollbar.set)

# Create an error label
error_label = tk.Label(frame, text='', font=('Arial', 14))
error_label.pack()

window.mainloop()
