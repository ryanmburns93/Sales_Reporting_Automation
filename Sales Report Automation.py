# -*- coding: utf-8 -*-
"""
Created on Sat Oct 16 22:59:18 2021

@author: Ryan
"""

from __future__ import print_function
import excel2img
import pandas as pd
import os
from openpyxl import load_workbook
from openpyxl.styles import Alignment
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import base64
import os.path
from google.oauth2.credentials import Credentials


#  pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib


def get_col_widths(dataframe):
    # First we find the maximum length of the index column
    idx_max = max([len(str(s)) for s in dataframe.index.values] + [len(str(dataframe.index.name))])
    # Then, we concatenate this to the max of the lengths of column name and its values for each column, left to right
    return [idx_max] + [max([len(str(s)) for s in dataframe[col].values] + [len(col)]) for col in dataframe.columns]


def gather_input_info():
    '''
    Gather information from the user for customizing the program outputs.

    Returns
    -------
    image_date : String.
        String of the date for the end of the period. Used to label each image
        produced. Ex. 'NOV 15' for the November 1 - November 15 reporting period.
    sheetname : String.
        Name of the worksheet in the Excel workbook for the reporting period.
        Ex. 'NOV 1 to 15' for the November 1 - November 15 reporting period.
    period_start_date : String.
        The starting date of the reporting period for customizing the outgoing
        email. Ex. 'November 1st' for the November 1 - November 15 reporting period.
    period_end_date : String.
        The ending date of the reporting period for customizing the outgoing
        email. Ex. 'November 15th' for the November 1 - November 15 reporting period.
    current_full_report : DataFrame.
        The current reporting period transactions report in DataFrame format.
    vendor_info_df : DataFrame.
        The vendor information table with vendor ID, contact information, and
        payment details in DataFrame format.

    '''
    image_date = input('Please type the date to mark the screenshots with: ')
    sheetname = input('Please type the name of the sheet for this period: ')
    period_start_date = input('Please type the date to mark the start of'
                              ' the period for the email blast:')
    period_end_date = input('Please type the date to mark the end of'
                            'the period for the email blast:')
    current_full_report = pd.read_excel('C:/Users/Ryan/BevCol/The Beverly Collective Sales.xlsx',
                                        sheetname)
    vendor_info_df = pd.read_excel('C:/Users/Ryan/BevCol/The Beverly Collective Sales.xlsx',
                                   'VENDOR LIST')
    return (image_date, sheetname, period_start_date,
            period_end_date, current_full_report, vendor_info_df)


def create_image_attachments(current_full_report, image_date):
    '''
    Creates an image screenshoting the transactions of each vendor who made sale
    transactions during the current reporting period and saves the image to the local
    directory.

    Parameters
    ----------
    current_full_report : DataFrame.
        DataFrame containing the current reporting period transactions report.
    image_date : String.
        The final report date to include in the title of each saved image file.

    Returns
    -------
    None.

    '''
    for vendor in current_full_report['VENDOR ID'].unique():
        temp_df = current_full_report[current_full_report['VENDOR ID']==vendor]
        writer = pd.ExcelWriter(f'{vendor}_temp.xlsx')
        temp_df.to_excel(writer, sheet_name='Sheet1', index=False)
        # Auto-adjust columns' width
        for column in temp_df:
            column_width = max(temp_df[column].astype(str).map(len).max(), len(column))
            if column_width < 10:
                column_width = 10
            col_idx = temp_df.columns.get_loc(column)
            writer.sheets['Sheet1'].set_column(col_idx, col_idx, column_width)
        writer.save()
        wb = load_workbook(filename = f'{vendor}_temp.xlsx')
        ws = wb.active
        for i in range(len(temp_df)):
            _cell = ws[f'I{i+2}']
            _cell.number_format = 'm/d/yyyy'
            _cell = ws[f'A{i+2}']
            _cell.alignment = Alignment(horizontal='center')
            _cell = ws[f'J{i+2}']
            _cell.number_format = 'h:mm AM/PM'
            for currency_col in ['E', 'F', 'G', 'H']:
                _cell = ws[f'{currency_col}{i+2}']
                _cell.number_format = '$#,##0.00_);[Red]($#,##0.00)'
        wb.save(f'{vendor}_temp.xlsx')
        excel2img.export_img(f'{vendor}_temp.xlsx',
                             f'{vendor} Sales {image_date}.png',
                             'Sheet1',
                             None)
        os.remove(f'{vendor}_temp.xlsx')
    return


def create_message(sender, to, subject, message_text):
    """Create a message for an email.
    Code from: https://developers.google.com/gmail/api/guides/drafts#python

    Args:
        sender: Email address of the sender.
        to: Email address of the receiver.
        subject: The subject of the email message.
        message_text: The text of the email message.

      Returns:
          An object containing a base64url encoded email object.
    """
    message = MIMEMultipart(message_text)
    message = MIMEText(message_text)
    message['to'] = to
    message['from'] = sender
    message['subject'] = subject
    return {'raw': base64.urlsafe_b64encode(message.as_string().
                                            encode('UTF-8')).
            decode('ascii')}


def create_draft(service, user_id, subject, message_body, to):
    """Create and insert a draft email. Print the returned draft's message and id.
    Some code of this function sourced from:
    https://developers.google.com/gmail/api/guides/drafts

    Args:
        service: Authorized Gmail API service instance.
        user_id: User's email address. The special value "me"
        can be used to indicate the authenticated user.
        message_body: The body of the email message, including headers.

    Returns:
        Draft object, including draft id and message meta data.
    """
    message = MIMEText(message_body)
    message['to'] = to
    message['from'] = user_id
    message['subject'] = subject
    draft_message = {'message': {'raw': base64.urlsafe_b64encode(message.as_string().encode('UTF-8')).decode('ascii')}}
    draft = service.users().drafts().create(userId=user_id, body=draft_message).execute()

    print(f'Draft id: {draft["id"]}')
    print(f'Draft message: {draft["message"]}')

    return draft



def create_email_text(vendor_firstname, date_start, date_end, amount, payment_method, email):
    text = f'''{email}\n
Hi {vendor_firstname},\n
I have attached a screenshot of your sales at The Beverly Collective for the period of {date_start} through {date_end}.\n
I will make the payment of {"${:,.2f}".format(amount)} to you via {payment_method} soon.\n
Thank you, and I hope you are having a great week!\n
Best,
Monica
'''
    return text


def create_zero_sales_email_text(vendor_firstname, date_start, date_end, email):
    text = f'''{email}\n
Hi {vendor_firstname}!\n
I am sorry to say that you didn't have any sales at The Beverly Collective for the period of (date_start) through {date_end}.\n
Thank you, and I hope you are having a great week!\n
Best,
Monica'''
    return text


def calculate_summary_amount(current_full_report):
    '''
    Creates an aggregate sales transactions report providing top-level sales
    information for each vendor.

    Parameters
    ----------
    current_full_report : DataFrame.
        DataFrame containing the current reporting period transactions report.

    Returns
    -------
    grouped_df : DataFrame.
        DataFrame containing a summary output of the aggregate net sales, sales tax,
        vendor earnings, and The Beverly Collective commission for each vendor.

    '''
    grouped_df = current_full_report.groupby(by='VENDOR ID').agg('sum')
    grouped_df = grouped_df.drop(columns='TRANSACTION')
    grouped_df.to_csv('C:/Users/Ryan/BevCol/Bev_Sales_Summary.csv')
    return grouped_df


def create_vendor_email_drafts(service, current_full_report, grouped_df, vendor_info_df, period_start_date, period_end_date, image_date):
    '''
    Creates an outgoing email draft in the Gmail account of the owner of The Beverly
    Collective containing customized subject and body. Some code of this function
    sourced from: https://developers.google.com/gmail/api/guides/drafts

    Parameters
    ----------
    service : Gmail API Connection.
        An active connection with live certification and token to the Gmail API.
    current_full_report : DataFrame.
        DataFrame containing the current reporting period transactions report.
    grouped_df : DataFrame.
        DataFrame containing a summary output of the aggregate net sales, sales tax,
        vendor earnings, and The Beverly Collective commission for each vendor.
    vendor_info_df : DataFrame.
        The vendor information table with vendor ID, contact information, and
        payment details in DataFrame format.
    period_start_date : String.
        The starting date of the reporting period for customizing the outgoing
        email. Ex. 'November 1st' for the November 1 - November 15 reporting period.
    period_end_date : String.
        The ending date of the reporting period for customizing the outgoing
        email. Ex. 'November 15th' for the November 1 - November 15 reporting period.
    image_date : String.
        String of the date for the end of the period. Used to label each image
        produced. Ex. 'NOV 15' for the November 1 - November 15 reporting period.

    Returns
    -------
    None.

    '''
    for vendor in current_full_report['VENDOR ID'].unique():
        vendor_amount = grouped_df['VENDOR PAYOUT'][vendor]
        vendor_email = str(vendor_info_df[vendor_info_df['VENDOR CODE']==vendor]
                           ['EMAIL'].values)[2:-2]
        vendor_firstname = str(vendor_info_df[vendor_info_df['VENDOR CODE']==vendor]
                               ['NAME']).split()[1]
        if vendor_amount == 0:
            create_draft(service=service,
                         user_id='monica.vercillo@gmail.com',
                         subject=f'{vendor} Update {image_date} The Beverly Collective',
                         message_body=create_zero_sales_email_text(vendor_firstname,
                                                                   period_start_date,
                                                                   period_end_date,
                                                                   vendor_email),
                         to=vendor_email)
            continue
        vendor_amount = grouped_df['VENDOR PAYOUT'][vendor]
        payment_method = str(vendor_info_df[vendor_info_df['VENDOR CODE']==vendor]
                             ['PREFERRED PAYMENT METHOD'].values[0])
        create_draft(service=service,
                     user_id='monica.vercillo@gmail.com',
                     subject=f'{vendor} Update {image_date} The Beverly Collective',
                     message_body=create_email_text(vendor_firstname,
                                                    period_start_date,
                                                    period_end_date,
                                                    vendor_amount,
                                                    payment_method,
                                                    vendor_email),
                     to=vendor_email)
    return


def establish_gmail_api_connection():
    ''' Establish Gmail API Connection, prompting user for app permission and
    returning active API connection. Most code of this function sourced from:
    https://developers.google.com/gmail/api/quickstart/python

    Returns
    -------
    service : Gmail API Connection.
        An active connection with live certification and token to the Gmail API.

    '''
    os.remove('C:/Users/Ryan/BevCol/token.json')
    # If modifying these scopes, delete the file token.json.
    SCOPES = ['https://www.googleapis.com/auth/gmail.readonly',
              'https://www.googleapis.com/auth/gmail.compose',
              'https://www.googleapis.com/auth/gmail.insert',
              'https://www.googleapis.com/auth/gmail.settings.sharing',
              'https://www.googleapis.com/auth/gmail.addons.current.message.metadata',
              'https://www.googleapis.com/auth/gmail.addons.current.action.compose',
              'https://www.googleapis.com/auth/gmail.addons.current.message.action',
              'https://www.googleapis.com/auth/gmail.labels']
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    service = build('gmail', 'v1', credentials=creds)

    # Call the Gmail API
    results = service.users().labels().list(userId='me').execute()
    labels = results.get('labels', [])

    if not labels:
        print('No labels found.')
    else:
        print('Labels:')
        for label in labels:
            print(label['name'])
    return service


def main():
    (image_date,
     sheetname,
     period_start_date,
     period_end_date,
     current_full_report,
     vendor_info_df) = gather_input_info()
    grouped_df = calculate_summary_amount(current_full_report)
    create_image_attachments(current_full_report, image_date)
    service = establish_gmail_api_connection()
    create_vendor_email_drafts(service, current_full_report, grouped_df,
                               vendor_info_df, period_start_date,
                               period_end_date, image_date)
