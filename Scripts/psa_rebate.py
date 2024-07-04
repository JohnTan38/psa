import streamlit as st
import pandas as pd
import numpy as np
#import matplotlib.pyplot as plt
import win32com.client
import glob, os, openpyxl, re
from datetime import datetime
import pythoncom
import warnings
warnings.filterwarnings("ignore")

import smtplib, email, ssl
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

st.set_page_config('PSA Rebates', page_icon="🏛️", layout='wide')
def title(url):
     st.markdown(f'<p style="color:#2f0d86;font-size:22px;border-radius:2%;">{url}</p>', unsafe_allow_html=True)
def title_main(url):
     st.markdown(f'<h1 style="color:#230c6e;font-size:42px;border-radius:2%;">{url}</h1>', unsafe_allow_html=True)

def success_df(html_str):
    html_str = f"""
        <p style='background-color:#fdfdcc;
        color: #09031e;
        font-size: 15px;
        border-radius:5px;
        padding-left: 12px;
        padding-top: 10px;
        padding-bottom: 12px;
        line-height: 18px;
        border-color: #03396c;
        text-align: left;'>
        {html_str}</style>
        <br></p>"""
    st.markdown(html_str, unsafe_allow_html=True)

title_main('PSA Rebates')
pythoncom.CoInitialize()


usr_name = st.sidebar.multiselect('Select your username', ['john.tan', 'linda.lim'], placeholder='Choose 1', 
                          max_selections=2)
if usr_name is not None:
    if st.sidebar.button('Confirm Username'):
            usr_email = usr_name[0]+ '@sh-cogent.com.sg' #your outlook email address
            st.sidebar.write(f'User email: {usr_email}')
            outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI") 
def user_email(usr_name):
    usr_email = usr_name[0] + '@sh-cogent.com.sg'
    return usr_email

psa_rebate = pd.read_csv(r'https://raw.githubusercontent.com/JohnTan38/Project-Income/main/psa_rebate.csv', encoding = "ISO-8859-1")
# psa rebate rates
offpeak_20_24 = psa_rebate.iloc[0, psa_rebate.columns.get_loc('offpeak_24')] #35
offpeak_20_48 = psa_rebate.iloc[0, psa_rebate.columns.get_loc('offpeak_48')] #15
offpeak_40_24 = psa_rebate.iloc[1, psa_rebate.columns.get_loc('offpeak_24')] #52.5
offpeak_40_48 = psa_rebate.iloc[1, psa_rebate.columns.get_loc('offpeak_48')] #22.5

peak_20_24 = psa_rebate.iloc[0, psa_rebate.columns.get_loc('peak_24')] #25
peak_20_48 = psa_rebate.iloc[0, psa_rebate.columns.get_loc('peak_48')] #10
peak_40_24 = psa_rebate.iloc[1, psa_rebate.columns.get_loc('peak_24')] #37.5
peak_40_48 = psa_rebate.iloc[1, psa_rebate.columns.get_loc('peak_48')] #15

def calculate_rebate(df):
    # Define a function to calculate rebate based on the conditions
    def rebate(row):
        if row['Nonpeak'] == 'Yes':
            if row['Size'] == 20:
                return offpeak_20_24 if row['24hr'] == 1 else offpeak_20_48 if row['48hr'] == 1 else 0
            elif row['Size'] == 40:
                return offpeak_40_24 if row['24hr'] == 1 else offpeak_40_48 if row['48hr'] == 1 else 0
        elif row['Nonpeak'] == 'No':
            if row['Size'] == 20:
                return peak_20_24 if row['24hr'] == 1 else peak_20_48 if row['48hr'] == 1 else 0
            elif row['Size'] == 40:
                return peak_40_24 if row['24hr'] == 1 else peak_40_48 if row['48hr'] == 1 else 0
        return 0

    # Apply the function to each row in the DataFrame to calculate the rebate
    df['Rebate'] = df.apply(rebate, axis=1)
    
    return df

def merge_dataframes(df1, df2, col_name):  
    merged_df = pd.merge(df1, df2, on=col_name) # Merge the two dataframes based on col_name
    return merged_df # Return the merged dataframe

def offpeak_rebate_sums(df_rebate):
    # Filter rows based on conditions
    offpeak_20_24 = df_rebate[(df_rebate['Size'] == 20) & (df_rebate['24hr'] == 1) & (df_rebate['Nonpeak'] == 'Yes')]['Rebate'].sum()
    offpeak_40_24 = df_rebate[(df_rebate['Size'] == 40) & (df_rebate['24hr'] == 1) & (df_rebate['Nonpeak'] == 'Yes')]['Rebate'].sum()
    offpeak_20_48 = df_rebate[(df_rebate['Size'] == 20) & (df_rebate['48hr'] == 1) & (df_rebate['Nonpeak'] == 'Yes')]['Rebate'].sum()
    offpeak_40_48 = df_rebate[(df_rebate['Size'] == 40) & (df_rebate['48hr'] == 1) & (df_rebate['Nonpeak'] == 'Yes')]['Rebate'].sum()

    # Create a new DataFrame with the calculated sums
    offpeak_df = pd.DataFrame({
        'offpeak_24hr': [offpeak_20_24, offpeak_40_24],
        'offpeak_48hr': [offpeak_20_48, offpeak_40_48]
    }, index=['20GP', '40GP'])

    return offpeak_df

def peak_rebate_sums(df_rebate):
    # Filter rows based on conditions
    peak_20_24 = df_rebate[(df_rebate['Size'] == 20) & (df_rebate['24hr'] == 1) & (df_rebate['Nonpeak'] == 'No')]['Rebate'].sum()
    peak_40_24 = df_rebate[(df_rebate['Size'] == 40) & (df_rebate['24hr'] == 1) & (df_rebate['Nonpeak'] == 'No')]['Rebate'].sum()
    peak_20_48 = df_rebate[(df_rebate['Size'] == 20) & (df_rebate['48hr'] == 1) & (df_rebate['Nonpeak'] == 'No')]['Rebate'].sum()
    peak_40_48 = df_rebate[(df_rebate['Size'] == 40) & (df_rebate['48hr'] == 1) & (df_rebate['Nonpeak'] == 'No')]['Rebate'].sum()

    # Create a new DataFrame with the calculated sums
    peak_df = pd.DataFrame({
        'peak_24hr': [peak_20_24, peak_40_24],
        'peak_48hr': [peak_20_48, peak_40_48]
    }, index=['20GP', '40GP'])

    return peak_df

def extract_week_number(filename):
    """
    Extracts the week number from a filename in the format "Week_{number}_Y{year}.xlsx".
    Args:
        filename: The filename to extract the week number from.
    Returns:
        The extracted week number as an integer, or None if the filename is not in the expected format.
    """
    match = re.search(r"Week_(\d+)_Y", filename)
    if match:
        return int(match.group(1))
    else:
        return None

def is_valid_filename(file_name):
    """
    This function checks if the given file_name has the format 'Week_'+week_number+'_Y'+year_number+'.xlsx'.
    Parameters:
    file_name (str): The name of the file to check.
    Returns:
    bool: True if the file_name matches the format, False otherwise.
    """
    pattern = r'^Week_\d+_Y\d+\.xlsx$' # Define the regular expression pattern for the filename format
    
    match = re.match(pattern, file_name) # Use the match function to check if the filename matches the pattern
    return match is not None # If there is a match, function returns True, otherwise it returns False

def check_week_in_sheet_names(list_sheetNames, week_number):
    week_number = str(week_number) # Convert week_number to string for comparison
    # Iterate over each sheet name in the list
    for sheet_name in list_sheetNames:
        # Check if the sheet name matches the format 'Week' + ' ' + n
        if sheet_name == 'Week ' + week_number:
            return True  # Return True if match is found

uploaded_file = st.file_uploader("Upload Transport KPI Monitoring excel file", type=['xlsx'])
if uploaded_file is None:
    st.write('Please upload an excel file')
elif uploaded_file:
    fileName = uploaded_file.name
    week_number = extract_week_number(fileName)
    uploaded_wb = openpyxl.load_workbook(uploaded_file)
    sheetNames = uploaded_wb.sheetnames
    
    if not is_valid_filename(fileName) | check_week_in_sheet_names(sheetNames,week_number):
        st.write('Please upload an excel file with valid filename format / worksheet does not exist')
    else:
        sheet_name = 'Week {}'.format(week_number)
        haulier_original = pd.read_excel(uploaded_file, sheet_name=sheet_name, engine='openpyxl')
        haulier_0 = haulier_original[['EventType', 'ContainerNumber', 'CarrierName', 'CarrierVoyage','EventTime', 'Size', '24hr', '48hr', 
                                      'Nonpeak']]

        haulier_0.sort_values(['EventTime', 'CarrierName'], ascending=[True, False], inplace=True)
        haulier_0.dropna(subset=['24hr', '48hr'], inplace=True)
        haulier=haulier_0.copy()

        calculated_rebate = calculate_rebate(haulier) # calculate rebate, get final dataframe

        col_name = 'ContainerNumber'
        rebate = merge_dataframes(haulier_original,calculated_rebate,col_name)
        rebate.drop(columns=['EventType_y', 'CarrierName_y', 'CarrierVoyage_y', 'EventTime_y', 'Size_y', '24hr_y', '48hr_y', 
                             'Nonpeak_y'], inplace=True) #Rebate_y
        rebate.rename(columns = {'EventType_x': 'EventType', 'CarrierName_x': 'CarrierName', 'CarrierVoyage_x': 'CarrierVoyage', 
                                 'EventTime_x': 'EventTime', 'Size_x': 'Size', '24hr_x': '24hr', '48hr_x': '48hr', 'Nonpeak_x': 'Nonpeak', 
                                 'Rebate_x': 'Rebate'}, inplace = True)

        rebate_final = rebate.copy() # total peak / offpeak rebates
        peak_offpeak_rebate = pd.merge(offpeak_rebate_sums(rebate_final), peak_rebate_sums(rebate_final), left_index=True, right_index=True)


        st.divider()
        title('PSA Rebate Summary ($): week ' + str(week_number))
        st.dataframe(peak_offpeak_rebate)

        #with pd.ExcelWriter("C:/Users/"+usr_name[0]+ "Downloads/"+ 'psa_rebate.csv') as writer_rebate:
        sheetName = 'psa_rebate_week'+ str(week_number)+ '_'+ datetime.now().strftime("%Y%m%d %H%M")
        try:
                    rebate_final.to_csv("C:/Users/"+usr_name[0]+ "/Downloads/"+ 'psa_rebate.csv', mode='x')
        except FileExistsError:
                    rebate_final.to_csv("C:/Users/"+usr_name[0]+ "/Downloads/"+ 'psa_rebate_1.csv')

def send_email_psa_reabte(df,usr_email,subj_email):
    usr_email = user_email(usr_name)
    email_receiver = usr_email
    #email_receiver = st.multiselect('Select one email', ['john.tan@sh-cogent.com.sg', 'vieming@yahoo.com'])
    email_sender = "john.tan@sh-cogent.com.sg"
    email_password = "PASSWORD" #st.secrets["password"]

    body = """
            <html>
            <head>
            <title>Dear User</title>
            </head>
            <body>
            <p style="color: blue;font-size:25px;">PSA Rebate ($) offpeak/peak.</strong><br></p>

            </body>
            </html>

            """+ df.to_html() +"""
        
            <br>This message is computer generated. """+ datetime.now().strftime("%Y%m%d %H:%M:%S")

    mailserver = smtplib.SMTP('smtp.office365.com',587)
    mailserver.ehlo()
    mailserver.starttls()
    mailserver.login(email_sender, email_password)
       
    try:
            if email_receiver is not None:
                try:
                    rgx = r'^([^@]+)@[^@]+$'
                    matchObj = re.search(rgx, email_receiver)
                    if not matchObj is None:
                        usr = matchObj.group(1)
                    
                except:
                    pass

            msg = MIMEMultipart()
            msg['From'] = email_sender
            msg['To'] = email_receiver
            msg['Subject'] = 'PSA Rebate Summary Week_'+ str(subj_email)+' '+ datetime.today().strftime("%Y%m%d %H:%M")
            msg['Cc'] = 'john.tan@sh-cogent.com.sg'
        
            msg.attach(MIMEText(body, 'html'))
            text = msg.as_string()

            with smtplib.SMTP("smtp.office365.com", 587) as server:
                server.ehlo()
                server.starttls()
                server.login(email_sender, email_password)
                server.sendmail(email_sender, email_receiver, text)
                server.quit()
            #st.success(f"Email sent to {email_receiver} 💌 🚀")
            success_df(f"Email sent to {email_receiver} 💌 🚀")
    except Exception as e:
            st.error(f"Email not sent: {e}")

if st.button('Email'):
    usr_email = user_email(usr_name)
    send_email_psa_reabte(peak_offpeak_rebate,usr_email,week_number)

footer_html = """
    <div class="footer">
    <style>
        .footer {
            position: fixed;
            bottom: 0;
            left: 0;
            right: 0;
            background-color: #f0f2f6;
            padding: 10px 20px;
            text-align: center;
        }
        .footer a {
            color: #4a4a4a;
            text-decoration: none;
        }
        .footer a:hover {
            color: #3d3d3d;
            text-decoration: underline;
        }
    </style>
        All rights reserved @2024. Cogent Holdings IT Solutions.      
    </div>
"""
st.markdown(footer_html,unsafe_allow_html=True)
