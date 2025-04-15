import streamlit as st
from datetime import datetime
import pandas as pd
import win32com.client as win32
import pythoncom

def table(col,df):
    
    html_table = """<table border="1">"""
    
    html_table += "<tr>"
    
    for c in col:
        html_table += f"<th>{c}</th>"
        
    html_table += "</tr>"
    
    for i in range(len(df)):
        row = df[i]
        html_table += "<tr>"
        
        for j in row:
            html_table += f"<td>{j}</td>"
        html_table += "</tr>"
            
    html_table += "</table>"
    
    return html_table.replace("nan","")


def send(body):
    
    pythoncom.CoInitialize()
    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)

    # Paramétrer l'email
    mail.Subject = "Test Tableau dans un Email"
    mail.Body = f"{body}"
    mail.HTMLBody = f"<html><body>{mail.Body}</body></html>"

    # Ajouter le destinataire
    mail.To = "Morgan.Stanley.Swiss@morganstanley.com"

    # Envoyer l'email
    mail.Send() 

def ms_pricer(initial_df):
    
    col_ms = ['Request ID', 'Product', 'Wrapper', 'Issuer', 'Currency', 'Size',
       'Reoffer (%)', 'Strike Date', 'Tenor (m)', 'BBG Code 1', 'BBG Code 2',
       'BBG Code 3', 'BBG Code 4', 'BBG Code 5', 'Gearing', 'Strike (%)',
       'KI Barrier (%)', 'Barrier Type', 'Early Termination Period',
       'Early Termination Level (%)', 'Autocall from Period X',
       'Coupon Frequency', 'Trigger Level (%) ', 'Periodic Coupon (%)',
       'Memory coupon']
    
    
    

    
    df = pd.DataFrame([],index = range(len(initial_df)),columns=col_ms)

    df["Currency"] = initial_df["Currency"]
    df["Size"] = initial_df["Size"].apply(lambda x: f"{x:,}")
    df["Reoffer (%)"] = initial_df["Reoffer"].apply(lambda x: "{:.2f}".format(x)+"%")
    df["Strike Date"] = initial_df["Strike_Date"].apply(lambda x: x.strftime("%d-%b-%y"))
    df["Tenor (m)"] = initial_df["Tenor"]
    for i in range(1,6):
        
        df[f"BBG Code {i}"] = initial_df[f"Underlying {i}"]
        
    df['Strike (%)'] = initial_df["Strike"].apply(lambda x: "{:.2f}".format(x)+"%")
    df['Barrier Type'] = len(initial_df)*[""]
    try:
        df['KI Barrier (%)'] = initial_df["KI_Barrier"]
        df[initial_df['KI_Barrier'] != ""]['Barrier Type'] = len(initial_df[initial_df['KI_Barrier'] != ""])*["European"]
    except:
        pass
    
    df['Early Termination Period'] = initial_df["Frequency"]
    df['Early Termination Level (%)'] = initial_df["Autocall_Barrier"].apply(lambda x: "{:.2f}".format(x)+"%")
    df['Autocall from Period X'] = len(initial_df)*[1]
    
    df["Product"] = len(initial_df)*["Phoenix Autocall"]
    df["Wrapper"] = len(initial_df)*["Note"]
    df["Issuer"] = len(initial_df)*["MSBV"]
    

    return df








today = datetime.today()
# Title of the application
st.title("Pricer Interface")







































if "params" not in st.session_state:
    st.session_state.params = dict()

if "submit" not in st.session_state:
    st.session_state.submit = None

if "valid" not in st.session_state:
    st.session_state.valid = False

# Product Summary Section

if not st.session_state.valid: 
    with st.expander("Product Summary"):
        col1, col2 = st.columns(2)
        with col1:
            st.session_state.params["Currency"] = st.selectbox("Currency", ["USD", "EUR"])
            st.session_state.params["Strike_Date"] = st.date_input("Strike Date", today)
        with col2:
            st.session_state.params["Tenor"] = st.selectbox("Tenor", [24,18,12, 6, 3],index=2)
            st.session_state.params["Frequency"] = st.selectbox("Frequency", ["Monthly","Quarterly", "Semi-Annually", "Annually"],index=2)
            st.session_state.params["Size"] = st.number_input("Size", min_value=100000)
            
    # Payoff Section
    with st.expander("Payoff"):
        col1, col2 = st.columns(2)
        with col1:
            st.session_state.params["Coupon_Type"] = st.selectbox("Coupon Type", ["Fixed", "Conditional"])
            st.session_state.params["Downside_Type"] = st.selectbox("Downside Type", ["Put", "KI Put"],index=0)
            st.session_state.params["Reoffer"] = st.number_input("Reoffer", min_value=0)
        with col2:
            st.session_state.params["Autocall_Barrier"] = st.number_input("Autocall Barrier", min_value=0)
            st.session_state.params["Strike"] = st.number_input("Strike", min_value=0)
            if st.session_state.params["Downside_Type"] == "KI Put":
                st.session_state.params["KI_Barrier"] = st.number_input("KI Barrier", min_value=0)
    # Underlyings Section
    with st.expander("Underlyings"):
        underlyings = []

        for i in range(1, 6):
            # Creating 5 lines of input for underlyings
            st.session_state.params[f"Underlying {i}"] = st.text_input(f"Underlying {i}", key=f"underlying_{i}")

    st.session_state.valid = st.button("Validate")
    if st.session_state.valid:
        st.rerun()
    
    
if st.session_state.valid:
    
    df = pd.DataFrame([st.session_state.params])
    
    body = table(ms_pricer(df).columns,ms_pricer(df).values)
    send(body)
    
