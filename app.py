import streamlit as st
import pandas as pd
from jinja2 import Environment, FileSystemLoader
from weasyprint import HTML
import os
import base64
import calendar
from datetime import datetime
import pickle
from num2words import num2words
from urllib.parse import quote
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build


# Email mapping
email_map = {
    "john doe": "john@example.com",
    "jane smith": "jane@example.com",
    "kavidhesh g": "kavidheshg@gmail.com",
    # Add more
}

SCOPES = ['https://www.googleapis.com/auth/gmail.send']

def gmail_authenticate():
    creds = None
    if os.path.exists('token.pkl'):
        with open('token.pkl', 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.pkl', 'wb') as token:
            pickle.dump(creds, token)
    return build('gmail', 'v1', credentials=creds)

def send_email_with_attachment_bytes(service, to, subject, body_text, attachment_bytes, filename):
    message = MIMEMultipart()
    message['to'] = to
    message['subject'] = subject

    message.attach(MIMEText(body_text, 'plain'))

    part = MIMEBase('application', 'pdf')
    part.set_payload(attachment_bytes)
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename="{filename}"')
    message.attach(part)

    raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode()
    return service.users().messages().send(userId='me', body={'raw': raw_message}).execute()


def count_week_offs(month_str):
    try:
        dt = datetime.strptime(month_str.strip(), "%B %Y")
        year = dt.year
        month = dt.month

        cal = calendar.Calendar()
        count = sum(
            1 for day, weekday in cal.itermonthdays2(year, month)
            if day != 0 and weekday in (5, 6)  # Saturday = 5, Sunday = 6
        )
        return count
    except Exception as e:
        return "-"
    
def safe_round(val):
    try:
        return round(float(val))
    except (ValueError, TypeError):
        return 0

def safe_format(val):
    try:
        num = float(val)
        return f"{int(round(num)):,}"
    except (ValueError, TypeError):
        return "0"


    
def send_payslips(df_raw, payslip_month):
    new_columns = [
        "S_no", "Emp_Code", "Name", "Budget_Code", "DOJ", "Location", "Department",
        "Designation", "Category", "Paid_Days",
        "E_Basic", "E_DA", "E_HRA", "E_CONV", "E_Medical", "E_Special_Allowance",
        "E_GROSS", "CTC",
        "P_Basic", "P_DA", "P_HRA", "P_CONV", "P_Medical", "P_Special",
        "E_Arrears", "E_Bonus", "E_Gross_Earnings",
        "D_PF_12", "D_ESI", "D_Prof_Tax", "D_Salary_Loan", "D_Advance", "D_LWF",
        "D_Other_Ded_Food", "D_IT", "D_Mobile_Deduction", "D_Gross_Deduction",
        "Net_Take_Home",
        "PF_Basic", "PF_Limit_Basic", "PF_Gross", "PF_EPS_8_33", "PF_EPF_3_67",
        "PF_Total", "PF_AC2", "PF_AC21", "PF_AC22", "PF_Total_Contribution",
        "ESI_Employer", "ESI_Total"
    ]

    df_raw.columns = new_columns
    df_cleaned = df_raw[1:].reset_index(drop=True)
    df_cleaned = df_cleaned.astype(str).fillna("")
    week_off_total = count_week_offs(payslip_month)
    

    
    env = Environment(loader=FileSystemLoader("."))
    template = env.get_template("payslip.html")

    logo_file_path = "N:/payslip automation/logo.png"
    logo_url = "file:///" + quote(logo_file_path.replace("\\", "/"))

    service = gmail_authenticate()

    for _, row in df_cleaned.iterrows():
        name_key = str(row["Name"]).strip().lower()
        recipient_email = email_map.get(name_key)

        if not recipient_email:
            st.warning(f"No email found for {row['Name']}, skipping.")
            continue
        try:
            doj = pd.to_datetime(row["DOJ"]).strftime("%d-%m-%Y")
        except:
            doj = row["DOJ"]

        net_pay_amount_raw = safe_round(row["Net_Take_Home"])
        net_pay_words = f"Rupees {num2words(net_pay_amount_raw, to='cardinal', lang='en_IN').title()} Only"
        net_pay_amount=safe_format(net_pay_amount_raw)

        paid_days = safe_round(row["Paid_Days"])
        present_days = paid_days - week_off_total


        html_out = template.render(
            payslip_month=payslip_month,
            Emp_Code=row["Emp_Code"],
            Name=row["Name"],
            Designation=row["Designation"],
            DOJ=doj,
            paid_days=paid_days,
            present_days=present_days,
            week_off=week_off_total,
            UAN_No="-",
            ESIC_No="-",
            leave_balance="-",
            E_Basic = safe_format(row["E_Basic"]),
            E_HRA = safe_format(row["E_HRA"]),
            E_CONV = safe_format(row["E_CONV"]),
            E_Medical = safe_format(row["E_Medical"]),
            E_Special_Allowance = safe_format(row["E_Special_Allowance"]),
            P_Basic = safe_format(row["P_Basic"]),
            P_HRA = safe_format(row["P_HRA"]),
            P_CONV = safe_format(row["P_CONV"]),
            P_Medical = safe_format(row["P_Medical"]),
            P_Special = safe_format(row["P_Special"]),
            D_PF_12 = safe_format(row["D_PF_12"]),
            D_Prof_Tax = safe_format(row["D_Prof_Tax"]),
            D_LWF = safe_format(row["D_LWF"]),
            D_IT = safe_format(row["D_IT"]),
            D_Other_Ded_Food = safe_format(row["D_Other_Ded_Food"]),
            total_earnings_gross = safe_format(row["E_Gross_Earnings"]),
            total_earnings_actual = safe_format(row["E_Gross_Earnings"]),
            total_deductions_gross = safe_format(row["D_Gross_Deduction"]),
            total_deductions_actual = safe_format(row["D_Gross_Deduction"]),
            net_pay=net_pay_amount,
            net_pay_words=net_pay_words,
            logo_path=logo_url
        )
        pdf_bytes = HTML(string=html_out, base_url=".").write_pdf()
        subject = f"Payslip for {payslip_month}"
        body = f"Dear {row['Name']},\n\nPlease find attached your payslip for {payslip_month}.\n\nBest regards,\nHR Team"

        try:
            send_email_with_attachment_bytes(
                service,
                to=recipient_email,
                subject=subject,
                body_text=body,
                attachment_bytes=pdf_bytes,
                filename=f"{row['Name']}_{payslip_month}.pdf"
            )
            st.success(f"✅ Sent to {row['Name']} ({recipient_email})")
        except Exception as e:
            st.error(f"❌ Failed to send to {row['Name']} ({recipient_email}): {e}")

# Streamlit UI
st.title("📤 Payroll Payslip Sender")

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])
payslip_month = st.text_input("Enter Payslip Month (e.g., May 2025)", placeholder="e.g., May 2025")

if uploaded_file:
    df = pd.read_excel(uploaded_file, skiprows=6, header=None)
    st.dataframe(df.head(25))

    if st.button("📧 Send Emails"):
        if not payslip_month.strip():
            st.error("❗ Please enter the payslip month before sending emails.")
        else:
            with st.spinner("Sending payslips..."):
                send_payslips(df, payslip_month)

