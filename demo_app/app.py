import os
import json
import requests
import msal
from flask import Flask, render_template, request, jsonify, url_for

app = Flask(__name__, static_folder='static', static_url_path='/static')

# Email credentials - set these in Render environment variables
EMAIL_CLIENT_ID = os.getenv('EMAIL_CLIENT_ID')
EMAIL_CLIENT_SECRET = os.getenv('EMAIL_CLIENT_SECRET')
EMAIL_TENANT_ID = os.getenv('EMAIL_TENANT_ID')
SENDER_EMAIL = os.getenv('SENDER_EMAIL', 'Tamana.Gupta@desri.com')

# Static invoice data (from your JSON file)
INVOICE_DATA = {
    "Matched_Results": [
        {
      "Worker": "Maanya Myadam",
      "Worker_Email": "maanya.myadam@desri.com",
      "Worker_Manager": "Sravya Eranti",
      "Worker_Manager_Email": "sravya.eranti@desri.com",
      "Supplier_Invoice": "290263",
      "Invoice_Number": "IN-121008",
      "Company": "TPE Alta Luna, LLC",
      "Supplier": "Beck Land & Cattle Company",
      "Invoice_amount": "4302",
      "Spend_Category_(Worktag)": "Land Rent (cash portion)",
      "Due_Date": "2026-01-01",
      "Invoice_Due_Within_(Days)": "22",
      "Aging": "16 - 30 Days"
    },
        {
      "Worker": "Maanya Myadam",
      "Worker_Email": "maanya.myadam@desri.com",
      "Worker_Manager": "Sravya Eranti",
      "Worker_Manager_Email": "sravya.eranti@desri.com",
      "Supplier_Invoice": "8099137",
      "Invoice_Number": "IN-120836",
      "Company": "Assembly Solar III, L.L.C.",
      "Supplier": "Stoel Rives LLP",
      "Invoice_amount": "8146.12",
      "Spend_Category_(Worktag)": "Legal Costs",
      "Due_Date": "2026-01-02",
      "Invoice_Due_Within_(Days)": "23",
      "Aging": "16 - 30 Days"
    },
    {
      "Worker": "Maanya Myadam",
      "Worker_Email": "maanya.myadam@desri.com",
      "Worker_Manager": "Sravya Eranti",
      "Worker_Manager_Email": "sravya.eranti@desri.com",
      "Supplier_Invoice": "700788032286_12022025",
      "Invoice_Number": "IN-120967",
      "Company": "Willow Springs Solar, LLC",
      "Supplier": "Southern California Edison",
      "Invoice_amount": "27027.35",
      "Spend_Category_(Worktag)": "Electric Utilities",
      "Due_Date": "2025-12-22",
      "Invoice_Due_Within_(Days)": "12",
      "Aging": "6 - 15 Days"
    },
    {
      "Worker": "Maanya Myadam",
      "Worker_Email": "maanya.myadam@desri.com",
      "Worker_Manager": "Sravya Eranti",
      "Worker_Manager_Email": "sravya.eranti@desri.com",
      "Supplier_Invoice": "108237",
      "Invoice_Number": "IN-121216",
      "Company": "TPE Alta Luna, LLC",
      "Supplier": "Radian Generation LLC",
      "Invoice_amount": "7222.5",
      "Spend_Category_(Worktag)": "Project IT Tools / Services (DACC/SCADA)",
      "Due_Date": "2025-12-31",
      "Invoice_Due_Within_(Days)": "21",
      "Aging": "16 - 30 Days"
    },
    {
      "Worker": "Lehar Vinnakota",
      "Worker_Email": "lehar.vinnakota@desri.com",
      "Worker_Manager": "Sravya Eranti",
      "Worker_Manager_Email": "sravya.eranti@desri.com",
      "Supplier_Invoice": "108373",
      "Invoice_Number": "IN-121315",
      "Company": "Red Horse III, LLC",
      "Supplier": "Radian Generation LLC",
      "Invoice_amount": "5775",
      "Spend_Category_(Worktag)": "Project IT Tools / Services (DACC/SCADA)",
      "Due_Date": "2025-12-31",
      "Invoice_Due_Within_(Days)": "21",
      "Aging": "16 - 30 Days"
    },
    {
      "Worker": "Anubhav Maskara",
      "Worker_Email": "anubhav.maskara@desri.com",
      "Worker_Manager": "Navneet Mohata",
      "Worker_Manager_Email": "navneet.mohata@desri.com",
      "Supplier_Invoice": "5884636",
      "Invoice_Number": "IN-120666",
      "Company": "Dolet Hills Solar, LLC",
      "Supplier": "Arthur J Gallagher Risk Management Services, Inc.",
      "Invoice_amount": "147397.69",
      "Spend_Category_(Worktag)": "Project Insurance",
      "Due_Date": "2025-11-26",
      "Invoice_Due_Within_(Days)": "-14",
      "Aging": "0 - 5 Days"
    },
    {
      "Worker": "Anubhav Maskara",
      "Worker_Email": "anubhav.maskara@desri.com",
      "Worker_Manager": "Navneet Mohata",
      "Worker_Manager_Email": "navneet.mohata@desri.com",
      "Supplier_Invoice": "Blu elk I_24",
      "Invoice_Number": "IN-120609",
      "Company": "Blue Elk Solar I, LLC",
      "Supplier": "Primoris Renewable Energy, Inc",
      "Invoice_amount": "62251.64",
      "Spend_Category_(Worktag)": "EPC Costs",
      "Due_Date": "2025-12-25",
      "Invoice_Due_Within_(Days)": "15",
      "Aging": "6 - 15 Days"
    },
    {
      "Worker": "Anubhav Maskara",
      "Worker_Email": "anubhav.maskara@desri.com",
      "Worker_Manager": "Navneet Mohata",
      "Worker_Manager_Email": "navneet.mohata@desri.com",
      "Supplier_Invoice": "2343628",
      "Invoice_Number": "IN-120117",
      "Company": "Dolet Hills Solar, LLC",
      "Supplier": "WEG Transformers USA LLC",
      "Invoice_amount": "7500",
      "Spend_Category_(Worktag)": "Storage/Handling Fees",
      "Due_Date": "2025-12-24",
      "Invoice_Due_Within_(Days)": "14",
      "Aging": "6 - 15 Days"
    },
    {
      "Worker": "Varun Pansari",
      "Worker_Email": "varun.pansari@desri.com",
      "Worker_Manager": "Navneet Mohata",
      "Worker_Manager_Email": "navneet.mohata@desri.com",
      "Supplier_Invoice": "04584",
      "Invoice_Number": "IN-120894",
      "Company": "Santa Teresa Solar, LLC",
      "Supplier": "Luminary Logistics Solutions LLC",
      "Invoice_amount": "2995",
      "Spend_Category_(Worktag)": "Storage/Handling Fees",
      "Due_Date": "2025-06-08",
      "Invoice_Due_Within_(Days)": "-185",
      "Aging": "0 - 5 Days"
    },
    {
      "Worker": "Varun Pansari",
      "Worker_Email": "varun.pansari@desri.com",
      "Worker_Manager": "Navneet Mohata",
      "Worker_Manager_Email": "navneet.mohata@desri.com",
      "Supplier_Invoice": "PI24-E000928-001",
      "Invoice_Number": "IN-103110",
      "Company": "Santa Teresa Storage, LLC",
      "Supplier": "SENERGY TECHNICAL SERVICES (USA) LLC",
      "Invoice_amount": "7500",
      "Spend_Category_(Worktag)": "Other Professional Services",
      "Due_Date": "2024-12-11",
      "Invoice_Due_Within_(Days)": "-364",
      "Aging": "0 - 5 Days"
    },
    {
      "Worker": "Varun Pansari",
      "Worker_Email": "varun.pansari@desri.com",
      "Worker_Manager": "Navneet Mohata",
      "Worker_Manager_Email": "navneet.mohata@desri.com",
      "Supplier_Invoice": "1148568",
      "Invoice_Number": "IN-121078",
      "Company": "Hunter Solar, LLC - CO",
      "Supplier": "John H. Hyatt",
      "Invoice_amount": "66954",
      "Spend_Category_(Worktag)": "Land Rent (cash portion)",
      "Due_Date": "2026-01-16",
      "Invoice_Due_Within_(Days)": "37",
      "Aging": "31 - 60 Days"
    },
    {
      "Worker": "Garima Surana",
      "Worker_Email": "garima.surana@desri.com",
      "Worker_Manager": "Sravya Eranti",
      "Worker_Manager_Email": "sravya.eranti@desri.com",
      "Supplier_Invoice": "106142_a2",
      "Invoice_Number": "IN-120825",
      "Company": "Assembly Solar II, L.L.C.",
      "Supplier": "Radian Generation LLC",
      "Invoice_amount": "4196.12",
      "Spend_Category_(Worktag)": "Regulatory Compliance (Non-Legal)",
      "Due_Date": "2025-12-31",
      "Invoice_Due_Within_(Days)": "21",
      "Aging": "16 - 30 Days"
    },
    {
      "Worker": "Garima Surana",
      "Worker_Email": "garima.surana@desri.com",
      "Worker_Manager": "Sravya Eranti",
      "Worker_Manager_Email": "sravya.eranti@desri.com",
      "Supplier_Invoice": "106142",
      "Invoice_Number": "IN-120823",
      "Company": "Assembly Solar I, L.L.C.",
      "Supplier": "Radian Generation LLC",
      "Invoice_amount": "1907.33",
      "Spend_Category_(Worktag)": "Regulatory Compliance (Non-Legal)",
      "Due_Date": "2025-12-31",
      "Invoice_Due_Within_(Days)": "21",
      "Aging": "16 - 30 Days"
    },
        {
      "Worker": "Garima Surana",
      "Worker_Email": "garima.surana@desri.com",
      "Worker_Manager": "Sravya Eranti",
      "Worker_Manager_Email": "sravya.eranti@desri.com",
      "Supplier_Invoice": "106142_a2",
      "Invoice_Number": "IN-120825",
      "Company": "Assembly Solar II, L.L.C.",
      "Supplier": "Radian Generation LLC",
      "Invoice_amount": "4196.12",
      "Spend_Category_(Worktag)": "Regulatory Compliance (Non-Legal)",
      "Due_Date": "2025-12-31",
      "Invoice_Due_Within_(Days)": "21",
      "Aging": "16 - 30 Days"
    }
    ]
}

def create_invoice_table_html(invoices):
    """Create HTML table for invoices."""
    if not invoices:
        return "<p>No invoices found.</p>"
    
    html = """
    <table border="1" cellpadding="8" cellspacing="0" style="border-collapse: collapse; width: 100%; font-family: Arial, sans-serif;">
        <thead>
            <tr style="background-color: #4A4A4A; color: white;">
                <th style="text-align: left; padding: 10px;">Aging</th>
                <th style="text-align: left; padding: 10px;">Invoice Number</th>
                <th style="text-align: left; padding: 10px;">Company</th>
                <th style="text-align: left; padding: 10px;">Supplier</th>
                <th style="text-align: right; padding: 10px;">Invoice Amount</th>
                <th style="text-align: left; padding: 10px;">Spend Category</th>
                <th style="text-align: left; padding: 10px;">Due Date</th>
                <th style="text-align: right; padding: 10px;">Days Until Due</th>
            </tr>
        </thead>
        <tbody>
    """
    
    for inv in invoices:
        html += f"""
            <tr>
                <td style="padding: 8px;">{inv.get('Aging', '')}</td>
                <td style="padding: 8px;">{inv.get('Invoice_Number', '')}</td>
                <td style="padding: 8px;">{inv.get('Company', '')}</td>
                <td style="padding: 8px;">{inv.get('Supplier', '')}</td>
                <td style="text-align: right; padding: 8px;">${inv.get('Invoice_amount', '0')}</td>
                <td style="padding: 8px;">{inv.get('Spend_Category_(Worktag)', '')}</td>
                <td style="padding: 8px;">{inv.get('Due_Date', '')}</td>
                <td style="text-align: right; padding: 8px;">{inv.get('Invoice_Due_Within_(Days)', '')}</td>
            </tr>
        """
    
    html += "</tbody></table>"
    return html

def send_email_graph(to_email, subject, body_html, cc_emails=None):
    """Send email using Microsoft Graph API."""
    authority = f"https://login.microsoftonline.com/{EMAIL_TENANT_ID}"
    
    app_msal = msal.ConfidentialClientApplication(
        client_id=EMAIL_CLIENT_ID,
        client_credential=EMAIL_CLIENT_SECRET,
        authority=authority
    )
    
    result = app_msal.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    
    if "access_token" not in result:
        raise Exception(f"Failed to get token: {result.get('error_description', 'Unknown error')}")
    
    access_token = result["access_token"]
    
    message = {
        "message": {
            "subject": subject,
            "body": {"contentType": "HTML", "content": body_html},
            "toRecipients": [{"emailAddress": {"address": to_email}}]
        }
    }
    
    if cc_emails:
        message["message"]["ccRecipients"] = [{"emailAddress": {"address": e}} for e in cc_emails]
    
    response = requests.post(
        f"https://graph.microsoft.com/v1.0/users/{SENDER_EMAIL}/sendMail",
        headers={"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"},
        json=message
    )
    
    if response.status_code == 202:
        return {"success": True}
    else:
        raise Exception(f"Failed: {response.status_code} - {response.text}")

@app.route('/')
def index():
    invoices = INVOICE_DATA.get('Matched_Results', [])
    
    # Calculate total amount
    total_amount = sum(float(inv.get('Invoice_amount', 0)) for inv in invoices)
    
    # Group by worker
    workers = {}
    for inv in invoices:
        email = inv.get('Worker_Email', '')
        if email not in workers:
            workers[email] = {
                'name': inv.get('Worker', ''),
                'email': email,
                'manager': inv.get('Worker_Manager', ''),
                'manager_email': inv.get('Worker_Manager_Email', ''),
                'invoices': []
            }
        workers[email]['invoices'].append(inv)
    
    return render_template('index.html', workers=workers, invoices=invoices, total_amount=total_amount)

@app.route('/send-email', methods=['POST'])
def send_email():
    data = request.json
    worker_email = data.get('worker_email')
    
    # Find worker's invoices
    invoices = [inv for inv in INVOICE_DATA['Matched_Results'] if inv.get('Worker_Email') == worker_email]
    
    if not invoices:
        return jsonify({'success': False, 'error': 'No invoices found for worker'})
    
    worker_name = invoices[0].get('Worker', '')
    manager_email = invoices[0].get('Worker_Manager_Email', '')
    
    subject = f"Pending Invoices for Approval - {len(invoices)} Invoice(s)"
    body = f"""
    <html>
    <body style="font-family: Arial, sans-serif; line-height: 1.6;">
        <p>Hi {worker_name},</p>
        <p>Please find below the invoices pending in your approval queue:</p>
        <br>
        {create_invoice_table_html(invoices)}
        <br>
        <p>Kindly review and approve at your earliest convenience.</p>
        <p>Thanks,<br>Accounts Payable</p>
    </body>
    </html>
    """
    
    try:
        cc_list = [manager_email] if manager_email and manager_email != worker_email else None
        send_email_graph(worker_email, subject, body, cc_list)
        return jsonify({'success': True, 'message': f'Email sent to {worker_email}'})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})

@app.route('/send-all', methods=['POST'])
def send_all():
    results = {'sent': 0, 'failed': 0, 'errors': []}
    
    # Group by worker
    workers = {}
    for inv in INVOICE_DATA['Matched_Results']:
        email = inv.get('Worker_Email', '')
        if email not in workers:
            workers[email] = []
        workers[email].append(inv)
    
    for worker_email, invoices in workers.items():
        worker_name = invoices[0].get('Worker', '')
        manager_email = invoices[0].get('Worker_Manager_Email', '')
        
        subject = f"Pending Invoices for Approval - {len(invoices)} Invoice(s)"
        body = f"""
        <html>
        <body style="font-family: Arial, sans-serif; line-height: 1.6;">
            <p>Hi {worker_name},</p>
            <p>Please find below the invoices pending in your approval queue:</p>
            <br>
            {create_invoice_table_html(invoices)}
            <br>
            <p>Kindly review and approve at your earliest convenience.</p>
            <p>Thanks,<br>Accounts Payable</p>
        </body>
        </html>
        """
        
        try:
            cc_list = [manager_email] if manager_email and manager_email != worker_email else None
            send_email_graph(worker_email, subject, body, cc_list)
            results['sent'] += 1
        except Exception as e:
            results['failed'] += 1
            results['errors'].append({'email': worker_email, 'error': str(e)})
    
    return jsonify(results)

if __name__ == '__main__':
    app.run(debug=True, port=5000)




