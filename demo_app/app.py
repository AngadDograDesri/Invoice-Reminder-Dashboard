import os
import json
import requests
import msal
from flask import Flask, render_template, request, jsonify

app = Flask(__name__)

# Email credentials - set these in Render environment variables
EMAIL_CLIENT_ID = os.getenv('EMAIL_CLIENT_ID')
EMAIL_CLIENT_SECRET = os.getenv('EMAIL_CLIENT_SECRET')
EMAIL_TENANT_ID = os.getenv('EMAIL_TENANT_ID')
SENDER_EMAIL = os.getenv('SENDER_EMAIL', 'Angad.Dogra@desri.com')

# Static invoice data (from your JSON file)
INVOICE_DATA = {
    "Matched_Results": [
        {
            "Worker": "Angad Dogra",
            "Worker_Email": "angad.dogra@desri.com",
            "Worker_Manager": "Vikas Agrawal",
            "Worker_Manager_Email": "vikas.agrawal@desri.com",
            "Supplier_Invoice": "373809",
            "Invoice_Number": "IN-117893",
            "Company": "Rocking R Solar, LLC",
            "Supplier": "EDKO LLC",
            "Invoice_amount": "44166.67",
            "Spend_Category_(Worktag)": "Vegetation Management",
            "Due_Date": "2025-12-28",
            "Invoice_Due_Within_(Days)": "33",
            "Aging": "31 - 60 Days"
        },
        {
            "Worker": "Angad Dogra",
            "Worker_Email": "angad.dogra@desri.com",
            "Worker_Manager": "Vikas Agrawal",
            "Worker_Manager_Email": "vikas.agrawal@desri.com",
            "Supplier_Invoice": "09262025_T",
            "Invoice_Number": "IN-118715",
            "Company": "DESRI Drew Solar Financing Holdings, L.L.C.",
            "Supplier": "Societe Generale",
            "Invoice_amount": "5000",
            "Spend_Category_(Worktag)": "LC Fees/Interest",
            "Due_Date": "2025-11-20",
            "Invoice_Due_Within_(Days)": "10",
            "Aging": "0 - 5 Days"
        },
        {
            "Worker": "Angad Dogra",
            "Worker_Email": "angad.dogra@desri.com",
            "Worker_Manager": "Vikas Agrawal",
            "Worker_Manager_Email": "vikas.agrawal@desri.com",
            "Supplier_Invoice": "09262025_Fin Hold LC Fees_Test",
            "Invoice_Number": "IN-118714",
            "Company": "DESRI Drew Solar Financing Holdings, L.L.C.",
            "Supplier": "Societe Generale",
            "Invoice_amount": "10397.26",
            "Spend_Category_(Worktag)": "LC Fees/Interest",
            "Due_Date": "2025-11-20",
            "Invoice_Due_Within_(Days)": "-5",
            "Aging": "0 - 5 Days"
        },
        {
            "Worker": "Angad Dogra",
            "Worker_Email": "angad.dogra@desri.com",
            "Worker_Manager": "Vikas Agrawal",
            "Worker_Manager_Email": "vikas.agrawal@desri.com",
            "Supplier_Invoice": "09232025_Drew A1_t2",
            "Invoice_Number": "IN-118713",
            "Company": "DESRI A1 Drew Solar Borrower, L.L.C.",
            "Supplier": "Societe Generale",
            "Invoice_amount": "5000",
            "Spend_Category_(Worktag)": "Commitment Fees",
            "Due_Date": "2025-11-20",
            "Invoice_Due_Within_(Days)": "-5",
            "Aging": "0 - 5 Days"
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

