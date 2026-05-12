import requests
import json
import os
from dotenv import load_dotenv

load_dotenv()

# --- CONFIGURATION (loaded from .env) ---
ZOHO_API_KEY = os.environ.get('ZOHO_API_KEY', '')
GST_USER_ID = os.environ.get('GST_USER_ID', '')
GST_PASSWORD = os.environ.get('GST_PASSWORD', '')
GST_CLIENT_ID = os.environ.get('GST_CLIENT_ID', '')
GST_CLIENT_SECRET = os.environ.get('GST_CLIENT_SECRET', '')

# STEP 1: Fetch sales invoices from Zoho Billing for February
def fetch_sales_invoices_feb():
    url = 'https://www.zohoapis.com/billing/v1/invoices'
    headers = {'Authorization': f'Zoho-oauthtoken {ZOHO_API_KEY}'}
    # 2026 is not a leap year — February ends on the 28th
    params = {
        'date_start': '2026-02-01',
        'date_end': '2026-02-28'
    }
    response = requests.get(url, headers=headers, params=params)
    response.raise_for_status()
    return response.json()

# STEP 1b: Map sales invoices to GSTR-1 format
def map_invoices_to_gstr1(invoices):
    # This is a simplified mapping. Adjust as per GST API schema.
    gstr1_data = {
        'invoices': []
    }
    for inv in invoices.get('invoices', []):
        gstr1_data['invoices'].append({
            'invoice_number': inv.get('invoice_number'),
            'invoice_date': inv.get('date'),
            'customer_gstin': inv.get('customer_gstin'),
            'total_amount': inv.get('total'),
            'tax_amount': inv.get('tax_total'),
            'items': inv.get('line_items', [])
        })
    return gstr1_data

# STEP 2: Authenticate with GST portal

# STEP 2: Authenticate with GST portal using user credentials
def authenticate_gst():
    url = 'https://api.gst.gov.in/authenticate'  # Replace with actual GST authentication endpoint
    payload = {
        'username': GST_USER_ID,
        'password': GST_PASSWORD,
        'client_id': GST_CLIENT_ID,
        'client_secret': GST_CLIENT_SECRET
    }
    response = requests.post(url, json=payload)
    response.raise_for_status()
    data = response.json()
    # Adjust parsing as per GST API response
    return data.get('access_token') or data.get('token')

# STEP 3: Push GSTR-1 data to GST portal
def push_gstr1_data(gstr1_data, access_token):
    url = 'https://api.gst.gov.in/gstr1/upload'
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }
    response = requests.post(url, headers=headers, json=gstr1_data)
    response.raise_for_status()
    return response.json()


if __name__ == '__main__':
    try:
        invoices = fetch_sales_invoices_feb()
        gstr1_data = map_invoices_to_gstr1(invoices)
        access_token = authenticate_gst()
        result = push_gstr1_data(gstr1_data, access_token)
        print('GSTR-1 Data pushed successfully:', result)
    except Exception as e:
        print('Error:', e)
