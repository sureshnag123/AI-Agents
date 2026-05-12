import os
import logging
from dotenv import load_dotenv
from flask import Flask, render_template, request, redirect, url_for, flash, session
import requests

load_dotenv()

app = Flask(__name__)
app.secret_key = os.environ.get('FLASK_SECRET_KEY', 'change-me-in-production')

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

ZOHO_CLIENT_ID = os.environ.get('ZOHO_CLIENT_ID', '')
ZOHO_CLIENT_SECRET = os.environ.get('ZOHO_CLIENT_SECRET', '')
ZOHO_REDIRECT_URI = os.environ.get('ZOHO_REDIRECT_URI', 'http://localhost:5000/callback')


# --- UI ROUTES ---

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        zoho_api_key = request.form['zoho_api_key']
        gst_user_id = request.form['gst_user_id']
        gst_password = request.form['gst_password']
        gst_client_id = request.form.get('gst_client_id', '')
        gst_client_secret = request.form.get('gst_client_secret', '')
        date_start = request.form['date_start']
        date_end = request.form['date_end']
        try:
            result = push_invoices(zoho_api_key, gst_user_id, gst_password,
                                   gst_client_id, gst_client_secret,
                                   date_start, date_end)
            flash(f'Success: {result}', 'success')
        except Exception as e:
            logger.error('Invoice push failed: %s', e)
            flash(f'Error: {e}', 'danger')
        return redirect(url_for('index'))
    # Pre-fill Zoho access token from session if available
    zoho_token = session.get('zoho_access_token', '')
    return render_template('index.html', zoho_token=zoho_token)


@app.route('/zoho/authorize')
def zoho_authorize():
    """Redirect user to Zoho OAuth consent screen."""
    auth_url = (
        'https://accounts.zoho.com/oauth/v2/auth'
        f'?response_type=code'
        f'&client_id={ZOHO_CLIENT_ID}'
        f'&scope=ZohoSubscriptions.invoices.READ'
        f'&redirect_uri={ZOHO_REDIRECT_URI}'
        f'&access_type=offline'
    )
    return redirect(auth_url)


@app.route('/callback')
def zoho_callback():
    """Handle Zoho OAuth callback and exchange code for access token."""
    code = request.args.get('code')
    error = request.args.get('error')

    if error or not code:
        flash(f'Zoho authorization failed: {error or "no code received"}', 'danger')
        return redirect(url_for('index'))

    try:
        token_data = exchange_zoho_code(code)
        session['zoho_access_token'] = token_data.get('access_token', '')
        session['zoho_refresh_token'] = token_data.get('refresh_token', '')
        flash('Zoho authorized successfully. Access token stored for this session.', 'success')
    except Exception as e:
        logger.error('Token exchange failed: %s', e)
        flash(f'Token exchange failed: {e}', 'danger')

    return redirect(url_for('index'))


# --- BACKEND LOGIC ---

def exchange_zoho_code(code):
    response = requests.post('https://accounts.zoho.com/oauth/v2/token', params={
        'grant_type': 'authorization_code',
        'client_id': ZOHO_CLIENT_ID,
        'client_secret': ZOHO_CLIENT_SECRET,
        'redirect_uri': ZOHO_REDIRECT_URI,
        'code': code,
    })
    response.raise_for_status()
    return response.json()


def refresh_zoho_token(refresh_token):
    response = requests.post('https://accounts.zoho.com/oauth/v2/token', params={
        'grant_type': 'refresh_token',
        'client_id': ZOHO_CLIENT_ID,
        'client_secret': ZOHO_CLIENT_SECRET,
        'refresh_token': refresh_token,
    })
    response.raise_for_status()
    data = response.json()
    session['zoho_access_token'] = data.get('access_token', '')
    return data.get('access_token', '')


def fetch_sales_invoices(zoho_api_key, date_start, date_end):
    url = 'https://www.zohoapis.com/billing/v1/invoices'
    headers = {'Authorization': f'Zoho-oauthtoken {zoho_api_key}'}
    params = {'date_start': date_start, 'date_end': date_end}
    response = requests.get(url, headers=headers, params=params)
    if response.status_code == 401:
        # Token expired — try refresh if we have a refresh token
        refresh_token = session.get('zoho_refresh_token')
        if refresh_token:
            new_token = refresh_zoho_token(refresh_token)
            headers['Authorization'] = f'Zoho-oauthtoken {new_token}'
            response = requests.get(url, headers=headers, params=params)
        response.raise_for_status()
    else:
        response.raise_for_status()
    return response.json()


def map_invoices_to_gstr1(invoices):
    gstr1_data = {'invoices': []}
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


def authenticate_gst(gst_user_id, gst_password, gst_client_id='', gst_client_secret=''):
    url = 'https://api.gst.gov.in/authenticate'
    payload = {
        'username': gst_user_id,
        'password': gst_password,
    }
    if gst_client_id:
        payload['client_id'] = gst_client_id
    if gst_client_secret:
        payload['client_secret'] = gst_client_secret
    response = requests.post(url, json=payload)
    response.raise_for_status()
    data = response.json()
    return data.get('access_token') or data.get('token')


def push_gstr1_data(gstr1_data, access_token):
    url = 'https://api.gst.gov.in/gstr1/upload'
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }
    response = requests.post(url, headers=headers, json=gstr1_data)
    response.raise_for_status()
    return response.json()


def push_invoices(zoho_api_key, gst_user_id, gst_password,
                  gst_client_id, gst_client_secret, date_start, date_end):
    logger.info('Fetching invoices from Zoho: %s to %s', date_start, date_end)
    invoices = fetch_sales_invoices(zoho_api_key, date_start, date_end)
    gstr1_data = map_invoices_to_gstr1(invoices)
    logger.info('Mapped %d invoices to GSTR-1 format', len(gstr1_data['invoices']))
    access_token = authenticate_gst(gst_user_id, gst_password, gst_client_id, gst_client_secret)
    result = push_gstr1_data(gstr1_data, access_token)
    logger.info('GSTR-1 push result: %s', result)
    return result


if __name__ == '__main__':
    app.run(debug=True)
