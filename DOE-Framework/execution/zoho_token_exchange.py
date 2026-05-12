import requests
import os
from dotenv import load_dotenv

load_dotenv()

client_id = os.environ.get('ZOHO_CLIENT_ID', '')
client_secret = os.environ.get('ZOHO_CLIENT_SECRET', '')
redirect_uri = os.environ.get('ZOHO_REDIRECT_URI', 'http://localhost:5000/callback')
code = os.environ.get('ZOHO_AUTH_CODE', '')  # Set ZOHO_AUTH_CODE in .env after getting it

url = "https://accounts.zoho.com/oauth/v2/token"
params = {
    "grant_type": "authorization_code",
    "client_id": client_id,
    "client_secret": client_secret,
    "redirect_uri": redirect_uri,
    "code": code
}

response = requests.post(url, params=params)
print(response.json())
