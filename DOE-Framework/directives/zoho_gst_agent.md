# Zoho Billing to GST Portal Agent Implementation

## Purpose
Automate the process of uploading sales invoices from Zoho Billing to the GST Portal (GSTR-1).

## Steps
1. **Zoho API Setup**
   - Register your app in Zoho API Console.
   - Obtain Client ID and Client Secret.
   - Authorize your app and get the OAuth code.
   - Exchange the code for an access token.
   - Use the access token as your Zoho API Key.

2. **GST Portal API Setup**
   - Enable API access on GST portal.
   - Obtain user ID, password, and any required client credentials.

3. **Agent Workflow**
   - User enters Zoho API Key, GST credentials, and date range in the UI.
   - Agent fetches sales invoices from Zoho Billing for the selected period.
   - Agent maps invoices to GSTR-1 format.
   - Agent authenticates with GST portal and uploads invoices.
   - Agent displays upload status and errors in the UI.

## Requirements
- Zoho Billing access token (OAuth).
- GST portal API credentials.
- Python environment with Flask and requests.

## Files
- `execution/pdf_editor/app_ui.py`: Flask UI app for integration.
- `execution/zoho_token_exchange.py`: Script to exchange OAuth code for Zoho access token.
- `execution/push_gstr1_agent.py`: Backend logic for invoice upload.

## Edge Cases
- Invalid or expired Zoho access token.
- GST portal authentication failures.
- Data format mismatches between Zoho and GST.

## Improvements
- Add error handling and logging.
- Support for other GST return types.
- Secure credential storage.

## References
- [Zoho API Console](https://api-console.zoho.com/)
- [Zoho OAuth Docs](https://www.zoho.com/accounts/protocol.html)
- [GST Portal API Docs](https://www.gst.gov.in/)

---
Update this document as you refine the agent or discover new requirements.
