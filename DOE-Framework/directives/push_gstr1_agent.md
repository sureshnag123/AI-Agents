# Push GSTR-1 Data Agent

## Purpose
Automate the process of extracting GSTR-1 data from Zoho Billing and submitting it to the GST portal via API.

## Steps
1. Export GSTR-1 data from Zoho Billing (JSON/Excel).
2. Format data as per GST API requirements.
3. Authenticate with GST portal (using API credentials, OTP, or e-sign).
4. Push GSTR-1 data to GST portal.
5. Handle response and errors.

## Requirements
- GST API access (credentials, client ID/secret, registration).
- Zoho Billing API access (for data extraction).
- Python environment with requests and any required libraries.

## Edge Cases
- Data format mismatches.
- Authentication failures.
- API rate limits or downtime.

## Script Location
execution/push_gstr1_agent.py

## Improvements
Update this directive as you discover new API constraints or workflow optimizations.