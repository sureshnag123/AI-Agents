# Business Planning Agent Directive

> Standard Operating Procedure for creating comprehensive business plans with market research, financial projections, and strategic analysis.

## Goal
Automate the creation of professional business plans by combining market research, competitive analysis, SWOT analysis, financial projections, and business model canvases into cohesive documents stored in Google Docs and Sheets.

## Inputs
| Input | Required | Description |
|-------|----------|-------------|
| Company Name | Yes | Name of the business |
| Industry | Yes | Industry sector (e.g., "SaaS", "E-commerce", "Healthcare") |
| Business Description | Yes | 2-3 sentence description of the business |
| Target Market | No | Description of target customers |
| Revenue Model | No | How the business makes money |
| Initial Investment | No | Starting capital amount |
| Competitors | No | Known competitor names |
| Google Drive Folder ID | No | Folder to store output documents |

## Tools/Scripts

### Market Research
- `serp_market_research.py` - Search Google, analyze competitors, track trends, get news
  - Modes: `search`, `competitors`, `trends`, `news`
  - Requires: `SERPAPI_API_KEY`

### Business Analysis
- `generate_business_plan.py` - Generate SWOT, financials, canvas, compile plans
  - Modes: `swot`, `financials`, `canvas`, `compile`
  - Requires: `OPENAI_API_KEY` or `ANTHROPIC_API_KEY`

### Google Docs
- `create_google_doc.py` - Create new documents
- `update_google_doc.py` - Update existing documents (append, prepend, replace)
  - Requires: `credentials.json` and OAuth flow

### Google Sheets
- `read_sheet.py` - Read data from sheets
- `update_sheet.py` - Create or update sheets
- `append_to_sheet.py` - Append rows to sheets
  - Requires: `credentials.json` and OAuth flow

## Workflow

### Phase 1: Market Research
1. **Industry Search**: Run market research to understand the industry landscape
   ```bash
   python execution/serp_market_research.py --mode search --query "[industry] market size trends 2024" --output .tmp/market_research.json
   ```

2. **Competitor Analysis**: Identify and analyze competitors
   ```bash
   python execution/serp_market_research.py --mode competitors --query "[company]" --industry "[industry]" --output .tmp/competitors.json
   ```

3. **Industry Trends**: Get trend data
   ```bash
   python execution/serp_market_research.py --mode trends --query "[industry]" --output .tmp/trends.json
   ```

4. **News Search**: Get recent industry news
   ```bash
   python execution/serp_market_research.py --mode news --query "[industry] news" --output .tmp/news.json
   ```

### Phase 2: Business Analysis
5. **SWOT Analysis**: Generate SWOT using market data
   ```bash
   python execution/generate_business_plan.py --mode swot --company "[company]" --industry "[industry]" --description "[description]" --market-data .tmp/market_research.json --output .tmp/swot.json
   ```

6. **Business Model Canvas**: Create the business model
   ```bash
   python execution/generate_business_plan.py --mode canvas --company "[company]" --industry "[industry]" --description "[description]" --target-market "[target]" --output .tmp/canvas.json
   ```

7. **Financial Projections**: Generate 5-year projections
   ```bash
   python execution/generate_business_plan.py --mode financials --company "[company]" --industry "[industry]" --revenue-model "[model]" --investment [amount] --years 5 --output .tmp/financials.json
   ```

### Phase 3: Compile and Deliver
8. **Compile Business Plan**: Combine all components
   ```bash
   python execution/generate_business_plan.py --mode compile --company "[company]" --industry "[industry]" --description "[description]" --market-data .tmp/market_research.json --swot-data .tmp/swot.json --financial-data .tmp/financials.json --canvas-data .tmp/canvas.json --output .tmp/business_plan.json
   ```

9. **Create Google Doc**: Create the business plan document
   ```bash
   python execution/create_google_doc.py --title "[Company] Business Plan" --content-file .tmp/business_plan_content.txt --folder-id [folder_id] --output .tmp/doc_result.json
   ```

10. **Create Financial Spreadsheet**: Export financials to Google Sheets
    ```bash
    python execution/update_sheet.py --title "[Company] Financial Projections" --data-file .tmp/financials.json
    ```

## Outputs
| Output | Location | Description |
|--------|----------|-------------|
| Business Plan Document | Google Docs | Full business plan with all sections |
| Financial Projections | Google Sheets | 5-year financial model |
| Market Research Data | `.tmp/` | JSON files with research data |
| SWOT Analysis | `.tmp/swot.json` | Structured SWOT data |
| Business Model Canvas | `.tmp/canvas.json` | BMC components |

## Edge Cases and Error Handling

### SerpAPI Rate Limits
- Free tier: 100 searches/month
- If rate limited, wait and retry or use cached data
- Consider upgrading for heavy usage

### Google OAuth
- First run requires browser authentication
- Token refreshes automatically after initial auth
- If token expires, delete `token.json` and re-authenticate

### LLM API Errors
- Supports both OpenAI and Anthropic
- Falls back to available provider
- Large documents may need to be chunked

### Missing Data
- If market research fails, proceed with available data
- Financial projections can be generated with minimal inputs
- Document which data sources were unavailable

## Best Practices

1. **Always save intermediate outputs** to `.tmp/` for debugging and reuse
2. **Run market research first** - it enriches all other analyses
3. **Review AI-generated content** before finalizing documents
4. **Use specific industry terms** in queries for better results
5. **Include competitor names** when known for better analysis
6. **Back up Google Docs links** - save document IDs

## Example Usage

### Quick Business Plan
```bash
# 1. Market research
python execution/serp_market_research.py --mode search --query "AI automation agency market 2024" --output .tmp/market.json

# 2. Generate SWOT
python execution/generate_business_plan.py --mode swot --company "AutomateAI" --industry "AI Automation" --description "AI-powered business process automation agency" --market-data .tmp/market.json --output .tmp/swot.json

# 3. Generate financials
python execution/generate_business_plan.py --mode financials --company "AutomateAI" --industry "AI Automation" --revenue-model "Monthly retainer + project fees" --investment 50000 --output .tmp/financials.json

# 4. Compile plan
python execution/generate_business_plan.py --mode compile --company "AutomateAI" --industry "AI Automation" --description "AI-powered business process automation agency" --swot-data .tmp/swot.json --financial-data .tmp/financials.json --market-data .tmp/market.json --output .tmp/plan.json

# 5. Create Google Doc
python execution/create_google_doc.py --title "AutomateAI Business Plan 2024" --content-file .tmp/plan_content.txt
```

## Required Environment Variables
```env
SERPAPI_API_KEY=your_serpapi_key
OPENAI_API_KEY=your_openai_key
ANTHROPIC_API_KEY=your_anthropic_key  # Optional, alternative to OpenAI
GOOGLE_APPLICATION_CREDENTIALS=credentials.json
```

## Dependencies
```
google-api-python-client
google-auth-httplib2
google-auth-oauthlib
google-search-results  # SerpAPI
openai
anthropic
python-dotenv
gspread
```
