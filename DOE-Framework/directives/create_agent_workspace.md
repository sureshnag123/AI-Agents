# Create Agent Workspace

## Goal
Generate a complete, standalone agent workspace in `outputs/` that can be copied to a separate location for independent development.

## When to Use
- User asks to "create a new agent" or "set up an agent workspace"
- User wants to start a new project based on a specific capability (lead gen, email, etc.)
- User wants a fresh DOE Framework workspace for a new use case

## Inputs
| Input | Required | Description |
|-------|----------|-------------|
| Agent Name | Yes | Human-readable name for the workspace (e.g., "My Lead Gen Agent") |
| Agent Type | Yes | One of: `lead_generation`, `email_automation`, `freelance_proposals`, `video_editing`, `crm_integration`, `full_stack`, `custom` |
| Additional Scripts | No | Extra scripts to include beyond the type's defaults |
| Additional Directives | No | Extra directives to include |
| Additional Packages | No | Extra Python packages to include |

## Available Agent Types

### lead_generation
- **Purpose**: Scrapes, enriches, and manages leads from Google Maps, SERP, and other sources
- **Includes**: Google Maps scraping, Apify integration, email enrichment, Google Sheets
- **API Keys Needed**: APIFY_API_TOKEN, GOOGLE_SHEETS_CREDENTIALS_FILE, OPENAI_API_KEY

### email_automation
- **Purpose**: Automates cold email campaigns and auto-replies via Instantly.ai
- **Includes**: Instantly API integration, welcome emails, campaign creation
- **API Keys Needed**: INSTANTLY_API_KEY, OPENAI_API_KEY

### freelance_proposals
- **Purpose**: Scrapes Upwork jobs and generates winning proposals
- **Includes**: Upwork scraping, proposal generation, PandaDoc integration
- **API Keys Needed**: APIFY_API_TOKEN, OPENAI_API_KEY, PANDADOC_API_KEY

### video_editing
- **Purpose**: Automates video editing with voice activity detection
- **Includes**: Jump cut automation, transition effects
- **API Keys Needed**: None (local processing)

### crm_integration
- **Purpose**: Manages data flow between sheets, webhooks, and cloud services
- **Includes**: Google Sheets CRUD, Modal webhooks, Slack notifications
- **API Keys Needed**: GOOGLE_SHEETS_CREDENTIALS_FILE, MODAL_TOKEN_ID, SLACK_WEBHOOK_URL

### full_stack
- **Purpose**: Complete agent with ALL capabilities
- **Includes**: Everything from all other types
- **API Keys Needed**: All of the above

### custom
- **Purpose**: Minimal base setup for custom workflows
- **Includes**: Basic Google Sheets integration, OpenAI/Anthropic
- **API Keys Needed**: OPENAI_API_KEY, GOOGLE_SHEETS_CREDENTIALS_FILE

## Execution

### Step 1: Determine Agent Type
Ask the user what kind of agent they want to create. If unclear, ask clarifying questions:
- "What tasks do you want this agent to perform?"
- "Do you need lead generation, email automation, video editing, or something custom?"

### Step 2: Get Agent Name
Ask for a name for the workspace. This will be used for the folder name.

### Step 3: Run the Generator Script

```bash
python execution/create_agent_workspace.py --name "{agent_name}" --type {agent_type}
```

With additional items:
```bash
python execution/create_agent_workspace.py \
  --name "{agent_name}" \
  --type {agent_type} \
  --additional-scripts "script1.py,script2.py" \
  --additional-directives "directive1.md" \
  --additional-packages "some-package"
```

### Step 4: Verify Output
The workspace will be created at `outputs/{agent-name}/` with:
- `AGENTS.md` - Customized system prompt
- `README.md` - Getting started guide
- `.env.example` - Required environment variables
- `requirements.txt` - Python dependencies
- `.github/agents/{name}.agent.md` - **VS Code custom agent (auto-generated)**
- `directives/` - Relevant SOPs
- `execution/` - Relevant scripts
- `.gitignore` - Standard ignores

### Step 5: Guide User on Next Steps
Tell the user:
1. The workspace location
2. To copy it to their desired location
3. To set up their `.env` file
4. To install dependencies with `pip install -r requirements.txt`
5. To open in VS Code — **the custom agent is already configured and ready to use**
6. To select their agent from the Copilot Chat agent dropdown (no manual setup needed)

## Output
- Complete agent workspace in `outputs/{agent-name}/`
- Ready-to-use VS Code custom agent in `.github/agents/`
- Ready to be copied and used independently

## Edge Cases

### User wants multiple agent types combined
Use `full_stack` type, or use a base type with `--additional-scripts` and `--additional-directives`.

### Workspace name already exists
The script will append a timestamp to create a unique folder.

### User doesn't know what type to choose
Walk them through the available types and their capabilities. If still unclear, suggest `custom` type and help them add specific scripts as needed.

### User needs a capability not covered
1. Create with `custom` type
2. Help them create new directives and scripts
3. Consider contributing back to the main framework

## Example Conversations

**User**: "I want to create a lead generation agent"
**Agent**: Runs `python execution/create_agent_workspace.py --name "Lead Gen Agent" --type lead_generation`

**User**: "Set up an agent for Upwork automation that also does email follow-ups"
**Agent**: Runs `python execution/create_agent_workspace.py --name "Upwork Outreach Agent" --type freelance_proposals --additional-scripts "instantly_autoreply.py,instantly_create_campaigns.py" --additional-directives "instantly_autoreply.md"`

**User**: "Create a new agent workspace"
**Agent**: "What kind of tasks do you want this agent to handle? I can create workspaces for: lead generation, email automation, freelance proposals, video editing, CRM integration, or a custom setup."
