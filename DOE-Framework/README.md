# The Complete Guide to Building Agentic Workflows

A comprehensive, step-by-step tutorial on building and deploying AI-powered agentic workflows for business automation.

> **This workspace is pre-configured with working examples.** Reference the live files in this repository to see real implementations of the concepts covered.

---

## Table of Contents

1. [Quick Start: Generate an Agent Workspace](#quick-start-generate-an-agent-workspace)
2. [Introduction to Agentic Workflows](#introduction-to-agentic-workflows)
3. [Understanding the AI Overhang](#understanding-the-ai-overhang)
4. [Documents, Chatbots, and Agents](#documents-chatbots-and-agents)
5. [The Five Components of an Agent](#the-five-components-of-an-agent)
6. [Why Frameworks Are Necessary](#why-frameworks-are-necessary)
7. [The DOE Framework](#the-doe-framework-directive-orchestration-execution)
8. [Setting Up VS Code + GitHub Copilot](#setting-up-vs-code--github-copilot)
9. [Workspace Structure & Key Files](#workspace-structure--key-files)
10. [Creating Your First Agentic Workflow](#creating-your-first-agentic-workflow)
11. [Writing Effective Directives](#writing-effective-directives)
12. [Building Execution Scripts](#building-execution-scripts)
13. [Claude Skills Framework](#claude-skills-framework)
14. [Model Context Protocol (MCP)](#model-context-protocol-mcp)
15. [Self-Annealing: Building Resilient Workflows](#self-annealing-building-resilient-workflows)
16. [Using Workflows in Practice](#using-workflows-in-practice)
17. [Meta-Directives: Chaining Workflows](#meta-directives-chaining-workflows)
18. [API Documentation Perusal](#api-documentation-perusal)
19. [Building CRM Wrappers](#building-crm-wrappers)
20. [Deploying to the Cloud](#deploying-to-the-cloud)
21. [Running Multiple Agents in Parallel](#running-multiple-agents-in-parallel)
22. [Sub-Agents](#sub-agents)
23. [Best Practices and Safety Guidelines](#best-practices-and-safety-guidelines)
24. [Quick Reference: This Workspace](#quick-reference-this-workspace)

---

## Quick Start: Generate an Agent Workspace

**This project can automatically generate standalone agent workspaces** that you can copy to separate folders for independent development.

### Prerequisites

Before generating workspaces, ensure you have:

1. **Python 3.10+** installed ([python.org](https://www.python.org/downloads/))
2. **VS Code** with **GitHub Copilot** extension (see [Setting Up VS Code + GitHub Copilot](#setting-up-vs-code--github-copilot))
3. **This repository cloned**:
   ```powershell
   git clone https://github.com/vjvarada/DOE-Framework-Agentic-AI.git
   cd DOE-Framework-Agentic-AI
   ```

> **Note**: The generator script itself has no dependencies. For the generated workspaces, you'll install their specific requirements after creation.

### Ask the Agent

Simply tell your AI agent:

> "Create a new lead generation agent called 'My Lead Bot'"

Or:

> "Set up an agent workspace for email automation"

The agent will use [directives/create_agent_workspace.md](directives/create_agent_workspace.md) to scaffold a complete workspace.

### Run Manually

```powershell
# List available agent types
python execution/create_agent_workspace.py --list-types

# Create a workspace
python execution/create_agent_workspace.py --name "My Agent" --type lead_generation
```

### Available Agent Types

| Type | Description | Key Capabilities |
|------|-------------|------------------|
| `lead_generation` | Scrapes and enriches leads | Google Maps, SERP, Apify, email enrichment |
| `email_automation` | Cold email campaigns | Instantly.ai integration, auto-replies |
| `freelance_proposals` | Job scraping & proposals | Upwork automation, AI proposal generation |
| `video_editing` | Video processing | Jump cuts via VAD, transitions |
| `crm_integration` | Data & webhook management | Google Sheets, Modal webhooks |
| `full_stack` | All capabilities combined | Everything above |
| `custom` | Minimal base setup | Basic structure, add your own |

### Generated Workspace Structure

```
outputs/my-agent/
├── AGENTS.md                    # Customized system prompt
├── README.md                    # Getting started guide
├── .env.example                 # Required API keys template
├── requirements.txt             # Python dependencies
├── .gitignore                   # Standard ignores
├── .github/
│   └── agents/
│       └── my-agent.agent.md    # ✨ VS Code custom agent (auto-generated)
├── directives/                  # Relevant SOPs only
│   └── *.md
└── execution/                   # Relevant scripts only
    └── *.py
```

> **✨ New:** The generator now creates a ready-to-use VS Code custom agent file (`.agent.md`) automatically. When you open the workspace in VS Code, the agent appears in the Copilot Chat agent dropdown immediately—no manual setup required!

### Combining Capabilities

Need email automation with a lead gen agent? Add extra scripts:

```powershell
python execution/create_agent_workspace.py \
  --name "Outreach Agent" \
  --type lead_generation \
  --additional-scripts "instantly_autoreply.py,instantly_create_campaigns.py" \
  --additional-directives "instantly_autoreply.md"
```

### After Generation

1. **Copy** the workspace from `outputs/` to your desired location
2. **Open** the new workspace folder in VS Code
3. **Select your agent** — The custom agent is already configured!
   - Open Copilot Chat Panel (`Ctrl+Shift+I`)
   - Click the **agent dropdown** at the top of the chat panel
   - Select your generated agent (e.g., "My Lead Bot")
   - Start chatting with your specialized agent immediately
4. **Set up environment** — Copy and configure API keys:
   ```powershell
   cp .env.example .env
   # Edit .env with your API keys
   ```
5. **Install** dependencies: `pip install -r requirements.txt`
6. **Start** using your specialized agent!

> **✨ Zero-Config Setup:** The generator creates a `.github/agents/*.agent.md` file that VS Code automatically detects. No manual agent creation required!

---

## Introduction to Agentic Workflows

Agentic workflows represent a fundamental shift in how we interact with AI systems. Rather than simply asking AI to answer questions or generate text (copy-paste interactions), agentic workflows allow AI to take **autonomous actions** in the real world—sending emails, managing databases, scraping data, creating documents, and orchestrating complex multi-step processes.

### What Makes Agentic Workflows Different

| Traditional AI Use | Agentic Workflows |
|-------------------|-------------------|
| One-time responses | Continuous execution loops |
| Manual copy-paste | Automated actions |
| Single-step tasks | Multi-step orchestration |
| Human-dependent | Semi-autonomous operation |
| Chat-based interface | Tool-based execution |

### The Business Case

Agentic workflows don't automate 100% of one role—they automate **90% of 10,000 roles**. This "horizontal leverage" is far more valuable:
- Automating 100% of one role = 1 unit of economic value
- Automating 90% of 10,000 roles = 9,000 units of economic value

---

## Understanding the AI Overhang

AI is currently in an **overhang state**—current capabilities far exceed what most people believe, expect, or know how to use.

### Three Reasons This Is Possible Now

**1. Intelligence Threshold Crossed**
- Frontier models (Claude, GPT, Gemini) score ~80% on SWE-bench verified
- This measures real software engineering ability on novel problems
- Models are now professional-grade in many domains

**2. Standardized Tool Integration**
- Model Context Protocol (MCP) standardizes AI-to-tool connections
- Frameworks like DOE formalize tool calling
- Cloud Skills and similar frameworks provide reliability

**3. Cost Economics**
- Token costs have plunged ~40x in the last year
- Claude Opus dropped from $15-75 to $5-25 per million tokens
- High-volume agentic workflows are now economically viable

---

## Documents, Chatbots, and Agents

Knowledge tools have evolved through three stages:

### 1. Documents (Static Knowledge)
- One-way information flow
- You read, but cannot interact
- Great for permanence (contracts, SOPs)
- Fixed once written

### 2. Chatbots (Dynamic Knowledge)
- Two-way interaction
- Ask questions, get responses
- Can clarify and dig deeper
- **Limitation**: Cannot take action in the world

### 3. Agents (Dynamic Action)
- Two-way interaction PLUS action
- Execute tasks, call APIs, modify systems
- Long execution times (minutes to hours)
- Autonomous decision-making

> **Key Insight**: An agent is not a chatbot. The chat interface is just a shell—agents use it temporarily until better interfaces emerge.

---

## The Five Components of an Agent

Every agent follows the **PTMRO** loop:

### 1. Planning
Breaking down high-level objectives into executable steps.

```
High-level task: "Eat at White Castle"
↓
Step 1: Get in the car
Step 2: Research GPS location
Step 3: Drive there
Step 4: Order food
Step 5: Enjoy meal
```

**Critical**: Planning errors at the start compound catastrophically. If you're off by 1% at the beginning, you could end up thousands of kilometers from your goal.

### 2. Tools
Tools are how agents interact with the real world:
- Calling APIs
- Executing code
- Database operations
- Web browsing
- File manipulation

> Think of tools like a caveman's spear—pre-built solutions for recurring problems.

### 3. Memory

| Memory Type | Description | Example |
|-------------|-------------|---------|
| Short-term | Reasoning tokens, discarded after use | Thinking/reasoning traces |
| Intermediate | Messages in current conversation | Chat history |
| Long-term | Persists across sessions | Files, system prompts, databases |

### 4. Reflection
The agent evaluates its own work:
- Examines outputs for errors
- Identifies failing approaches
- Knows when to pivot
- Self-corrects

### 5. Orchestration
The coordination layer that:
- Shuttles information between steps
- Takes planning results → feeds to tools
- Stores necessary data in memory
- Uses reflection results to adjust next planning cycle

---

## Why Frameworks Are Necessary

### The Stochasticity Problem

LLMs are **probabilistic, not deterministic**. They don't predict the single best next word—they predict a statistical distribution of options.

```
Input: "Hi, how are ___"
↓
Distribution: {you: 98%, things: 1%, your: 0.5%, ...}
↓
Selection uses temperature/top-p for randomness
```

### Compound Probability Disaster

If each step is 90% successful:

| Steps | Success Rate |
|-------|-------------|
| 1 | 90% |
| 5 | 59% |
| 10 | 35% |
| 20 | 12% |

**Example**: If you send wrong invoices 5% of the time, that doesn't cause 5% business impact—it causes ~95% impact because clients won't trust you.

### The Solution: Frameworks

We wrap the probabilistic AI in a deterministic framework:
- Give defined nodes between important steps
- Shorten the gaps where errors can accumulate
- Reserve AI for judgment, not mechanical operations
- Use code for deterministic tasks

---

## The DOE Framework (Directive, Orchestration, Execution)

DOE is a three-layer software architecture that constrains AI outputs for business reliability.

### Layer 1: Directives (The "What")
- SOPs written in natural language as Markdown files
- Define goals, inputs, tools, expected outputs, edge cases
- **No code**—readable by anyone in the organization
- Stored in `/directives` folder

```markdown
# lead_scraping.md

## Objective
Scrape B2B leads matching specified criteria

## Inputs
- Industry keyword
- Location
- Number of leads needed

## Process
1. Search LinkedIn Sales Navigator
2. Extract matching profiles
3. Verify 80% match rate
4. Enrich emails via secondary service
5. Export to Google Sheet

## Definition of Done
- Google Sheet URL with specified number of rows
- Each row has verified email address
```

### Layer 2: Orchestration (The "Who")
The AI agent that:
- Reads directives
- Makes routing decisions
- Coordinates tools
- Manages task flow
- Adapts to unexpected situations

### Layer 3: Execution (The "How")
Deterministic Python scripts that:
- Handle one specific task each
- Do the same thing every time given same inputs
- Don't hallucinate or guess
- Either work correctly or throw clear errors
- Stored in `/execution` folder

```
/execution
  ├── scrape_apollo.py
  ├── enrich_clearbit.py
  ├── send_email.py
  └── create_pandadoc.py
```

### Why DOE Works

| Component | Characteristic | Strength |
|-----------|---------------|----------|
| Directives | Natural language, flexible | Easy to understand/modify |
| Orchestration | AI-powered, adaptive | Handles ambiguity |
| Execution | Code, deterministic | Reliable, fast, precise |

**Key Insight**: Reserve LLM calls for judgment. Let code handle mechanical operations—it's 10,000-100,000x faster and essentially free.

---

## Setting Up VS Code + GitHub Copilot

This section provides a detailed step-by-step guide for setting up your agentic workflow environment using Visual Studio Code and GitHub Copilot.

### Prerequisites

Before starting, ensure you have:
- [ ] Visual Studio Code installed ([download here](https://code.visualstudio.com/))
- [ ] A GitHub account with Copilot subscription (Individual, Business, or Enterprise)
- [ ] Python 3.10+ installed
- [ ] Node.js 18+ installed (for Trigger.dev cloud functions)
- [ ] Git installed and configured

### Step 1: Install Required VS Code Extensions

Open VS Code and install these extensions (Ctrl+Shift+X):

| Extension | Purpose | Extension ID |
|-----------|---------|--------------|
| **GitHub Copilot** | AI code completion | `GitHub.copilot` |
| **GitHub Copilot Chat** | Conversational AI interface | `GitHub.copilot-chat` |
| **Python** | Python language support | `ms-python.python` |
| **Pylance** | Python IntelliSense | `ms-python.vscode-pylance` |

**Quick Install Command** (run in VS Code terminal):
```powershell
code --install-extension GitHub.copilot
code --install-extension GitHub.copilot-chat
code --install-extension ms-python.python
code --install-extension ms-python.vscode-pylance
```

### Step 2: Configure GitHub Copilot (Manual Steps Required)

> ⚠️ **Important**: The initial Copilot setup requires manual UI interaction—it cannot be fully automated.

**Step 2a: Sign In and Verify Copilot is Active**

1. **Click the Copilot icon** in the bottom-right status bar (or Activity Bar)
2. **Sign in** with your GitHub account if prompted
3. **Verify activation**: You should see "Copilot" status as active (not "Copilot: Inactive")
4. **Check subscription**: Ensure your GitHub account has an active Copilot subscription

> **Troubleshooting**: If Copilot isn't activating, go to **File → Preferences → Settings** (`Ctrl+,`), search for "Copilot", and ensure it's enabled.

**Step 2b: Understanding Copilot Customization Options**

VS Code offers multiple ways to customize GitHub Copilot's behavior:

| Customization Type | File Location | Purpose |
|--------------------|---------------|---------|
| **Custom Instructions** | `.github/copilot-instructions.md` | Project-wide coding standards and guidelines |
| **Agent Instructions** | `AGENTS.md` (root) | System prompt for AI agent behavior |
| **Custom Agents** | `.github/agents/*.agent.md` | Specialized AI personas with specific tools |

This workspace uses [AGENTS.md](AGENTS.md) as the agent instruction file — Copilot automatically reads it when working in this workspace.

**Step 2c: Create a Custom Agent (Optional but Recommended)**

Custom agents let you create specialized AI personas with specific tools and instructions. To create one:

1. **Open Copilot Chat Panel**: Press `Ctrl+Shift+I`
2. **Click the agent dropdown** at the top of the chat panel (shows "Agent" or current agent name)
3. **Select "Configure Custom Agents"**
4. **Click "Create new custom agent"** (or run `Chat: New Custom Agent` from Command Palette)
5. **Choose storage location**:
   - **Workspace** (`.github/agents/`) — shared with your team via Git
   - **User Profile** — personal, available across all workspaces
6. **Enter a file name** (e.g., `doe-agent`)
7. **Configure the `.agent.md` file** with instructions and available tools

> **For this workspace**: You can copy the contents of [AGENTS.md](AGENTS.md) into a custom agent file to get the DOE framework behavior as a switchable agent.

**Example Custom Agent File** (`.github/agents/doe-agent.agent.md`):
```yaml
---
description: DOE Framework agent for directive-based automation
name: DOE Agent
tools: ['search', 'fetch', 'terminal', 'editFiles']
---
# DOE Framework Agent
You operate within a 3-layer architecture (Directive, Orchestration, Execution).
Read the full instructions from AGENTS.md in the workspace root.
```

**Step 2d: Open Copilot Chat Panel**
   - Press `Ctrl+Shift+I` to open Copilot Chat
   - Or click the Copilot icon in the Activity Bar
   - Switch agents using the dropdown at the top of the chat panel

### Step 3: Open This Workspace

```powershell
# Clone or navigate to this repository
cd "C:\Users\YourUsername\Documents\Github"
git clone <repository-url> "DOE Framework Agentic AI"
code "DOE Framework Agentic AI"
```

### Step 4: Set Up Python Environment

This workspace includes a [requirements.txt](requirements.txt) with all necessary Python dependencies:

```powershell
# Create virtual environment
python -m venv .venv

# Activate it (PowerShell)
.\.venv\Scripts\Activate.ps1

# Install all dependencies from requirements.txt
pip install -r requirements.txt
```

**Key dependencies in [requirements.txt](requirements.txt):**

| Category | Packages |
|----------|----------|
| **AI/LLM** | `anthropic`, `openai`, `cohere`, `google-genai` |
| **Google Services** | `google-auth`, `google-api-python-client`, `gspread` |
| **Web Scraping** | `apify-client`, `beautifulsoup4`, `playwright`, `httpx` |
| **Data Processing** | `pandas`, `numpy`, `Pillow` |
| **Computer Vision** | `mediapipe`, `opencv-python` (for thumbnail face matching) |
| **Infrastructure** | `modal`, `python-dotenv` |

### Step 4b: Set Up Node.js Environment (Optional)

For Trigger.dev cloud functions, this workspace also includes [package.json](package.json):

```powershell
# Install Node.js dependencies
npm install

# Run Trigger.dev development server
npm run dev

# Deploy to Trigger.dev cloud
npm run deploy
```

**Key dependencies in [package.json](package.json):**
- `@anthropic-ai/sdk` - Claude API for Node.js
- `@trigger.dev/sdk` - Cloud function orchestration (alternative to Modal)
- `googleapis` - Google APIs for Node.js
- `zod` - Runtime type validation

### Step 5: Configure Environment Variables

Create a `.env` file in the workspace root:

```env
# AI API Keys
ANTHROPIC_API_KEY=your_anthropic_key_here
OPENAI_API_KEY=your_openai_key_here

# Data APIs
APIFY_API_TOKEN=your_apify_token_here
ANYMAILFINDER_API_KEY=your_anymailfinder_key_here

# Service APIs
PANDADOC_API_KEY=your_pandadoc_key_here
INSTANTLY_API_KEY=your_instantly_key_here

# Notifications
SLACK_WEBHOOK_URL=your_slack_webhook_here
```

> **Security**: Never commit `.env` to git. Add it to `.gitignore`.

### Step 6: Verify Agent System Prompt

The agent reads [AGENTS.md](AGENTS.md) at the start of every conversation. This file teaches the agent:
- The DOE framework architecture
- How to read directives and call execution scripts
- Self-annealing behavior for error handling
- File organization conventions

**Verify the system prompt is loaded**:
1. Open Copilot Chat (Ctrl+Shift+I)
2. Type: "What framework are you operating under?"
3. The agent should reference the 3-layer DOE architecture

### Step 7: Test Your Setup

Run this test workflow in Copilot Chat:

```
Set up my workspace in accordance with AGENTS.md. 
Verify that the directives/ and execution/ folders exist 
and list what workflows are available.
```

Expected response: Agent should list the directives and confirm the workspace structure.

---

## Workspace Structure & Key Files

This workspace is already configured with working examples. Here's what each file does:

### System Configuration Files

| File | Purpose | View It |
|------|---------|---------|
| [AGENTS.md](AGENTS.md) | Universal system prompt for all AI agents | Main agent instructions |
| [requirements.txt](requirements.txt) | Python dependencies | `pip install -r requirements.txt` |
| [package.json](package.json) | Node.js dependencies (Trigger.dev) | `npm install` |
| `.env` | API keys and secrets (create this yourself) | Not committed to git |

### Directives (The "What")

Location: [`directives/`](directives/)

| Directive | Purpose |
|-----------|---------|
| [create_agent_workspace.md](directives/create_agent_workspace.md) | **Generate new agent workspaces** |
| [scrape_leads.md](directives/scrape_leads.md) | Lead scraping with industry verification |
| [create_proposal.md](directives/create_proposal.md) | PandaDoc proposal generation |
| [gmaps_lead_generation.md](directives/gmaps_lead_generation.md) | Google Maps lead enrichment pipeline |
| [google_serp_lead_scraper.md](directives/google_serp_lead_scraper.md) | Google search results scraping |
| [instantly_autoreply.md](directives/instantly_autoreply.md) | Cold email auto-reply automation |
| [upwork_scrape_apply.md](directives/upwork_scrape_apply.md) | Upwork job scraping and application |
| [jump_cut_vad.md](directives/jump_cut_vad.md) | Video editing automation |

### Execution Scripts (The "How")

Location: [`execution/`](execution/)

**Lead Generation Scripts:**
| Script | Function |
|--------|----------|
| [scrape_apify.py](execution/scrape_apify.py) | Single Apify lead scrape |
| [scrape_apify_parallel.py](execution/scrape_apify_parallel.py) | Parallel lead scraping (3-4x faster) |
| [scrape_google_maps.py](execution/scrape_google_maps.py) | Google Maps business scraping |
| [gmaps_lead_pipeline.py](execution/gmaps_lead_pipeline.py) | Full Google Maps enrichment pipeline |
| [gmaps_parallel_pipeline.py](execution/gmaps_parallel_pipeline.py) | Parallel Google Maps processing |
| [enrich_emails.py](execution/enrich_emails.py) | Email enrichment via AnyMailFinder |

**Google Sheets Scripts:**
| Script | Function |
|--------|----------|
| [read_sheet.py](execution/read_sheet.py) | Read data from Google Sheets |
| [update_sheet.py](execution/update_sheet.py) | Write/update Google Sheets |
| [append_to_sheet.py](execution/append_to_sheet.py) | Append rows to sheets |

**Proposal & Email Scripts:**
| Script | Function |
|--------|----------|
| [create_proposal.py](execution/create_proposal.py) | Generate PandaDoc proposals |
| [welcome_client_emails.py](execution/welcome_client_emails.py) | Send onboarding emails |
| [instantly_autoreply.py](execution/instantly_autoreply.py) | Process Instantly webhook replies |
| [instantly_create_campaigns.py](execution/instantly_create_campaigns.py) | Create cold email campaigns |

**Text Processing Scripts:**
| Script | Function |
|--------|----------|
| [casualize_batch.py](execution/casualize_batch.py) | Batch text casualization |
| [casualize_first_names_batch.py](execution/casualize_first_names_batch.py) | First name formatting |
| [casualize_company_names_batch.py](execution/casualize_company_names_batch.py) | Company name formatting |
| [casualize_city_names_batch.py](execution/casualize_city_names_batch.py) | City name formatting |

**Upwork Automation:**
| Script | Function |
|--------|----------|
| [upwork_scraper.py](execution/upwork_scraper.py) | Scrape Upwork job listings |
| [upwork_apify_scraper.py](execution/upwork_apify_scraper.py) | Apify-based Upwork scraping |
| [upwork_proposal_generator.py](execution/upwork_proposal_generator.py) | Generate Upwork proposals |

**Cloud & Webhooks:**
| Script | Function |
|--------|----------|
| [modal_webhook.py](execution/modal_webhook.py) | Modal cloud webhook server |
| [webhooks.json](execution/webhooks.json) | Webhook → directive mappings |
| [onboarding_post_kickoff.py](execution/onboarding_post_kickoff.py) | Post-kickoff automation |

**Video Processing:**
| Script | Function |
|--------|----------|
| [jump_cut_vad_singlepass.py](execution/jump_cut_vad_singlepass.py) | Voice activity detection for editing |
| [insert_3d_transition.py](execution/insert_3d_transition.py) | 3D transition insertion |

**Workspace Generator:**
| Script | Function |
|--------|----------|
| [create_agent_workspace.py](execution/create_agent_workspace.py) | **Generate standalone agent workspaces** |
| [agent_templates.json](execution/agent_templates.json) | Agent type configurations |

---

## Creating Your First Agentic Workflow

Now that your environment is set up, let's walk through creating and running your first workflow using the existing files in this workspace.

### Understanding the Flow

```
┌─────────────────────────────────────────────────────────────────┐
│                    YOUR NATURAL LANGUAGE REQUEST                │
│              "Scrape 100 plumbers in Austin, Texas"             │
└─────────────────────────────────────────────────────────────────┘
                                ↓
┌─────────────────────────────────────────────────────────────────┐
│                  GITHUB COPILOT (Orchestrator)                  │
│  1. Reads AGENTS.md to understand its role                      │
│  2. Finds directives/scrape_leads.md for instructions           │
│  3. Executes execution/scrape_apify.py with parameters          │
│  4. Handles errors, retries, updates directives                 │
└─────────────────────────────────────────────────────────────────┘
                                ↓
┌─────────────────────────────────────────────────────────────────┐
│                         OUTPUT                                   │
│              Google Sheet URL with enriched leads               │
└─────────────────────────────────────────────────────────────────┘
```

### Step-by-Step: Running the Lead Scraper

**Step 1: Open Copilot Chat**
- Press `Ctrl+Shift+I` or click the Copilot Chat icon
- The agent automatically reads [AGENTS.md](AGENTS.md) for context

**Step 2: Issue Your Command**
```
Scrape 100 plumbers in Austin, Texas. 
Give me the Google Sheet URL when done.
```

**Step 3: Agent Reads the Directive**

The agent will find and read [directives/scrape_leads.md](directives/scrape_leads.md), which specifies:
- Use [execution/scrape_apify.py](execution/scrape_apify.py) for small scrapes (<1000 leads)
- Run a test scrape of 25 leads first
- Verify 80%+ industry match
- Enrich emails using [execution/enrich_emails.py](execution/enrich_emails.py)
- Upload to Google Sheets using [execution/update_sheet.py](execution/update_sheet.py)

**Step 4: Watch the Execution**

The agent will execute commands like:
```powershell
# Test scrape (25 leads)
python execution/scrape_apify.py --query "Plumbers" --location "Austin, Texas" --max_items 25 --no-email-filter

# Verify results in .tmp/test_leads.json

# Full scrape (100 leads)
python execution/scrape_apify.py --query "Plumbers" --location "Austin, Texas" --max_items 100 --no-email-filter

# Enrich emails
python execution/enrich_emails.py --input .tmp/leads_[timestamp].json

# Upload to sheet
python execution/update_sheet.py --input .tmp/leads_[timestamp].json
```

**Step 5: Receive Your Deliverable**

The agent returns a Google Sheet URL with your enriched leads.

### Example: Creating a Proposal

**Your Command:**
```
Create a proposal for John Smith at ACME Corp (john@acme.com).
Project: AI Automation Implementation
Value: $15,000 over 3 months
Problems: Manual data entry, slow lead response, inconsistent follow-up, no analytics
```

**Agent Reads:** [directives/create_proposal.md](directives/create_proposal.md)

**Agent Executes:** [execution/create_proposal.py](execution/create_proposal.py)

**Output:** PandaDoc proposal link sent to client's email

### Example: Google Maps Lead Generation

**Your Command:**
```
Find 50 HVAC contractors in Miami, FL using Google Maps.
I want deep contact enrichment including owner info.
```

**Agent Reads:** [directives/gmaps_lead_generation.md](directives/gmaps_lead_generation.md)

**Agent Executes:**
1. [execution/gmaps_lead_pipeline.py](execution/gmaps_lead_pipeline.py) - Scrapes Google Maps
2. Enriches each business with website scraping + DuckDuckGo search
3. Uses Claude to extract structured contact data
4. Uploads to Google Sheets

**Output:** Google Sheet with 36 fields per lead including owner name, email, phone, LinkedIn

---

## Writing Effective Directives

Directives are the instruction manuals that tell the agent **what** to do. This workspace contains several production-ready examples.

### Anatomy of a Directive

Study [directives/scrape_leads.md](directives/scrape_leads.md) as a reference implementation:

```markdown
# Lead Scraping & Verification

## Goal
Scrape leads using Apify, verify relevance (industry match > 80%), 
and save them to a Google Sheet.

## Inputs
- **Industry**: Target industry (e.g., "Plumbers")
- **Location**: Target location (e.g., "New York")
- **Total Count**: Number of leads desired

## Tools/Scripts
- Script: `execution/scrape_apify.py` (for <1000 leads)
- Script: `execution/scrape_apify_parallel.py` (for 1000+ leads)
- Script: `execution/update_sheet.py`

## Process
1. Test Scrape - Run with max_items=25
2. Verification - Check 80%+ match rate
3. Full Scrape - Run with full count
4. Upload to Google Sheet (DELIVERABLE)
5. Enrich Missing Emails

## Edge Cases
- If Vain rate limited, wait 60 seconds and retry
- If <50% leads have emails, notify user
```

### Four Essential Components

| Component | What It Does | Example from This Workspace |
|-----------|-------------|----------------------------|
| **Goal** | Clear objective statement | "Scrape B2B leads and save to Google Sheet" |
| **Inputs** | Required data to start | Industry, Location, Total Count |
| **Tools/Scripts** | Which execution scripts to use | `execution/scrape_apify.py` |
| **Process** | Step-by-step sequence | Test → Verify → Full Scrape → Upload |

### Additional Elements (Recommended)

- **Edge Cases**: API quirks, rate limits, error handling
- **Fallback Behavior**: What to do if primary approach fails
- **Definition of Done**: Quality criteria for completion

### Real Examples in This Workspace

| Directive | Key Features to Study |
|-----------|----------------------|
| [scrape_leads.md](directives/scrape_leads.md) | Parallel processing, verification gates, email enrichment |
| [create_proposal.md](directives/create_proposal.md) | YAML front matter, multi-step content generation |
| [gmaps_lead_generation.md](directives/gmaps_lead_generation.md) | Deep enrichment pipeline, 36-field output schema |
| [instantly_autoreply.md](directives/instantly_autoreply.md) | Webhook-triggered automation, response classification |

### Creating a New Directive

**Option 1: From Scratch via Copilot Chat**
```
Create a new directive for sending weekly report emails.
The directive should:
- Read data from a specified Google Sheet
- Generate a summary using Claude
- Send formatted email to a list of recipients
- Run every Monday at 9 AM
```

**Option 2: Convert Existing SOPs**

If you already have SOPs (Standard Operating Procedures) for your business, you're **halfway there**:

1. Drag your SOP document into the workspace
2. Tell the agent:
   ```
   I just uploaded a file into the workspace. Turn it into a 
   directive and build the execution scripts to make it happen.
   ```
3. The agent will:
   - Create the directive in `/directives`
   - Build necessary scripts in `/execution`
   - Ask for any missing API tokens

**Why this works**: SOPs are literally already directives—they contain goals, steps, inputs, outputs. You're just reformatting into a more token-efficient structure.

**Bonus**: If your SOPs are incomplete, the agent will ask clarifying questions, forcing you to resolve ambiguities. The resulting directive ends up **better than the original SOP**.

**Option 3: From Voice Dump**

No documentation? No problem:
```
[Voice transcribe your stream of consciousness]

"I basically want to take new leads from our website form, 
check if they're qualified based on company size, if they are 
send them a welcome email with our calendar link, if not add 
them to a nurture sequence, and log everything in our CRM"
```

The agent converts your casual explanation into a structured directive.

The agent will:
1. Create `directives/weekly_report.md`
2. Define inputs, process, and edge cases
3. Reference or create necessary execution scripts

---

## Building Execution Scripts

Execution scripts are the deterministic Python code that does the actual work. This workspace has 20+ production scripts you can study and reuse.

### Key Principles

1. **One Job Per Script** - Each script handles one specific task
2. **Deterministic Behavior** - Same inputs → Same outputs
3. **Let AI Build Them** - Describe what you want; the agent writes the code

### Script Categories in This Workspace

#### Data Scraping Scripts

| Script | What It Does | Usage |
|--------|-------------|-------|
| [scrape_apify.py](execution/scrape_apify.py) | Scrape leads via Apify | `python scrape_apify.py --query "Plumbers" --location "NYC" --max_items 100` |
| [scrape_apify_parallel.py](execution/scrape_apify_parallel.py) | Parallel scraping (3-4x faster) | `python scrape_apify_parallel.py --total_count 4000 --strategy regions` |
| [scrape_google_maps.py](execution/scrape_google_maps.py) | Google Maps business data | Used by gmaps_lead_pipeline.py |

**Study [scrape_apify.py](execution/scrape_apify.py)** to understand the pattern:
- Loads API keys from `.env` using `python-dotenv`
- Takes command-line arguments with `argparse`
- Returns structured JSON output
- Saves to `.tmp/` directory (intermediates)

#### Google Sheets Scripts

| Script | What It Does |
|--------|-------------|
| [read_sheet.py](execution/read_sheet.py) | Read data from any Google Sheet |
| [update_sheet.py](execution/update_sheet.py) | Write/update sheet data |
| [append_to_sheet.py](execution/append_to_sheet.py) | Add rows to existing sheet |

These scripts use Google Service Account authentication via `credentials.json`.

#### Email & Proposal Scripts

| Script | What It Does |
|--------|-------------|
| [create_proposal.py](execution/create_proposal.py) | Generate PandaDoc proposals |
| [welcome_client_emails.py](execution/welcome_client_emails.py) | Send welcome emails |
| [instantly_autoreply.py](execution/instantly_autoreply.py) | Process cold email replies |
| [instantly_create_campaigns.py](execution/instantly_create_campaigns.py) | Create email campaigns |

#### Text Processing Scripts

| Script | What It Does |
|--------|-------------|
| [casualize_batch.py](execution/casualize_batch.py) | Make text more conversational |
| [casualize_first_names_batch.py](execution/casualize_first_names_batch.py) | Format first names naturally |
| [casualize_company_names_batch.py](execution/casualize_company_names_batch.py) | Format company names |

### Creating a New Script

In Copilot Chat:
```
Create an execution script that takes a CSV of company names 
and returns their LinkedIn company page URLs. Save it as 
execution/linkedin_lookup.py.
```

The agent will:
1. Write the Python script
2. Include proper error handling
3. Use environment variables for API keys
4. Add argparse for command-line usage
5. Test the script

### Why Code Over AI for Execution?

| Task | LLM Time | Python Script | Improvement |
|------|----------|---------------|-------------|
| Sort 200 items | 30 seconds | 53 ms | 500x faster |
| Math operations | Minutes | Microseconds | 100,000x faster |
| Repeat task 1000x | Variable | Identical | Perfect consistency |

**Key Insight from [AGENTS.md](AGENTS.md)**: Reserve LLM calls for judgment. Let code handle mechanical operations.

---

## Claude Skills Framework

Claude Skills is Anthropic's approach to standardized agentic workflows—specific to Claude models.

### Structure

```
/skills
├── /generate-report
│   ├── skill.md
│   └── /scripts
│       └── report_generator.py
├── /big-query
│   ├── skill.md
│   └── data_sources.md
└── /nda-review
    ├── skill.md
    └── rules.md
```

### YAML Front Matter

Skills use metadata at the top of files:

```yaml
---
name: generate-report
description: Creates weekly weather reports from public APIs
allowed_tools:
  - web_fetch
  - file_write
---
```

This allows the agent to understand what a skill does **without reading the entire file**, saving tokens.

### Pre-Built Skills

Claude has optimized skills for common tasks:
- PDF creation
- Word documents
- Excel spreadsheets
- PowerPoint presentations

These are battle-tested across thousands of runs.

---

## Model Context Protocol (MCP)

MCP is the "USB-C for AI"—a universal adapter letting any AI connect to any data source.

### Components

**MCP Clients**: The AI apps (VS Code, Claude Desktop, Anti-gravity)

**MCP Servers**: Specific tool integrations (Apollo MCP, Google Drive MCP, Gmail MCP)

### How It Works

```
Your Agent (Client)
    ↓
Calls MCP Server
    ↓
Server executes action
    ↓
Returns results
```

### Communication Types

| Type | Description |
|------|-------------|
| Resources | Structured data (documents, database records) |
| Tools | Functions the agent can call |
| Prompts | Instructions for how to interact with the server |

### Setting Up an MCP

```
Hey, I want to set up a Gmail MCP so I can send emails 
from my email address.
```

The agent will:
1. Research available MCPs
2. Guide authentication
3. Configure the connection
4. Test functionality

### MCP Cautions

**Token Cost**: MCPs load many functions into context
- 5 servers × 10 tools × 300 tokens = 15,000 tokens before doing anything

**Quality Variance**: Not all MCPs are equal
- Some are rushed to market
- Test thoroughly before relying on them

**Context Pollution**: More tokens = worse performance
- Use MCP selectively
- Consider custom execution scripts instead

---

## Self-Annealing: Building Resilient Workflows

Self-annealing is how workflows strengthen themselves over time. This concept is embedded in [AGENTS.md](AGENTS.md) as a core operating principle.

### The Self-Annealing Loop (from AGENTS.md)

```
Error Occurs
    ↓
Diagnose Problem (read error message and stack trace)
    ↓
Fix the Script
    ↓
Test the Fix (unless it uses paid API credits - ask user first)
    ↓
Update the Directive (document what was learned)
    ↓
System is now stronger
```

### How It's Configured in This Workspace

From [AGENTS.md](AGENTS.md):

```markdown
## Operating Principles

**2. Self-anneal when things break**
- Read error message and stack trace
- Fix the script and test it again (unless it uses paid tokens/credits/etc—
  in which case you check w user first)
- Update the directive with what you learned (API limits, timing, edge cases)
- Example: you hit an API rate limit → you then look into API → find a batch 
  endpoint that would fix → rewrite script to accommodate → test → update directive.

**3. Update directives as you learn**
Directives are living documents. When you discover API constraints, better 
approaches, common errors, or timing expectations—update the directive.
```

### Real Example: Parallel Scraping Evolution

Look at [directives/scrape_leads.md](directives/scrape_leads.md). The parallel processing section exists because:
1. Sequential scraping was too slow for large datasets
2. Agent discovered geographic partitioning doesn't increase API costs
3. Agent created [execution/scrape_apify_parallel.py](execution/scrape_apify_parallel.py)
4. Directive was updated with new workflow option

### Employee A vs Employee B

| Employee A (Blocker) | Employee B (Self-Capable) |
|---------------------|--------------------------|
| Every problem becomes your problem | Tries solutions before escalating |
| Work halts waiting for you | Only escalates when necessary |
| Makes same mistakes repeatedly | Documents solutions for team |

**Self-annealing agents behave like Employee B.**

### Safety Guidelines for Self-Annealing

These are built into [AGENTS.md](AGENTS.md):

1. **Check before expensive operations**: "If it uses paid tokens/credits, check with user first"
2. **Never modify credentials**: API keys in `.env` are sacred
3. **Log all changes**: Changelogs at bottom of directives enable rollback
4. **Test before deploying**: Verify fixes work before updating production directives

---

## Using Workflows in Practice

### The Text Box Interface

Your primary interface is now a single text box. Everything happens through natural language commands.

### Maximize Bandwidth: Voice Transcription

This is a game-changer that most people underestimate:

| Method | Speed | Multiplier |
|--------|-------|------------|
| Typing | 50-70 WPM | 1x baseline |
| Speaking | 150-200 WPM | **3-4x faster input** |
| Reading | 300-500 WPM (skimming to 1000+) | **5-10x faster output** |

**Optimal Flow**: Voice transcription for input → Read output

**How to set this up in VS Code:**
1. Use Windows built-in voice typing: `Win+H` starts dictation
2. Or use tools like Whisper-based transcription (OpenAI)
3. Speak naturally—modern models handle casual speech well

**Pro tip**: You don't need formal syntax or precise technical language. Just speak like you're messaging a colleague:
- ❌ "Execute the lead scraping directive with parameters: industry=HVAC, location=Texas"
- ✅ "Hey, scrape me 200 HVAC companies in Texas, enrich the emails, give me the sheet"

### Running Workflows

Simply ask:
```
Scrape 200 HVAC companies in Texas, verify emails, 
personalize them, give me the Google sheet.
```

Be specific upfront to reduce back-and-forth:
- Specify quantities
- Name the target
- Describe desired output format

### When to Watch, When to Let Run

**Watch closely (first 10-15 runs)**:
- Develop intuition for model reasoning
- Catch errors early
- Identify optimization opportunities

**Let run autonomously**:
- Well-tested workflows
- Low-sensitivity tasks
- Use hooks/notifications to alert on completion

### Setting Up Notifications

```
Set up a hook that plays a chime sound every time 
an agent task completes. I usually have you alt-tabbed 
while doing other things.
```

### Claude Code Hooks (Step-by-Step)

Hooks execute shell commands in response to agent events—perfect for staying aware of long-running tasks.

**Setting up a completion chime:**

```
Hey, I'd like you to set me up a hook that plays a nice chime sound 
every time one of my agents is done with a task. That way, I'll know 
to go back to the task because I normally have you alt-tabbed while 
doing other things.
```

The agent will:
1. Research Claude Code hook syntax
2. Create a hooks configuration
3. Add a sound-playing script
4. Test the hook

**Different hooks for different events:**
- Task completion → Pleasant chime
- Error encountered → Alert tone  
- Human input required → Different sound

**Why this matters**: When juggling 3-4 agent instances, a significant portion of your time is wasted not knowing an agent finished and is waiting for you. Sound notifications solve this completely.

---

## Meta-Directives: Chaining Workflows

Once you have multiple individual workflows working, you can chain them into **meta-directives** (umbrella workflows).

### The Concept

Instead of manually triggering:
1. Lead scraping workflow
2. Email enrichment workflow
3. Personalization workflow
4. Campaign creation workflow

Create one **meta-directive** that chains them all:

```
New Lead Campaign (Meta-Directive)
├── scrape_leads.md      → Returns lead list
├── enrich_emails.md     → Adds email addresses  
├── personalize.md       → Creates first lines
└── create_campaign.md   → Uploads to Instantly
```

### Example Meta-Directive

Create `directives/new_campaign_endtoend.md`:

```markdown
# End-to-End New Campaign

## Goal
Take a single input (industry + location) and deliver a fully 
personalized cold email campaign ready to send.

## Process
1. **Scrape Leads**: Run scrape_leads.md with industry/location
2. **Enrich Emails**: Run enrich_emails on output
3. **Personalize**: Generate first lines using casualize_batch.py
4. **Create Campaign**: Upload to Instantly via instantly_create_campaigns.py
5. **Notify**: Send Slack message with campaign summary

## Inputs
- Industry keyword
- Location
- Number of leads (default: 200)

## Definition of Done
- Instantly campaign link
- Google Sheet with all leads
- Slack notification sent
```

### Running Meta-Directives

Simply ask:
```
Run the end-to-end new campaign workflow for HVAC contractors 
in Miami, FL. I want 200 leads.
```

The agent handles all the handoffs between sub-workflows automatically.

### Benefits

| Before (Manual Handoffs) | After (Meta-Directive) |
|-------------------------|----------------------|
| 4 separate commands | 1 command |
| You're the bottleneck | Agent handles transitions |
| Risk forgetting steps | Complete pipeline guaranteed |
| 30+ min of attention | Issue command, walk away |

---

## API Documentation Perusal

When building workflows that connect to new APIs, you'll often need to help your agent understand the API.

### The Challenge

Not all API documentation pages are AI-friendly:
- Some load content via JavaScript (invisible to web scrapers)
- Some require authentication to view full specs
- Some lack examples or detailed schemas

### Solution 1: Copy-Paste API Docs

For JavaScript-rendered docs, manually copy the content:

```
[Copy the API documentation page contents]

Here's the API documentation for the Vain lead scraper. 
Create an execution script that:
1. Authenticates with this API
2. Scrapes a LinkedIn Sales Navigator search URL
3. Returns structured lead data
```

### Solution 2: Look for "Copy for LLM" Buttons

Modern API providers (like Apify) now include AI-friendly features:

- **Copy for LLM** button → Exports markdown version
- **Open in ChatGPT/Claude** → Direct integration
- These are optimized for token efficiency

### Solution 3: Feed the OpenAPI Spec

Many APIs provide an `openapi.json` or Swagger spec:

```
Here's the OpenAPI specification for the XYZ API:
[paste JSON]

List all available endpoints and their parameters.
```

### Pro Tip: Agent as API Explorer

Don't just have the agent build—have it explore first:

```
Before building anything, research the Instantly API. 
Tell me:
1. What endpoints are available?
2. What authentication is required?
3. What are the rate limits?
4. Are there bulk/batch endpoints?

Then give me 3 approaches for building an email campaign creator.
```

---

## Building CRM Wrappers

One powerful pattern is creating an "agent wrapper" around your CRM (ClickUp, HubSpot, Salesforce, etc.) so you can manage it entirely through natural language.

### Two Approaches

**Approach 1: MCP Server**

```
I want to connect to my ClickUp CRM. Is there an MCP for this?
```

The agent will:
1. Search for official/community ClickUp MCP
2. Guide authentication setup
3. Configure the MCP connection
4. You can now say "Create a new lead named John Smith" and it just works

**Approach 2: Custom API Scripts**

```
Build a series of ClickUp directives so I can automate adding records, 
updating them, and so on. I want you to act as my ClickUp wrapper.
Do this via API calls, not MCP.
```

The agent will:
1. Create `execution/clickup_client.py` (base API client)
2. Create CRUD scripts (create, read, update, delete)
3. Create directives for each operation
4. Update `.env` with ClickUp API key

### Trade-offs

| MCP Approach | Custom API Approach |
|--------------|-------------------|
| Quick setup | More control |
| Loads tokens into context | Leaner execution |
| May lack custom fields | Full API access |
| Community-maintained | You maintain it |

### Example Commands After Setup

```
# Add a new lead
Add Peter Rockwell to my CRM in the Sales Pipeline

# Update status
Move John Smith to "Proposal Sent" status

# Query
Show me all leads in "Meeting Booked" status

# Bulk action
Send onboarding emails to all leads that closed this week
```

### Weaving CRM into Other Workflows

The real power is connecting CRM actions to other workflows:

```markdown
# In new_client_onboarding.md

## Process
1. Create client record in ClickUp
2. Update status to "Onboarding"
3. Run welcome_email workflow
4. Generate proposal using create_proposal.md
5. Add comment to ClickUp record: "Proposal sent at [timestamp]"
```

Now "Onboard new client John Smith" handles everything end-to-end.

---

## Deploying to the Cloud

This workspace supports two cloud deployment options:
- **Modal** (Python) - [execution/modal_webhook.py](execution/modal_webhook.py)
- **Trigger.dev** (Node.js) - Configured in [package.json](package.json)

### Choosing Your Platform

| Feature | Modal | Trigger.dev |
|---------|-------|-------------|
| **Language** | Python | TypeScript/JavaScript |
| **Setup** | `pip install modal` | `npm install @trigger.dev/sdk` |
| **Deploy** | `modal deploy` | `npm run deploy` |
| **Pricing** | Pay-per-use (very cheap) | Free tier + pay-per-use |
| **Best for** | Python scripts, ML workloads | TypeScript projects, complex workflows |

### Using Modal (Python)

This workspace includes [execution/modal_webhook.py](execution/modal_webhook.py) for Modal cloud deployment.

### Cloud Webhooks Architecture

From [AGENTS.md](AGENTS.md):

```markdown
## Cloud Webhooks (Modal)

The system supports event-driven execution via Modal webhooks. 
Each webhook maps to exactly one directive with scoped tool access.

**Key files:**
- `execution/webhooks.json` - Webhook slug → directive mapping
- `execution/modal_webhook.py` - Modal app (do not modify unless necessary)
```

### Current Webhook Configuration

See [execution/webhooks.json](execution/webhooks.json):

```json
{
  "webhooks": {
    "create-proposal": {
      "directive": "create_proposal_webhook",
      "description": "Generate PandaDoc proposal and send follow-up email",
      "tools": ["create_proposal", "send_email", "web_fetch"]
    }
  }
}
```

### Setting Up a New Webhook

**Step 1: Create the Directive**

Ask Copilot:
```
Create a webhook directive for automatically processing new leads 
from our CRM. When triggered, it should:
1. Read the lead data from the webhook payload
2. Enrich the email using AnyMailFinder
3. Add to our master Google Sheet
4. Send a Slack notification
```

**Step 2: Add to webhooks.json**

The agent will update [execution/webhooks.json](execution/webhooks.json):
```json
{
  "webhooks": {
    "process-lead": {
      "directive": "process_lead_webhook",
      "description": "Enrich new leads and notify team",
      "tools": ["enrich_emails", "update_sheet", "send_slack"]
    }
  }
}
```

**Step 3: Deploy to Modal**

```powershell
modal deploy execution/modal_webhook.py
```

**Step 4: Get Your Endpoint URLs**

Endpoint format:
```
Endpoints:
- https://[your-username]--claude-orchestrator-list-webhooks.modal.run - List webhooks
- https://[your-username]--claude-orchestrator-directive.modal.run?slug={slug} - Execute directive
```

### Scheduled Triggers

For recurring tasks, use Modal's cron scheduling:

```python
@app.function(schedule=modal.Cron("0 9 * * MON"))  # Every Monday at 9 AM
def weekly_report():
    # Your automation code
    pass
```

### Monitoring

> **All webhook activity streams to Slack in real-time.**

Configure your `SLACK_WEBHOOK_URL` in `.env` for notifications.

### Using Trigger.dev (Node.js Alternative)

This workspace also includes [package.json](package.json) configured for Trigger.dev—useful if you prefer TypeScript or need complex workflow orchestration.

**Setup:**
```powershell
# Install dependencies
npm install

# Start local development server
npm run dev

# Deploy to Trigger.dev cloud
npm run deploy
```

**When to use Trigger.dev over Modal:**
- You're building TypeScript-first workflows
- You need visual workflow monitoring dashboard
- You want built-in retry logic and error handling
- You're integrating with other Node.js services

**Example Trigger.dev task:**
```typescript
import { task } from "@trigger.dev/sdk/v3";
import Anthropic from "@anthropic-ai/sdk";

export const enrichLeadTask = task({
  id: "enrich-lead",
  run: async (payload: { email: string }) => {
    // Use Claude to enrich the lead
    const anthropic = new Anthropic();
    // ... enrichment logic
    return enrichedLead;
  },
});
```

---

## Running Multiple Agents in Parallel

### Why Parallelize?

- Explore more solutions in same time
- Different approaches for comparison
- Avoid single-point failures

### Practical Day-to-Day: Managing Multiple Workflows

Here's what a productive agentic workflow session looks like—running three agent instances for three different business functions:

**Instance 1 (Left): Agency Client Work**
```
Run the post kickoff flow for the demo kickoff call transcript in .tmp/
```
→ Agent analyzes transcript, scrapes leads, enriches emails, creates campaign

**Instance 2 (Middle): Content Research**
```
Run the YouTube outlier workflow and find me 10-20 outliers for "agentic workflows"
```
→ Agent scrapes YouTube API, finds high-performing videos, returns ideas

**Instance 3 (Right): Community Management**
```
Pull the top 10 most recent school posts from my community
```
→ Agent pulls posts, you reply to questions, it formats and sends responses

**The Result**: In 30 minutes of orchestration, you accomplish what would take 3-4 hours of sequential work.

### Practical Limits

| Agents | Recommendation |
|--------|----------------|
| 2 | Good baseline |
| 3-4 | Soft maximum |
| 5+ | Usually counterproductive |

### Multi-Approach Building

1. Ask one agent: "Give me 3 distinct approaches to build this lead scraper"
2. Open 3 agent instances
3. Assign one approach to each
4. Have each work in separate `/tmp` folder
5. Compare results
6. Pick winner, move to main workspace

### Cross-Model Parallelization

Use different models to stay under rate limits:
- Instance 1: Claude
- Instance 2: Gemini  
- Instance 3: GPT-5

Same system prompts = interchangeable behavior

---

## Sub-Agents

Sub-agents solve the **context pollution** problem by isolating tasks.

### Understanding Context Pollution

This is one of the most critical concepts for maintaining workflow quality.

**The Problem**: As your conversation continues, the context window fills with:
- Previous messages
- Tool outputs
- MCP function definitions
- Intermediate debugging steps
- Error logs and retries

**The Science**: Research shows performance degrades significantly as tokens increase:

| Documents in Context | Accuracy |
|---------------------|----------|
| 5 documents | ~75% |
| 10 documents | ~65% |
| 20 documents | ~55% |
| 40 documents | ~50% |

This is why your agent might nail a task at the start of a session but struggle with similar tasks after an hour of back-and-forth.

**Token Budget Reality Check** (use `/context` in Claude Code to see yours):
- System prompt: ~1-2%
- MCP tools: **8-15%** (each tool = 200-400 tokens!)
- Your messages: Variable
- Agent messages: Variable

With 5 MCP servers × 10 tools × 300 tokens = **15,000 tokens gone before you do anything**.

### The Sub-Agent Solution

Sub-agents get **fresh, clean context windows**:
1. Parent agent spawns sub-agent
2. Sub-agent receives only relevant task info
3. Sub-agent works independently
4. Returns only essential results to parent
5. Sub-agent's working context is discarded

### Architecture

```
Parent Agent
    ├── Sub-Agent A (research task)
    ├── Sub-Agent B (code review)
    └── Sub-Agent C (documentation)
```

**Note**: Sub-agents cannot spawn more sub-agents (safety constraint).

### Two Recommended Sub-Agents

**1. Reviewer Sub-Agent**
- Fresh eyes on execution scripts
- Evaluates quality without bias
- Suggests improvements
- Like having someone else proofread your essay

```
Create a reviewer sub-agent that examines execution scripts 
with fresh eyes and provides improvement suggestions to 
the main agent.
```

**2. Document Sub-Agent**
- Reviews scripts after self-annealing
- Updates directives to match current script behavior
- Keeps documentation synchronized

```
Create a document sub-agent that reviews scripts and 
updates directives so everything stays aligned.
```

### Permissions

Give sub-agents **least privilege**:
- Reviewer: Read execution scripts, no write access
- Document: Read executions, write only to directives

---

## Best Practices and Safety Guidelines

### Human-in-the-Loop Decision Framework

**Always involve humans when**:
- High magnitude outcomes (major business impact)
- High quality sensitivity (small errors = big problems)
- Examples: Cold email templates, financial documents, proposals

**Safe to automate**:
- Low sensitivity tasks (web scraping, data formatting)
- Linear quality-to-impact relationship
- Multiple drafts for later selection

### The 10x Rule for Optimization

Only optimize if you can achieve **10x improvement** in a key metric:
- 3 minutes → 2 minutes is not worth the risk
- 20 minutes → 2 minutes is worth pursuing

Small optimizations reduce reliability for marginal gains.

### Security Best Practices

1. **Never share API keys**
   - Keep in .env file
   - Never hardcode in scripts

2. **Use containers for autonomous agents**
   - Don't give full system access
   - Isolate execution environment

3. **Set spending caps**
   - API usage limits
   - Check before expensive operations

4. **Log everything**
   - Maintain audit trails
   - Enable rollback capabilities

### Building Your Workflow Library

Every workflow you build becomes a **permanent reusable asset**:

**What Your Library Can Eventually Do:**
- Automated lead scraping
- Email enrichment
- Personalized cold email replies
- Voice agent calls (initiate AI calls to prospects)
- Proposal generation
- Slide deck creation (matching your brand/tone)
- CRM management
- Community engagement
- Content ideation

**Portability:**
- Copy directives across workspaces
- Share with team members (just copy the folder)
- Put on GitHub for collaborative development
- Deploy anywhere via Modal webhooks

**The Compound Effect:**

| Week | Workflows Built | Cumulative Value |
|------|----------------|------------------|
| 1 | Lead scraper | 1 workflow |
| 2 | Email enrichment | 2 workflows (chain them!) |
| 3 | Personalization | 3 workflows (full pipeline!) |
| 4 | Proposal gen | 4 workflows |
| 8 | CRM wrapper | Multiple meta-directives |
| 12 | Full business ops | Your IDE runs your business |

Over time, your IDE becomes a **treasure chest** deployable anywhere you want.

**Real Example**: Automating forgotten tasks

> "I kept forgetting to post a weekly community call thread. Gave it to my agent, asked if automating it was straightforward. It found a pre-existing school system I'd built, created a scraping spec, and automated my school posts in 3 minutes flat using a schedule timer."

This is the power of accumulating workflows—tiny problems that bug you become trivial to solve.

---

## Quick Reference: This Workspace

### Complete File Map

```
DOE Framework Agentic AI/
├── 📄 AGENTS.md                      # Universal system prompt for AI agents
├── 📄 README.md                      # This tutorial
├── 📄 requirements.txt               # Python dependencies (pip install -r)
├── 📄 package.json                   # Node.js dependencies (npm install)
├── 📁 directives/                    # The "WHAT" layer
│   ├── create_agent_workspace.md     # 🆕 Generate new agent workspaces
│   ├── create_proposal.md            # PandaDoc proposal generation
│   ├── gmaps_lead_generation.md      # Google Maps lead enrichment
│   ├── google_serp_lead_scraper.md   # Google search scraping
│   ├── instantly_autoreply.md        # Cold email automation
│   ├── jump_cut_vad.md               # Video editing
│   ├── scrape_leads.md               # Apify lead scraping
│   └── upwork_scrape_apply.md        # Upwork automation
├── 📁 execution/                     # The "HOW" layer
│   ├── 🏭 Workspace Generator
│   │   ├── create_agent_workspace.py # 🆕 Generate agent workspaces
│   │   ├── agent_templates.json      # Agent type configurations
│   │   └── templates/                # Base templates
│   │       └── AGENTS_BASE.md
│   ├── 📊 Lead Generation
│   │   ├── scrape_apify.py
│   │   ├── scrape_apify_parallel.py
│   │   ├── scrape_google_maps.py
│   │   ├── gmaps_lead_pipeline.py
│   │   ├── gmaps_parallel_pipeline.py
│   │   └── enrich_emails.py
│   ├── 📋 Google Sheets
│   │   ├── read_sheet.py
│   │   ├── update_sheet.py
│   │   └── append_to_sheet.py
│   ├── 📧 Email & Proposals
│   │   ├── create_proposal.py
│   │   ├── welcome_client_emails.py
│   │   ├── instantly_autoreply.py
│   │   └── instantly_create_campaigns.py
│   ├── ✏️ Text Processing
│   │   ├── casualize_batch.py
│   │   ├── casualize_first_names_batch.py
│   │   ├── casualize_company_names_batch.py
│   │   └── casualize_city_names_batch.py
│   ├── 💼 Upwork
│   │   ├── upwork_scraper.py
│   │   ├── upwork_apify_scraper.py
│   │   └── upwork_proposal_generator.py
│   ├── ☁️ Cloud & Webhooks
│   │   ├── modal_webhook.py
│   │   ├── webhooks.json
│   │   └── onboarding_post_kickoff.py
│   └── 🎬 Video
│       ├── jump_cut_vad_singlepass.py
│       └── insert_3d_transition.py
├── 📁 outputs/                       # 🆕 Generated agent workspaces (git-ignored)
├── 📁 .tmp/                          # Temporary files (git-ignored)
└── 📄 .env                           # API keys (create yourself)
```

### Quick Commands for Copilot Chat

| Task | Command |
|------|---------|
| **Scrape Leads** | "Scrape 200 plumbers in Austin, Texas" |
| **Create Proposal** | "Create a proposal for John at ACME Corp..." |
| **Google Maps Leads** | "Find 50 HVAC contractors in Miami using Google Maps" |
| **Check Webhook Status** | "List all configured webhooks" |
| **Deploy to Cloud** | "Deploy the create_proposal directive as a Modal webhook" |
| **Run Upwork Scraper** | "Find Upwork jobs for Python automation" |

### VS Code + GitHub Copilot Shortcuts

| Shortcut | Action |
|----------|--------|
| `Ctrl+Shift+I` | Open Copilot Chat |
| `Ctrl+I` | Inline Copilot prompt |
| `Ctrl+Enter` | Accept Copilot suggestion |
| `Ctrl+Shift+P` | Command Palette |
| `` Ctrl+` `` | Toggle Terminal |

### The DOE Mantra

> **Directives** ([directives/](directives/)): Tell the agent WHAT to do (natural language)
> **Orchestration** (GitHub Copilot): Agent decides HOW to route (AI flexibility)  
> **Execution** ([execution/](execution/)): Scripts do the WORK (deterministic code)

### Essential .env Variables

```env
# Required for most workflows
ANTHROPIC_API_KEY=         # Claude API
APIFY_API_TOKEN=           # Lead scraping
GOOGLE_CREDENTIALS=        # Google Sheets (base64 encoded)

# Optional depending on workflows
OPENAI_API_KEY=            # GPT fallback
ANYMAILFINDER_API_KEY=     # Email enrichment
PANDADOC_API_KEY=          # Proposal generation
INSTANTLY_API_KEY=         # Cold email
SLACK_WEBHOOK_URL=         # Notifications
```

---

## Conclusion

This workspace is a fully functional agentic workflow system. The key files to understand:

1. **Start here**: [AGENTS.md](AGENTS.md) - The system prompt that teaches the agent the DOE framework
2. **Generate workspaces**: Ask the agent to "create a new agent" or run [create_agent_workspace.py](execution/create_agent_workspace.py)
3. **Learn from examples**: [directives/](directives/) - Real production directives you can copy and modify
4. **Study the code**: [execution/](execution/) - Deterministic scripts that do the actual work
5. **Deploy to cloud**: [execution/modal_webhook.py](execution/modal_webhook.py) - Event-driven automation

The skill of building agentic workflows is now one of the highest-ROI skills available. Use this workspace as your starting point—generate specialized agents, copy directives, modify scripts, and build your own library of automated workflows.

---

*This tutorial was created for the VS Code + GitHub Copilot workflow. All file links reference actual files in this workspace.*
