#!/usr/bin/env python3
"""
Generate business plan components using AI - SWOT analysis, financial projections,
business model canvas, and compile full business plans.
Part of the Business Planning Agent toolkit.

Supports two modes:
1. Copilot Mode (default): Outputs structured prompts for Copilot to process
2. Standalone Mode: Uses OPENAI_API_KEY or ANTHROPIC_API_KEY for direct LLM calls

When running with GitHub Copilot as orchestrator, LLM API keys are optional.
"""

import os
import sys
import json
import argparse
from datetime import datetime
from dotenv import load_dotenv

load_dotenv()

# Check if we have LLM API keys available
HAS_LLM_KEYS = bool(os.getenv("OPENAI_API_KEY") or os.getenv("ANTHROPIC_API_KEY"))


def get_llm_client():
    """Get the appropriate LLM client (OpenAI or Anthropic). Returns None if no keys available."""
    openai_key = os.getenv("OPENAI_API_KEY")
    anthropic_key = os.getenv("ANTHROPIC_API_KEY")
    
    if anthropic_key:
        try:
            import anthropic
            return ("anthropic", anthropic.Anthropic(api_key=anthropic_key))
        except ImportError:
            pass
    
    if openai_key:
        try:
            from openai import OpenAI
            return ("openai", OpenAI(api_key=openai_key))
        except ImportError:
            pass
    
    return None  # No LLM available - Copilot mode


def call_llm(prompt, system_prompt=None, max_tokens=4000):
    """Call the LLM with the given prompt. Returns None if no LLM available (Copilot mode)."""
    client_info = get_llm_client()
    
    if client_info is None:
        # Copilot mode - return the prompt for Copilot to process
        return None
    
    provider, client = client_info
    
    if provider == "anthropic":
        messages = [{"role": "user", "content": prompt}]
        response = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=max_tokens,
            system=system_prompt or "You are a business strategy expert.",
            messages=messages
        )
        return response.content[0].text
    else:
        messages = []
        if system_prompt:
            messages.append({"role": "system", "content": system_prompt})
        messages.append({"role": "user", "content": prompt})
        
        response = client.chat.completions.create(
            model="gpt-4o",
            messages=messages,
            max_tokens=max_tokens
        )
        return response.choices[0].message.content


def generate_swot_analysis(company_name, industry, description, market_data=None, copilot_mode=False):
    """
    Generate a SWOT analysis for a business.
    
    Args:
        company_name: Name of the company
        industry: Industry sector
        description: Business description
        market_data: Optional market research data
        copilot_mode: If True, return prompt for Copilot to process
        
    Returns:
        Dictionary with SWOT analysis or prompt data for Copilot
    """
    context = f"""
Company: {company_name}
Industry: {industry}
Description: {description}
"""
    if market_data:
        context += f"\nMarket Research Data:\n{json.dumps(market_data, indent=2)}"
    
    prompt = f"""
Based on the following business information, create a comprehensive SWOT analysis:

{context}

Provide the analysis in the following JSON format:
{{
    "strengths": ["strength 1", "strength 2", ...],
    "weaknesses": ["weakness 1", "weakness 2", ...],
    "opportunities": ["opportunity 1", "opportunity 2", ...],
    "threats": ["threat 1", "threat 2", ...],
    "summary": "Brief executive summary of the SWOT analysis"
}}

Focus on actionable insights relevant to the business strategy.
Return ONLY the JSON object, no additional text.
"""
    
    system_prompt = "You are a business strategy consultant specializing in SWOT analysis. Provide actionable, specific insights."
    
    # Check if we should use Copilot mode
    if copilot_mode or not HAS_LLM_KEYS:
        # Return prompt data for Copilot to process
        return {
            "mode": "copilot",
            "task": "swot_analysis",
            "company": company_name,
            "industry": industry,
            "context": context,
            "prompt": prompt,
            "system_prompt": system_prompt,
            "expected_format": "JSON with strengths, weaknesses, opportunities, threats, summary",
            "timestamp": datetime.now().isoformat(),
            "message": "No LLM API keys configured. Process this prompt with Copilot or provide OPENAI_API_KEY/ANTHROPIC_API_KEY."
        }
    
    response = call_llm(prompt, system_prompt)
    
    # Parse JSON from response
    try:
        # Try to extract JSON from response
        start = response.find('{')
        end = response.rfind('}') + 1
        json_str = response[start:end]
        swot = json.loads(json_str)
    except:
        swot = {"raw_analysis": response}
    
    swot["company"] = company_name
    swot["industry"] = industry
    swot["timestamp"] = datetime.now().isoformat()
    
    return swot


def generate_financial_projections(company_name, industry, revenue_model, 
                                    initial_investment=None, years=5, assumptions=None, copilot_mode=False):
    """
    Generate financial projections for a business plan.
    
    Args:
        company_name: Name of the company
        industry: Industry sector
        revenue_model: Description of revenue model
        initial_investment: Starting capital
        years: Number of years to project
        assumptions: Key assumptions for projections
        copilot_mode: If True, return prompt for Copilot to process
        
    Returns:
        Dictionary with financial projections or prompt data for Copilot
    """
    context = f"""
Company: {company_name}
Industry: {industry}
Revenue Model: {revenue_model}
Initial Investment: {initial_investment or 'Not specified'}
Projection Period: {years} years
"""
    if assumptions:
        context += f"\nKey Assumptions:\n{json.dumps(assumptions, indent=2)}"
    
    prompt = f"""
Based on the following business information, create realistic financial projections:

{context}

Provide projections in the following JSON format:
{{
    "revenue_projections": {{
        "year_1": {{"revenue": 0, "growth_rate": "X%"}},
        "year_2": {{"revenue": 0, "growth_rate": "X%"}},
        ...
    }},
    "expense_breakdown": {{
        "fixed_costs": {{"item": amount, ...}},
        "variable_costs": {{"item": "X% of revenue", ...}}
    }},
    "profitability": {{
        "gross_margin": "X%",
        "net_margin_year_1": "X%",
        "net_margin_year_5": "X%",
        "break_even_month": 0
    }},
    "key_metrics": {{
        "customer_acquisition_cost": 0,
        "lifetime_value": 0,
        "ltv_cac_ratio": 0
    }},
    "assumptions": ["assumption 1", ...],
    "risks": ["risk 1", ...],
    "summary": "Executive summary of financial outlook"
}}

Use realistic industry benchmarks. Return ONLY the JSON object.
"""
    
    system_prompt = "You are a financial analyst specializing in startup and business projections. Provide realistic, data-driven projections."
    
    # Check if we should use Copilot mode
    if copilot_mode or not HAS_LLM_KEYS:
        return {
            "mode": "copilot",
            "task": "financial_projections",
            "company": company_name,
            "industry": industry,
            "context": context,
            "prompt": prompt,
            "system_prompt": system_prompt,
            "expected_format": "JSON with revenue_projections, expense_breakdown, profitability, key_metrics",
            "timestamp": datetime.now().isoformat(),
            "message": "No LLM API keys configured. Process this prompt with Copilot or provide OPENAI_API_KEY/ANTHROPIC_API_KEY."
        }
    
    response = call_llm(prompt, system_prompt)
    
    try:
        start = response.find('{')
        end = response.rfind('}') + 1
        json_str = response[start:end]
        projections = json.loads(json_str)
    except:
        projections = {"raw_projections": response}
    
    projections["company"] = company_name
    projections["projection_years"] = years
    projections["timestamp"] = datetime.now().isoformat()
    
    return projections


def generate_business_model_canvas(company_name, industry, description, target_market=None, copilot_mode=False):
    """
    Generate a Business Model Canvas.
    
    Args:
        company_name: Name of the company
        industry: Industry sector
        description: Business description
        target_market: Target market description
        copilot_mode: If True, return prompt for Copilot to process
        
    Returns:
        Dictionary with Business Model Canvas or prompt data for Copilot
    """
    context = f"""
Company: {company_name}
Industry: {industry}
Description: {description}
Target Market: {target_market or 'Not specified'}
"""
    
    prompt = f"""
Based on the following business information, create a comprehensive Business Model Canvas:

{context}

Provide the canvas in the following JSON format:
{{
    "key_partners": ["partner 1", "partner 2", ...],
    "key_activities": ["activity 1", "activity 2", ...],
    "key_resources": ["resource 1", "resource 2", ...],
    "value_propositions": ["value prop 1", "value prop 2", ...],
    "customer_relationships": ["relationship type 1", ...],
    "channels": ["channel 1", "channel 2", ...],
    "customer_segments": ["segment 1", "segment 2", ...],
    "cost_structure": ["cost 1", "cost 2", ...],
    "revenue_streams": ["stream 1", "stream 2", ...],
    "summary": "Brief description of the business model"
}}

Be specific and actionable. Return ONLY the JSON object.
"""
    
    system_prompt = "You are a business model innovation expert. Create practical, implementable business model canvases."
    
    # Check if we should use Copilot mode
    if copilot_mode or not HAS_LLM_KEYS:
        return {
            "mode": "copilot",
            "task": "business_model_canvas",
            "company": company_name,
            "industry": industry,
            "context": context,
            "prompt": prompt,
            "system_prompt": system_prompt,
            "expected_format": "JSON with key_partners, key_activities, key_resources, value_propositions, etc.",
            "timestamp": datetime.now().isoformat(),
            "message": "No LLM API keys configured. Process this prompt with Copilot or provide OPENAI_API_KEY/ANTHROPIC_API_KEY."
        }
    
    response = call_llm(prompt, system_prompt)
    
    try:
        start = response.find('{')
        end = response.rfind('}') + 1
        json_str = response[start:end]
        canvas = json.loads(json_str)
    except:
        canvas = {"raw_canvas": response}
    
    canvas["company"] = company_name
    canvas["industry"] = industry
    canvas["timestamp"] = datetime.now().isoformat()
    
    return canvas


def compile_business_plan(company_name, industry, description, 
                          swot=None, financials=None, canvas=None,
                          executive_summary=None, market_research=None, copilot_mode=False):
    """
    Compile a full business plan document from components.
    
    Args:
        company_name: Name of the company
        industry: Industry sector
        description: Business description
        swot: SWOT analysis data
        financials: Financial projections data
        canvas: Business Model Canvas data
        executive_summary: Custom executive summary
        market_research: Market research data
        copilot_mode: If True, return prompt for Copilot to process
        
    Returns:
        Dictionary with compiled business plan or prompt data for Copilot
    """
    components = f"""
Company: {company_name}
Industry: {industry}
Description: {description}

SWOT Analysis: {json.dumps(swot, indent=2) if swot else 'Not provided'}

Financial Projections: {json.dumps(financials, indent=2) if financials else 'Not provided'}

Business Model Canvas: {json.dumps(canvas, indent=2) if canvas else 'Not provided'}

Market Research: {json.dumps(market_research, indent=2) if market_research else 'Not provided'}

Custom Executive Summary: {executive_summary or 'Generate one'}
"""
    
    prompt = f"""
Based on the following business information and components, compile a comprehensive business plan:

{components}

Create a professional business plan with the following sections:
1. Executive Summary
2. Company Description
3. Market Analysis
4. Organization & Management
5. Service/Product Line
6. Marketing & Sales Strategy
7. Financial Projections
8. Funding Requirements (if applicable)
9. Appendix (key data points)

Format the output as a well-structured document that can be directly used in a Google Doc.
Use clear headings and bullet points where appropriate.
"""
    
    system_prompt = "You are a business plan writer with expertise in creating investor-ready documents. Write clear, compelling, and professional business plans."
    
    # Check if we should use Copilot mode
    if copilot_mode or not HAS_LLM_KEYS:
        return {
            "mode": "copilot",
            "task": "compile_business_plan",
            "company": company_name,
            "industry": industry,
            "components": components,
            "prompt": prompt,
            "system_prompt": system_prompt,
            "expected_format": "Full business plan document with all sections",
            "components_provided": {
                "swot": swot is not None,
                "financials": financials is not None,
                "canvas": canvas is not None,
                "market_research": market_research is not None
            },
            "timestamp": datetime.now().isoformat(),
            "message": "No LLM API keys configured. Process this prompt with Copilot or provide OPENAI_API_KEY/ANTHROPIC_API_KEY."
        }
    
    response = call_llm(prompt, system_prompt, max_tokens=8000)
    
    return {
        "company": company_name,
        "industry": industry,
        "business_plan": response,
        "components_used": {
            "swot": swot is not None,
            "financials": financials is not None,
            "canvas": canvas is not None,
            "market_research": market_research is not None
        },
        "timestamp": datetime.now().isoformat()
    }


def main():
    parser = argparse.ArgumentParser(description='Business Plan Generator (supports Copilot mode)')
    parser.add_argument('--mode', required=True,
                        choices=['swot', 'financials', 'canvas', 'compile'],
                        help='Generation mode')
    parser.add_argument('--company', required=True, help='Company name')
    parser.add_argument('--industry', required=True, help='Industry sector')
    parser.add_argument('--description', help='Business description')
    parser.add_argument('--revenue-model', help='Revenue model (for financials)')
    parser.add_argument('--investment', type=float, help='Initial investment')
    parser.add_argument('--years', type=int, default=5, help='Projection years')
    parser.add_argument('--target-market', help='Target market')
    parser.add_argument('--market-data', help='Path to market research JSON')
    parser.add_argument('--swot-data', help='Path to SWOT analysis JSON')
    parser.add_argument('--financial-data', help='Path to financial projections JSON')
    parser.add_argument('--canvas-data', help='Path to business model canvas JSON')
    parser.add_argument('--output', help='Output JSON file path')
    parser.add_argument('--copilot', action='store_true', 
                        help='Force Copilot mode (output prompts instead of calling LLM)')

    args = parser.parse_args()
    
    # Determine if we're in Copilot mode
    copilot_mode = args.copilot or not HAS_LLM_KEYS
    
    if copilot_mode:
        print("Running in Copilot mode (LLM API keys not configured or --copilot flag used)")
        print("Output will contain prompts for Copilot to process.\n")

    # Load external data files if provided
    market_data = None
    if args.market_data and os.path.exists(args.market_data):
        with open(args.market_data, 'r') as f:
            market_data = json.load(f)

    swot_data = None
    if args.swot_data and os.path.exists(args.swot_data):
        with open(args.swot_data, 'r') as f:
            swot_data = json.load(f)

    financial_data = None
    if args.financial_data and os.path.exists(args.financial_data):
        with open(args.financial_data, 'r') as f:
            financial_data = json.load(f)

    canvas_data = None
    if args.canvas_data and os.path.exists(args.canvas_data):
        with open(args.canvas_data, 'r') as f:
            canvas_data = json.load(f)

    try:
        if args.mode == 'swot':
            result = generate_swot_analysis(
                company_name=args.company,
                industry=args.industry,
                description=args.description or '',
                market_data=market_data,
                copilot_mode=copilot_mode
            )
        elif args.mode == 'financials':
            result = generate_financial_projections(
                company_name=args.company,
                industry=args.industry,
                revenue_model=args.revenue_model or '',
                initial_investment=args.investment,
                years=args.years,
                copilot_mode=copilot_mode
            )
        elif args.mode == 'canvas':
            result = generate_business_model_canvas(
                company_name=args.company,
                industry=args.industry,
                description=args.description or '',
                target_market=args.target_market,
                copilot_mode=copilot_mode
            )
        elif args.mode == 'compile':
            result = compile_business_plan(
                company_name=args.company,
                industry=args.industry,
                description=args.description or '',
                swot=swot_data,
                financials=financial_data,
                canvas=canvas_data,
                market_research=market_data,
                copilot_mode=copilot_mode
            )

        print(json.dumps(result, indent=2))

        if args.output:
            with open(args.output, 'w') as f:
                json.dump(result, f, indent=2)
            print(f"\nResult saved to: {args.output}")

        return result

    except Exception as e:
        print(f"Error: {e}")
        sys.exit(1)


if __name__ == '__main__':
    main()
