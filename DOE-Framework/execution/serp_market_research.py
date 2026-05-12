#!/usr/bin/env python3
"""
Market research using SerpAPI - search Google, analyze competitors, track trends.
Part of the Business Planning Agent toolkit.
"""

import os
import sys
import json
import argparse
from datetime import datetime
from dotenv import load_dotenv

try:
    from serpapi import GoogleSearch
except ImportError:
    print("Installing serpapi...")
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "google-search-results"])
    from serpapi import GoogleSearch

load_dotenv()


def search_market(query, num_results=10, location="United States", language="en"):
    """
    Perform a Google search for market research.
    
    Args:
        query: Search query
        num_results: Number of results to return
        location: Geographic location for search
        language: Language code
        
    Returns:
        Dictionary with search results
    """
    api_key = os.getenv("SERPAPI_API_KEY")
    if not api_key:
        raise ValueError("SERPAPI_API_KEY environment variable not set")
    
    params = {
        "engine": "google",
        "q": query,
        "api_key": api_key,
        "num": num_results,
        "location": location,
        "hl": language
    }
    
    search = GoogleSearch(params)
    results = search.get_dict()
    
    # Extract organic results
    organic_results = []
    for result in results.get("organic_results", []):
        organic_results.append({
            "title": result.get("title"),
            "link": result.get("link"),
            "snippet": result.get("snippet"),
            "position": result.get("position")
        })
    
    # Extract related questions (People Also Ask)
    related_questions = []
    for question in results.get("related_questions", []):
        related_questions.append({
            "question": question.get("question"),
            "snippet": question.get("snippet")
        })
    
    return {
        "query": query,
        "total_results": results.get("search_information", {}).get("total_results"),
        "organic_results": organic_results,
        "related_questions": related_questions,
        "timestamp": datetime.now().isoformat()
    }


def analyze_competitors(company_name, industry=None, num_results=20):
    """
    Analyze competitors for a given company or industry.
    
    Args:
        company_name: Name of the company
        industry: Optional industry context
        num_results: Number of results
        
    Returns:
        Dictionary with competitor analysis
    """
    api_key = os.getenv("SERPAPI_API_KEY")
    if not api_key:
        raise ValueError("SERPAPI_API_KEY environment variable not set")
    
    # Search for competitors
    query = f"{company_name} competitors"
    if industry:
        query = f"{company_name} competitors {industry}"
    
    params = {
        "engine": "google",
        "q": query,
        "api_key": api_key,
        "num": num_results
    }
    
    search = GoogleSearch(params)
    results = search.get_dict()
    
    competitors = []
    for result in results.get("organic_results", []):
        competitors.append({
            "title": result.get("title"),
            "link": result.get("link"),
            "snippet": result.get("snippet")
        })
    
    # Also search for market share
    market_query = f"{industry or company_name} market share analysis"
    params["q"] = market_query
    market_search = GoogleSearch(params)
    market_results = market_search.get_dict()
    
    market_insights = []
    for result in market_results.get("organic_results", [])[:5]:
        market_insights.append({
            "title": result.get("title"),
            "link": result.get("link"),
            "snippet": result.get("snippet")
        })
    
    return {
        "company": company_name,
        "industry": industry,
        "competitors": competitors,
        "market_insights": market_insights,
        "timestamp": datetime.now().isoformat()
    }


def get_industry_trends(industry, timeframe="today 12-m"):
    """
    Get industry trends using Google Trends.
    
    Args:
        industry: Industry or keyword to track
        timeframe: Time range (today 12-m, today 3-m, today 1-m)
        
    Returns:
        Dictionary with trend data
    """
    api_key = os.getenv("SERPAPI_API_KEY")
    if not api_key:
        raise ValueError("SERPAPI_API_KEY environment variable not set")
    
    params = {
        "engine": "google_trends",
        "q": industry,
        "api_key": api_key,
        "date": timeframe,
        "data_type": "TIMESERIES"
    }
    
    search = GoogleSearch(params)
    results = search.get_dict()
    
    interest_over_time = results.get("interest_over_time", {})
    timeline_data = interest_over_time.get("timeline_data", [])
    
    trend_points = []
    for point in timeline_data:
        trend_points.append({
            "date": point.get("date"),
            "values": point.get("values", [])
        })
    
    # Get related queries
    params["data_type"] = "RELATED_QUERIES"
    related_search = GoogleSearch(params)
    related_results = related_search.get_dict()
    
    related_queries = {
        "rising": related_results.get("related_queries", {}).get("rising", []),
        "top": related_results.get("related_queries", {}).get("top", [])
    }
    
    return {
        "industry": industry,
        "timeframe": timeframe,
        "trend_data": trend_points,
        "related_queries": related_queries,
        "timestamp": datetime.now().isoformat()
    }


def search_news(query, num_results=10):
    """
    Search for news articles related to a topic.
    
    Args:
        query: Search query
        num_results: Number of results
        
    Returns:
        Dictionary with news results
    """
    api_key = os.getenv("SERPAPI_API_KEY")
    if not api_key:
        raise ValueError("SERPAPI_API_KEY environment variable not set")
    
    params = {
        "engine": "google_news",
        "q": query,
        "api_key": api_key,
        "num": num_results
    }
    
    search = GoogleSearch(params)
    results = search.get_dict()
    
    news_results = []
    for article in results.get("news_results", []):
        news_results.append({
            "title": article.get("title"),
            "link": article.get("link"),
            "source": article.get("source", {}).get("name"),
            "date": article.get("date"),
            "snippet": article.get("snippet")
        })
    
    return {
        "query": query,
        "news_results": news_results,
        "timestamp": datetime.now().isoformat()
    }


def main():
    parser = argparse.ArgumentParser(description='Market Research with SerpAPI')
    parser.add_argument('--mode', required=True,
                        choices=['search', 'competitors', 'trends', 'news'],
                        help='Research mode')
    parser.add_argument('--query', required=True, help='Search query or company name')
    parser.add_argument('--industry', help='Industry context (for competitor analysis)')
    parser.add_argument('--num-results', type=int, default=10, help='Number of results')
    parser.add_argument('--location', default='United States', help='Geographic location')
    parser.add_argument('--timeframe', default='today 12-m',
                        help='Timeframe for trends (today 12-m, today 3-m, today 1-m)')
    parser.add_argument('--output', help='Output JSON file path')

    args = parser.parse_args()

    try:
        if args.mode == 'search':
            result = search_market(
                query=args.query,
                num_results=args.num_results,
                location=args.location
            )
        elif args.mode == 'competitors':
            result = analyze_competitors(
                company_name=args.query,
                industry=args.industry,
                num_results=args.num_results
            )
        elif args.mode == 'trends':
            result = get_industry_trends(
                industry=args.query,
                timeframe=args.timeframe
            )
        elif args.mode == 'news':
            result = search_news(
                query=args.query,
                num_results=args.num_results
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
