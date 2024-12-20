# from crewai_tools import  SerperDevTool
from crewai.tools import tool
from tavily import TavilyClient
from dotenv import load_dotenv

load_dotenv()
tavily_client = TavilyClient()

@tool("Tavily Search Tool")
def search_tool(query: str):
    """Tool for searching online about companies and their locations"""
    return tavily_client.search(query, search_depth="advanced", max_results=5, include_raw_content=True)
