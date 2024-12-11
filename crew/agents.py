from tools import search_tool
from dotenv import load_dotenv

from crewai import Agent
from langchain_openai import ChatOpenAI
from langchain_groq import ChatGroq

load_dotenv()

llm = ChatOpenAI(model="gpt-4o", temperature=0.1, verbose=True, max_tokens=10000)
# llm = ChatGroq(model="mixtral-8x7b-32768", temperature=0.1, verbose=True)

company_research_agent = Agent(
    role="Company Research Analyst",
    goal="""
        You are an expert researcher working for Crowe where you need to
        perform an in-depth analysis to locate all global branches of our client {company_name}. 
        Identify the headquarters (HQ) country (only the country name) and the sector of operation 
        (which industry the company primarily operates in).

        Provide the data in the form of:
        - A list of countries where the company has operations (branches).
        - The country where the HQ is located.
        - The sector or industry of operation.
    """,
    backstory="""
        You are a highly analytical researcher specializing in gathering accurate and comprehensive 
        information about the global operations of companies. You excel at identifying branch locations, 
        headquarters, and their industries. Your expertise ensures that data about {company_name}, 
        operating in the sector {sector_of_operation}, is precise and reliable.
    """,
    tools=[search_tool],
    llm=llm,  # Optional
    verbose=True,  # Optional
)
