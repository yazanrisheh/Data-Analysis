from tools import search_tool
from dotenv import load_dotenv

from crewai import Agent
from langchain_openai import ChatOpenAI
from langchain_groq import ChatGroq

load_dotenv()

llm = ChatOpenAI(model = "gpt-4o-mini", temperature = 0.2, verbose = True)
# llm = ChatGroq(model = "llama3-8b-8192", temperature = 0.2, verbose = True)

research_agent = Agent(
  role='Market Research Analyst',
  goal="""Search for the company name {client_name} whose sector of operation is {sector_of_operation} and find
    the country location for its branches worldwide""",
  backstory="""
    An analytical researcher skilled in gathering and verifying {client_name} working in the sector {sector_of_operation} 
    location data, with an eye for accurate and comprehensive information on global branch distributions.
""",
  tools=[search_tool],
  llm=llm,  # Optional
  max_iter=15,  # Optional
  verbose=True,  # Optional
)


data_entry_agent = Agent(
  role='Data Entry Specialist',
  goal="""Input the number 1 in the excel sheet if the research_agent found the company
    {client_name} that is working in the sector {sector_of_operation} and the country location for all its branches 
    worldwide""",
  backstory="""
You are a data entry specialist that inputs the value 1 if the company {client_name} in the sector {sector_of_operation} 
exists in the country location thats available in the excel file. If it does not exist then start from row 2 to create
a new column for that country location after the column 'AH' and add the value 1
""",
  tools=[],
  llm=llm,  # Optional
  max_iter=15,  # Optional
  verbose=True,  # Optional
  allow_code_execution=True
)
