# from crewai import Task
# from agents import company_research_agent
# from pydantic import BaseModel
# from typing import List

# class Company(BaseModel):
#     Branches: List[str]
#     HQ: str
#     Sector_of_Operation: str

# company_research_task = Task(
#     description="""
#         Perform an extensive search to locate all global branches of {company_name} that operates in the sector (if given)
#         {sector_of_operation} and its HQ Location.

#         A JSON format of all the countries for {company_name}, its HQ location with just country name only
#         , and its sector of operation which is basically what industry are they in.

#         Ensure that your findings are as recent and accurate as possible.
#         The output must be a JSON format of the output. Nothing more and nothing less.
#     """,
#     agent=company_research_agent,
#     expected_output="""It must be a JSON format of the output. Nothing more and nothing less such as the example below.

#       "Branches": [
#         "United Kingdom",
#         "United States of America",
#         "Greece"
#       ],
#       "HQ": "United Arab Emirates",
#       "Sector of Operation": "pharma industry"
#     """,
#     output_pydantic=Company
# )
