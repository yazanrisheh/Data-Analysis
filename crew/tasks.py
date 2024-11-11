from crewai import Task
from agents import research_agent, data_entry_agent

research_task = Task(
    description="""
        Perform an extensive search to locate all global branches of {client_name} that operates in the sector
        {sector_of_operation}.
        Focus on identifying and verifying the country of each branch location, ensuring accuracy
        and comprehensiveness in coverage.

        Your research should include:
        - The country of each branch
        - Notable regions or clusters of branches, if applicable

        Your final report MUST be a structured list or table of countries with branch presence.

        Ensure that your findings are as recent and accurate as possible.
    """,
    agent=research_agent,
    expected_output="A structured JSON format of countries with branch presence for {client_name}.",
)
data_entry_task = Task(
    description="""
        Verify the existence of the company {client_name} that is operating in the sector {sector_of_operation} in the
        provided Excel sheet.
        For each branch location provided by the research_agent, check if the corresponding country column exists in the
        Excel sheet, starting from the column after 'AH'.
        
        For each country:
        - If the column for the country exists, input the number '1' in the respective row under that country.
        - If the column for the country does not exist, create a new column for that country after the existing columns, 
        starting from the column after 'AH', and add the value '1' in the first available row.
        
        You must ensure that all branches listed in the research_agent's findings are correctly recorded in the Excel sheet.
        The Excel file should only contain '1' entries for valid branches of {client_name} found in specific countries.

        Your final task is to save the updated Excel file with all data entries completed.
    """,
    agent=data_entry_agent,
    expected_output="""An updated Excel file with '1' entered in appropriate columns for each branch country location of
    {client_name} found by the research_agent.""",
    output_file="Updated_Clients_List.xlsx"
)
