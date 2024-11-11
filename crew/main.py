from agents import data_entry_agent, research_agent
from tasks import data_entry_task, research_task

import openpyxl
from crewai import Crew, Process

# Load the Excel file
file_path = "Clients list Test.xlsx"
wb = openpyxl.load_workbook(file_path)
ws = wb.active

# Iterate over each row, starting from row 3 to skip the empty row (row 1) and header row (row 2)
for row in ws.iter_rows(min_row=3, max_row=ws.max_row, values_only=True):
    client_name = row[1]  # Column B: Client name
    hq_company_name = row[2]  # Column C: HQ company name (parent/mother company name)
    sector_of_operation = row[7]  # Column H: Sector of operation
    
    # Prepare the input dictionary for the crew
    inputs = {
        "client_name": client_name,
        "hq_company_name": hq_company_name,
        "sector_of_operation": sector_of_operation
    }
    
    # Call the crew with the prepared inputs
    crew = Crew(
        agents=[research_agent, data_entry_agent],  # Define the agents in advance
        tasks=[research_task, data_entry_task],
        process=Process.sequential,
          verbose = True
    )
    
    # Execute the crew for each row with specific inputs
    result = crew.kickoff(inputs=inputs)
    print(result)

# Close the workbook if needed
wb.close()
