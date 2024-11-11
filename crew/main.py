from agents import data_entry_agent, research_agent
from tasks import data_entry_task, research_task
from crewai import Crew, Process
import pandas as pd

# Load the Excel file into a DataFrame just once
file_path = "Clients list Test.xlsx"
df = pd.read_excel(file_path)

# Iterate over each row, starting from row 3 to skip the empty row (row 1) and header row (row 2)
for row in df.itertuples(index=False):
    client_name = row[1]  # Column B: Client name
    sector_of_operation = row[7]  # Column H: Sector of operation
    
    # Prepare the input dictionary for the crew (excluding the DataFrame)
    inputs = {
        "client_name": client_name,
        "sector_of_operation": sector_of_operation,
    }
    
    # Call the crew with the prepared inputs
    crew = Crew(
        agents=[research_agent, data_entry_agent],  # Define the agents in advance
        tasks=[research_task, data_entry_task],
        process=Process.sequential,
        verbose=True
    )
    
    # Execute the crew for each row with specific inputs
    result = crew.kickoff(inputs=inputs)
    print(result)

# After processing, save the DataFrame back to Excel or any desired format if needed
df.to_excel("Updated_Clients_List.xlsx", index=False)
