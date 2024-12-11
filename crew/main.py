from agents import company_research_agent
from crewai import Crew, Task
import pandas as pd
from pydantic import BaseModel
from typing import List
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import time

countries = [
    "Austria",
    "Belgium",
    "Bulgaria",
    "Croatia",
    "Cyprus",
    "Czech Republic",
    "France",
    "Germany",
    "Greece",
    "Hungary",
    "Italy",
    "Luxembourg",
    "Netherlands",
    "Poland",
    "Romania",
    "Serbia",
    "Slovakia",
    "Slovenia",
    "Sweden",
    "Switzerland",
    "Ukraine",
    "United Arab Emirates",
    "United Kingdom",
    "United States of America"
]

# A dictionary to normalize returned country names to the standardized form in 'countries'
country_normalization_map = {
    "united states": "United States of America",
    "us": "United States of America",
    "u.s.": "United States of America",
    "u.s": "United States of America",
    "usa": "United States of America",
    "uae": "United Arab Emirates",
    "u.a.e.": "United Arab Emirates",
    "united arab emirates": "United Arab Emirates",
    "uk": "United Kingdom",
    "u.k.": "United Kingdom",
    "u.k": "United Kingdom",
    "united kingdom": "United Kingdom"
}


def normalize_country_name(country: str) -> str:
    if not country or not isinstance(country, str):
        return country
    # Strip whitespace and lowercase for matching
    lower_country = country.strip().lower()
    # Replace with the mapped name if available
    normalized = country_normalization_map.get(lower_country, country)
    return normalized


class Company(BaseModel):
    Branches: List[str]
    HQ: str
    Sector_of_Operation: str

company_research_task = Task(
    description="""
        Perform an extensive search to locate all global branches of {company_name} that operates in the sector (if given)
        {sector_of_operation} and its HQ Location.

        A JSON format of all the countries for {company_name}, its HQ location with just country name only
        , and its sector of operation which is basically what industry are they in.

        Ensure that your findings are as recent and accurate as possible.
        The output must be a JSON format of the output. Nothing more and nothing less.
    """,
    agent=company_research_agent,
    expected_output="""It must be a JSON format of the output. Nothing more and nothing less such as the example below.

      "Branches": [
        "United Kingdom",
        "United States of America",
        "Greece"
      ],
      "HQ": "United Arab Emirates",
      "Sector of Operation": "pharma industry"
    """,
    output_pydantic=Company
)
path = r"C:\Users\Asus\Documents\Crowe\crew\clients_test.xlsx"

# Load the Excel file into a DataFrame, with headers on the second row
df = pd.read_excel(path, header=1)

# Columns in Excel (adjust if needed):
hq_column_name = "Country of HQ location"    # Ensure this matches your actual header
sector_column_name = "Sector of operation"   # Ensure this matches your actual header

for index, row in df.iterrows():
    client_name = row["Client name"]
    sector_of_operation = row["Sector of operation"]

    # Prepare inputs for Crew
    inputs = {
        "company_name": client_name,
        "sector_of_operation": sector_of_operation,
    }

    # Run the Crew
    crew = Crew(
        agents=[company_research_agent],
        tasks=[company_research_task],
        memory=False,
        verbose=True,
        cache=False,
    )
    result = crew.kickoff(inputs=inputs)

    # Add a 5-second delay after each request to prevent rate limit issues
    time.sleep(15)

    hq = result["HQ"]
    sector = result["Sector_of_Operation"]
    branches = result["Branches"]

    # Normalize the HQ name
    hq = normalize_country_name(hq)
    # Normalize each branch country
    normalized_branches = [normalize_country_name(b) for b in branches]

    # Update HQ column if empty and we have a valid HQ
    if (hq_column_name in df.columns) and pd.isna(df.at[index, hq_column_name]) and hq:
        # Only write HQ if it's in your standardized list or if you just want to trust normalization
        if hq in countries:
            df.at[index, hq_column_name] = hq

    # Update Sector column if empty and we have a valid sector
    if (sector_column_name in df.columns) and pd.isna(df.at[index, sector_column_name]) and sector:
        df.at[index, sector_column_name] = sector

    # For branches, mark "1" in the respective country columns if found in normalized_branches
    for country in countries:
        if (country in df.columns) and (country in normalized_branches) and pd.isna(df.at[index, country]):
            df.at[index, country] = "1"

# Now use openpyxl to preserve formatting
wb = load_workbook(path)
sheet = wb.active

# Adjust if data starts at row 3 in Excel
for index, row in df.iterrows():
    excel_row = index + 3  # If header=1, index=0 corresponds to row 3 in Excel

    # Write HQ if we set it in df
    if hq_column_name in df.columns:
        hq_val = row[hq_column_name]
        if pd.notna(hq_val):
            # HQ is column D, adjust if needed
            sheet["D" + str(excel_row)] = hq_val

    # Write Sector if we set it in df
    if sector_column_name in df.columns:
        sector_val = row[sector_column_name]
        if pd.notna(sector_val):
            # Sector is column H, adjust if needed
            sheet["H" + str(excel_row)] = sector_val

    # Write branch countries (K to AH)
    start_col = 11  # K is the 11th column (A=1, B=2, ..., K=11)
    for idx, country in enumerate(countries):
        if country in df.columns:
            country_val = row[country]
            if pd.notna(country_val):
                col_letter = get_column_letter(start_col + idx)
                sheet[col_letter + str(excel_row)] = country_val

# Save the updated workbook
wb.save(r'C:\Users\Asus\Documents\Crowe\crew\clients_test_updated.xlsx')

print("Processing complete. The updated file with preserved formatting has been saved.")
