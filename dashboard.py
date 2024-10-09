import os
import pandas as pd
import plotly.express as px
import warnings
warnings.filterwarnings("ignore")

# Create 'html' directory if it doesn't exist
if not os.path.exists('html'):
    os.makedirs('html')

# Load the Excel file
file_path = 'نموذج بيانات الموظفين.xlsx'  # Replace with the correct path to your Excel file
data = pd.read_excel(file_path, sheet_name='بيانات الموظفين')

# Preprocess the data
data['Joining Date \n تاريخ الالتحاق'] = pd.to_datetime(data['Joining Date \n تاريخ الالتحاق'], errors='coerce')
data['Date of Birth \n تاريخ الميلاد'] = pd.to_datetime(data['Date of Birth \n تاريخ الميلاد'], errors='coerce')

# Calculate employee age and tenure
current_year = pd.to_datetime('today').year
data['Age'] = current_year - data['Date of Birth \n تاريخ الميلاد'].dt.year
data['Tenure'] = current_year - data['Joining Date \n تاريخ الالتحاق'].dt.year

# Additional Visualizations

def display_additional_visualizations():
    # 1. Gender Distribution by Job Title
    fig1 = px.histogram(data, x='Job Title \n المسمى الوظيفي', color='Gender \n الجنس', title='Gender Distribution by Job Title',
                        barmode='stack', color_discrete_sequence=px.colors.qualitative.Set1)
    
    # 2. Correlation Between Age and Tenure
    fig2 = px.scatter(data, x='Age', y='Tenure', trendline="ols", title='Correlation Between Age and Tenure', 
                      color_discrete_sequence=px.colors.qualitative.Plotly)
    
    # 3. Employee Growth Over Time
    data['Year of Joining'] = data['Joining Date \n تاريخ الالتحاق'].dt.year
    hires_by_year = data.groupby('Year of Joining').size().cumsum().reset_index(name='Cumulative Count')
    fig3 = px.line(hires_by_year, x='Year of Joining', y='Cumulative Count', title='Employee Growth Over Time', markers=True, 
                   color_discrete_sequence=px.colors.qualitative.Bold)
    
    # 4. Sector-wise Tenure Distribution
    fig4 = px.box(data, x='Sector \n القطاع', y='Tenure', title='Sector-wise Tenure Distribution',
                  color_discrete_sequence=px.colors.qualitative.Plotly)
    
    # 5. Average Job Grade by Sector
# Fix to safely convert job grade values with numeric content, ignoring empty strings
    job_grade_numeric = data['Job Grade\n الدرجة الوظيفية'].apply(lambda x: int(''.join(filter(str.isdigit, str(x)))) if ''.join(filter(str.isdigit, str(x))) else None)
    sector_avg_grade = data.groupby('Sector \n القطاع').apply(lambda x: job_grade_numeric.loc[x.index].mean()).reset_index(name='Avg Grade')
    fig5 = px.bar(sector_avg_grade, x='Sector \n القطاع', y='Avg Grade', title='Average Job Grade by Sector', color='Avg Grade',
                  color_continuous_scale=px.colors.sequential.Viridis)

    # 6. Educational Qualification vs. Tenure
    fig6 = px.histogram(data, x='Educational Qualification \n المؤهل العلمي', y='Tenure', title='Educational Qualification vs Tenure',
                        color_discrete_sequence=px.colors.qualitative.Plotly)
    
    # 7. Nationality vs. Sector Distribution
    fig7 = px.histogram(data, x='Nationality \n الجنسية', color='Sector \n القطاع', title='Nationality vs. Sector Distribution',
                        barmode='stack', color_discrete_sequence=px.colors.qualitative.Set3)
    
    # 8. Top 5 Job Titles by Employee Count
    # 8. Top 5 Job Titles by Employee Count
    # 8. Top 5 Job Titles by Employee Count
    top_job_titles = data['Job Title \n المسمى الوظيفي'].value_counts().nlargest(5).reset_index(name='Count')
    top_job_titles.rename(columns={'index': 'Job Title \n المسمى الوظيفي'}, inplace=True)  # Ensure correct column name
    fig8 = px.bar(top_job_titles, x='Job Title \n المسمى الوظيفي', y='Count', title='Top 5 Job Titles by Employee Count', color='Count', 
              color_continuous_scale=px.colors.sequential.Teal)


    
    # 9. Average Tenure by Department
    department_avg_tenure = data.groupby('Department \n الإدارة')['Tenure'].mean().reset_index(name='Avg Tenure')
    fig9 = px.bar(department_avg_tenure, x='Department \n الإدارة', y='Avg Tenure', title='Average Tenure by Department',
                  color='Avg Tenure', color_continuous_scale=px.colors.sequential.Blues)
    
    # 10. Job Category Breakdown by Gender
    fig10 = px.histogram(data, x='Job Category \n الفئة الوظيفية', color='Gender \n الجنس', title='Job Category Breakdown by Gender',
                         barmode='stack', color_discrete_sequence=px.colors.qualitative.Dark24)
    
    # Return the visualizations
    return [fig1, fig2, fig3, fig4, fig5, fig6, fig7, fig8, fig9, fig10]

# Save visualizations only if the HTML file doesn't already exist
additional_html_filenames = [
    "html/gender_distribution_by_job_title.html",
    "html/age_vs_tenure_correlation.html",
    "html/employee_growth_over_time.html",
    "html/sector_tenure_distribution.html",
    "html/avg_job_grade_by_sector.html",
    "html/education_vs_tenure.html",
    "html/nationality_vs_sector.html",
    "html/top_5_job_titles.html",
    "html/avg_tenure_by_department.html",
    "html/job_category_by_gender.html"
]

# Generate the additional visualizations
additional_visualizations = display_additional_visualizations()

# Save each visualization only if the file doesn't exist
for i, fig in enumerate(additional_visualizations):
    if not os.path.exists(additional_html_filenames[i]):
        fig.write_html(additional_html_filenames[i])

print("Additional interactive HTML files have been saved successfully!")
