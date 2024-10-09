import pandas as pd
import plotly.express as px
import warnings
import os

# Ignore warnings
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

# Generate visualizations
def display_visualizations():
    # 1. Employee Age Distribution
    fig1 = px.histogram(data, x='Age', nbins=20, title='Employee Age Distribution', color_discrete_sequence=px.colors.sequential.Blues)
    
    # 2. Job Title by Sector
    fig2 = px.histogram(data, x='Job Title \n المسمى الوظيفي', color='Sector \n القطاع', title='Job Title Distribution by Sector',
                        color_discrete_sequence=px.colors.qualitative.Pastel)
    fig2.update_layout(xaxis_title='Job Title', yaxis_title='Count', barmode='stack')
    
    # 3. Educational Qualification by Gender
    fig3 = px.histogram(data, x='Educational Qualification \n المؤهل العلمي', color='Gender \n الجنس', title='Educational Qualification by Gender',
                        barmode='group', color_discrete_sequence=px.colors.qualitative.Set3)
    
    # 4. Nationality Distribution
    fig4 = px.pie(data, names='Nationality \n الجنسية', title='Nationality Distribution', hole=0.3, color_discrete_sequence=px.colors.sequential.Plasma_r)
    
    # 5. Marital Status Distribution
    fig5 = px.pie(data, names='Marital Status \n الحالة الإجتماعية', title='Marital Status Distribution', hole=0.3, color_discrete_sequence=px.colors.sequential.RdBu)
    
    # 6. Job Grade vs. Gender
    fig6 = px.histogram(data, x='Job Grade\n الدرجة الوظيفية', color='Gender \n الجنس', title='Job Grade Distribution by Gender', 
                        barmode='group', color_discrete_sequence=px.colors.qualitative.D3)
    
    # 7. Department-wise Employee Count
    fig7 = px.histogram(data, x='Department \n الإدارة', title='Department-wise Employee Count', 
                        color_discrete_sequence=px.colors.sequential.Viridis)
    fig7.update_layout(xaxis_title='Department', yaxis_title='Count')
    
    # 8. Employee Tenure Distribution
    fig8 = px.histogram(data, x='Tenure', nbins=10, title='Employee Tenure Distribution', color_discrete_sequence=px.colors.sequential.Teal)
    
    # 9. New Hires Trend (by Joining Date)
    data['Year of Joining'] = data['Joining Date \n تاريخ الالتحاق'].dt.year
    hires_by_year = data.groupby('Year of Joining').size().reset_index(name='Count')
    fig9 = px.line(hires_by_year, x='Year of Joining', y='Count', title='New Hires Over Time', markers=True, 
                   color_discrete_sequence=px.colors.qualitative.Bold)
    
    # 10. Job Category vs. Employee Number
    fig10 = px.histogram(data, x='Job Category \n الفئة الوظيفية', color='Job Category \n الفئة الوظيفية', title='Employee Number by Job Category',
                         color_discrete_sequence=px.colors.qualitative.Pastel1)

    # Return the visualizations
    return [fig1, fig2, fig3, fig4, fig5, fig6, fig7, fig8, fig9, fig10]

# Generate the visualizations
visualizations = display_visualizations()

# File names for saving interactive HTML in the 'html' directory
html_filenames = [
    "html/employee_age_distribution.html",
    "html/job_title_by_sector.html",
    "html/educational_qualification_by_gender.html",
    "html/nationality_distribution.html",
    "html/marital_status_distribution.html",
    "html/job_grade_by_gender.html",
    "html/department_employee_count.html",
    "html/employee_tenure_distribution.html",
    "html/new_hires_trend.html",
    "html/job_category_employee_number.html"
]

# Save each visualization as an HTML file for interactivity
for i, fig in enumerate(visualizations):
    fig.write_html(html_filenames[i])

print("Interactive HTML files have been saved successfully in the 'html' directory!")
