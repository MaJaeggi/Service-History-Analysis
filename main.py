#A file to analyse tasks complete by poolwerx Boronia + Lilydale from 30/5/19 until 31/1/24

#libraries
import pandas as pd
import matplotlib.pyplot as plt

#Replace 'your_file.xlsx' with the actual file name
file_path = 'C:/Users/mjaeg/OneDrive/Desktop/Big folder/Coding/PythonTasksAnalysis/ServiceTasksCount.xlsx'

#Read the Excel file into a pandas DataFrame
df = pd.read_excel(file_path)


#Convert the 'Next Service Start Date' column to datetime format if it's not already
df['NextService Start Date'] = pd.to_datetime(df['NextService Start Date'])
#Set the 'Next Service Start Date' column as the index
df.set_index('NextService Start Date', inplace=True)
#Check and handle NaN values in the 'Project Name' column
df['Project Name'].fillna('', inplace=True)
#Add a new column 'Territory' with a default value of 'Other'
df['Territory'] = 'Other' 
#Check and handle NaN values in the 'Address' column
df['Address'].fillna('', inplace=True)


#Update 'Territory' column based on 'Address' substring matches
df.loc[df['Address'].str.contains('Lilydale|Montrose|Kilsyth|Mooroolbark|The Basin|Chirnside Park|Evelyn|Yereing|Chum Creek|Healesville|Woori|Yellingbo|Yarra Junction|Wesburn|Kalorama|Wandin|Coldstream|Gruyere|Tecoma|Belgrave|Olinda|Sassafras|Yarra Glen|Upwey|Seville|Launching|Ferny Creek|Sherbrooke|Silvan|Monbulk|Killara|Hoddles Creek|Don Valley|Warbuton|Macclesfield|Dixon|Steels', case=False, regex=True), 'Territory'] = 'Lilydale'
df.loc[df['Address'].str.contains('Boronia|Ferntree|Bayswater|Bayswater North|Wantirna|Knox City|Mountain Gate|Studfield', case=False, regex=True), 'Territory'] = 'Boronia'
df.loc[df['Address'].str.contains('Rowville|Mulgrave|Clarinda|Wheelers Hill|Hughsdale|Huntingdale|Notting Hill|Oakleigh|Springvale|Clayton|Lysterfield|Scoresby|Knoxfield', case=False, regex=True), 'Territory'] = 'Rowville'
df.loc[df['Address'].str.contains('Berwick|Beaconsfield|Narre Warren|Fountain Gate|Guys Hill|Harkaway|Clyde|Gembrook|Cockatoo', case=False, regex=True), 'Territory'] = 'Berwick'
df.loc[df['Address'].str.contains('Warrandyte|Wonga Park|Donvale|Kangaroo|Guys Hill|Harkaway|Clyde|Gembrook|Cockatoo', case=False, regex=True), 'Territory'] = 'Warrandyte'
#Filter rows where 'status' is 'completed' and dates are before or on 31/01/2024
completed_tasks = df[(df['Status'] == 'Completed') & (df.index <= '2024-01-31')]
#Filter rows where 'status' is 'Bronze', 'Silver', or 'Gold' and dates are before or on 31/01/2024
regular_services = df[df['Project Name'].str.contains('Bronze|Silver|Gold', case=False, regex=True) & (df.index <= '2024-01-31')]
#Filter rows for 'Casual Services' (not meeting the criteria for regular services) and dates before or on 31/01/2024
casual_services_condition = ~df['Project Name'].str.contains('Bronze|Silver|Gold', case=False, regex=True) & (df.index <= '2024-01-31')
casual_services = df[casual_services_condition]

#Group by month and count the number of total services for each month
completed_tasks_count = completed_tasks.resample('M').size().reset_index(name='completed_tasks_count')
#Group by month and count the number of regular services for each month
regular_services_count = regular_services.resample('M').size().reset_index(name='regular_services_count')
#Group by month and count the number of casual services for each month
casual_services_count = casual_services.resample('M').size().reset_index(name='casual_services_count')

#Plot the line graph
plt.figure(figsize=(10, 6))
lines = {}
lines['Total Services'], = plt.plot(completed_tasks_count['NextService Start Date'], completed_tasks_count['completed_tasks_count'], marker='o', linestyle='-', color='b', label='Total Services')
lines['Regular Services'], = plt.plot(regular_services_count['NextService Start Date'], regular_services_count['regular_services_count'], marker='o', linestyle='-', color='orange', label='Regular Services')
lines['Casual Services'], = plt.plot(casual_services_count['NextService Start Date'], casual_services_count['casual_services_count'], marker='o', linestyle='-', color='purple', label='Casual Services')
#Create visibility status dictionary
visibility_status = {label: True for label in lines}

#Create custom legend with line colors
legend = plt.legend(loc='upper left', bbox_to_anchor=(1, 0.5))
legend_lines = legend.get_lines()
for line, label in zip(legend_lines, lines.keys()):
    line.set_picker(True)
    line.set_pickradius(5)
    line.set_visible(visibility_status[label])


#Shade lockdown periods with legend
lockdown_periods = [
    ('2020-03-31', '2020-05-12'),
    ('2020-07-09', '2020-10-27'),
    ('2021-02-13', '2021-02-17'),
    ('2021-05-28', '2021-06-10'),
    ('2021-07-16', '2021-07-27'),
    ('2021-08-05', '2021-10-21'),
]

#Set a single label for all lockdown periods
lockdown_label = 'Lockdown'

#Plot shaded areas without label
for i, (start_date, end_date) in enumerate(lockdown_periods):
    if i == 0:  # Add label only for the first shaded area
        plt.axvspan(pd.to_datetime(start_date), pd.to_datetime(end_date), facecolor='red', alpha=0.2, label=lockdown_label)
    else:
        plt.axvspan(pd.to_datetime(start_date), pd.to_datetime(end_date), facecolor='red', alpha=0.2)

#Add vertical line at Lilydale opening date
lilydale_date = pd.to_datetime('2023-03-18')
plt.axvline(lilydale_date, color='green', linestyle='--', label='Lilydale Opens')

#Set x-axis limits to start at the same point as the first x-value
plt.xlim(completed_tasks_count['NextService Start Date'].min(), completed_tasks_count['NextService Start Date'].max())


#Create legend with line colurs
legend = plt.legend(loc='upper left')
legend_lines = legend.get_lines()


#Graph Vitals for Figure #1
plt.title('Services Per Month: Jun 2019 - Jan 2024')
plt.xlabel('Date')
plt.ylabel('Count')
plt.xticks(rotation=45, ha='right')

#Show the plot
plt.tight_layout()
plt.show()


