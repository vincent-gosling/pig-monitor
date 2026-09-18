# Pig Monitor Software 

A custom Python desktop application developed to automate monthly Key Performance Indicator (KPI) performance reporting for commercial pig farming. The system replaces manual record-keeping by ingesting raw herd data as a user input, generating fully styled multi-tab Excel reports in under one second, and displaying trend visualizations through a custom Tkinter dashboard.

## Full Project Summary

Data Pipeline: Engineered an end-to-end ETL pipeline using Pandas, Os and XlsxWriter. The user fills out a customisable excel input sheet, when the sheet is processed this data is used to fill out a new monthly data column and stored in a .txt file. The data is used to calculate KPIs and financial data then this is collated as an excel workbook using Xlsxwriter which can be opened from the main dashboard.

User Interface: Designed a GUI using Tkinter, including functions to create new reports, a treeview to view previous reports and input sheets, KPI metric cards and customisable plot visuals built straight onto the dashboard.

![Alt Text](https://github.com/vincent-gosling/pig-monitor/blob/main/1786970987741.jpg?raw=true)

NOTE: All data displayed in this image has been modified to maintain confidentiality, for the same reason there are no data files available in the repository.

Desktop Access: currently the project only runs the raw python code and uses a .bat file placed in the users desktop to allow a non-technical user to run the application.

## Repository Structure

├── Pig_Monitor.py               # Data processing and report writing script  
├── Pig_Monitor_Dashboard.py     # Tkinter dashboard, visualisations and data input script  
├── LICENSE                      # License file  
├── .gitignore                   # Git ignore configuration  
└── README.md                    # Project documentation  


## Future changes

In the future I would like to replace the .txt files with .csv's to improve the data readability, work on improving the insights that are available to the user through the dashboards plots, and build a more long term solution for opening the software in the users desktop.

## Contact

Vincent Gosling

LinkedIn: [LinkedIn](https://www.linkedin.com/in/vincent-gosling-58b205370/)
Email: vincent.gosling314@gmail.com
