# Overview
An ETL Powershell script that transforms and consolidates sales csv files into one final excel report.

# More detailed description
The script can be scheduled to run at custom time of the day and scans a custom directory for new csv sales files.
If it finds ones, it extracts the data, unifies datetime format to match the US one and also does a currency conversion to USD.
The result is a final excel workbook, combining everything

In addition, the script can accept the optional parameters: csv directory, time and exchange rates, while also logging every event along the way.
