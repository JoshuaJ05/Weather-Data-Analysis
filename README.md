# Weather-Data-Analysis

Overview
    This program analyzes weather data from an Excel file, providing insightful visualizations and statistics based on user input. The program validates the user's Excel file, processes the data, and generates both an average monthly temperature analysis and a precipitation analysis. It outputs the results in a new Excel workbook, complete with bar chart visualizations for clarity.


Features
Excel File Integration
Validates and loads user-provided Excel files.
Allows users to analyze data stored in a compatible Excel sheet.

  Temperature Analysis
Calculates average monthly temperatures.
Identifies the month with the highest average temperature.
Generates a bar chart for average monthly temperatures.

  Precipitation Analysis
Analyzes daily precipitation data.
Identifies the day with the highest precipitation.
Creates a bar chart for daily precipitation levels.

'''Future versions may allow users to customize chart colors and styling.'''


How It Works
  Load Excel File
Users provide the file name of their weather data Excel file. The program checks if the file exists and loads it for processing.

  Analyze Data
Temperature Analysis: The program calculates average temperatures for each month and identifies the month with the highest average temperature.
Precipitation Analysis: The program analyzes daily precipitation values and identifies the day with the highest rainfall.

  Generate Output
Results are saved in a new Excel file, which includes:
Temperature Analysis Sheet with data and a bar chart.
Precipitation Analysis Sheet with data and a bar chart.

  Output File
The processed results are saved/appended to the Excel file.

  Dependencies
The program requires the following Python libraries:

openpyxl: For reading, writing, and manipulating Excel files.
datetime: For processing date-related data.

  How to Use
Ensure you have Python installed.
Install the openpyxl library:
pip install openpyxl
Place the Excel file to be analyzed in the same directory as the program.
Run the program and provide the name of the Excel file when prompted.
Check the output file (WeatherDataAnalysis.xlsx) for the analysis results and visualizations.
