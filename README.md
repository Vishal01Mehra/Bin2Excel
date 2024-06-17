# Bin2Excel
This Python script converts Ardupiot .bin log file into Microsoft excel format.

## Install Required Libraries:
pymavlink: For parsing ArduPilot log files.
pandas: For handling data manipulation.
openpyxl: For writing to Excel files.
You can install these libraries using pip:

## Explanation:
1. parse_ardupilot_log: This function reads the ArduPilot log file using pymavlink and parses the messages into a dictionary. Each message type is stored as a list of dictionaries.
2. convert_to_excel: This function takes the parsed data and writes each message type to a separate sheet in an Excel file using pandas.
3. main: This function orchestrates the process, calling the parse and convert functions and handling file paths.
## Running the Script:
- Replace **'path_to_your_log_file.bin'** with the path to your ArduPilot log file.
- Replace **'output.xlsx'** with the desired path for the output Excel file.
- Run the script.
This script will parse the log file, extract the relevant data, and write it to an Excel file with each message type in a separate sheet.
**Feel free to modify and re-distribute this script with appropriate Credits.**
