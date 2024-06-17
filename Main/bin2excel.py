import pandas as pd
from pymavlink import mavutil

def parse_ardupilot_log(file_path):
    # Initialize a connection to the log file
    mlog = mavutil.mavlink_connection(file_path)

    # Dictionary to hold all parsed data
    data = {}

    # Read messages from the log
    while True:
        try:
            msg = mlog.recv_match(blocking=False)
            if msg is None:
                break

            # Skip heartbeat messages to avoid excessive data
            if msg.get_type() == 'HEARTBEAT':
                continue

            # Add message to data dictionary
            msg_dict = msg.to_dict()
            msg_type = msg.get_type()

            if msg_type not in data:
                data[msg_type] = []

            data[msg_type].append(msg_dict)
        except Exception as e:
            print(f"Error: {e}")
            break

    return data

def convert_to_excel(data, output_file):
    # Create a Pandas Excel writer using XlsxWriter as the engine
    writer = pd.ExcelWriter(output_file, engine='openpyxl')

    # Convert each message type to a separate sheet in the Excel file
    for msg_type, msg_list in data.items():
        df = pd.DataFrame(msg_list)
        df.to_excel(writer, sheet_name=msg_type, index=False)

    # Save the Excel file
    writer._save()

def main(input_log, output_excel):
    data = parse_ardupilot_log(input_log)
    convert_to_excel(data, output_excel)
    print(f"Data successfully written to {output_excel}")

if __name__ == "__main__":
    input_log = 'path_to_your_log_file.bin'  
    output_excel = 'output.xlsx'
    main(input_log, output_excel)
