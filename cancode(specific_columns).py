import pandas as pd
import cantools
import os
from datetime import datetime
from tkinter import Tk
from tkinter.filedialog import askopenfilename

# Hide the main Tkinter window
root = Tk()
root.withdraw()

# Show a dialog to select the CSV file
csv_file_path = askopenfilename(title="Select the CSV File", filetypes=[("CSV files", "*.csv")])
if not csv_file_path:
    raise FileNotFoundError("No CSV file selected")

# Show a dialog to select the DBC file
dbc_file_path = askopenfilename(title="Select the DBC File", filetypes=[("DBC files", "*.dbc")])
if not dbc_file_path:
    raise FileNotFoundError("No DBC file selected")

# Load the DBC file
db = cantools.database.load_file(dbc_file_path)

# Read the CSV file and skip the first two rows which seem to contain metadata
df_csv = pd.read_csv(csv_file_path, delimiter=';', skiprows=2)

# Manually rename the columns based on their positions
df_csv.columns = ['Index', 'Timestamp', 'Time', 'Type', 'Frame ID', 'Length', 'Data']

# Check if necessary columns are present
if 'Frame ID' not in df_csv.columns or 'Data' not in df_csv.columns:
    print("The required 'Frame ID' or 'Data' columns are missing in the CSV file.")
    print("Current columns in the CSV file:", df_csv.columns)
    raise KeyError("The required 'Frame ID' or 'Data' columns are missing in the CSV file.")

# Function to decode CAN message using cantools
def decode_can_message(row):
    try:
        message_id = int(row['Frame ID'], 16)
        if message_id in [msg.frame_id for msg in db.messages]:
            message = db.get_message_by_frame_id(message_id)
            data = bytes.fromhex(row['Data'].replace(' ', ''))
            decoded = message.decode(data)

            # Print the decoded message for debugging
            print(f"Decoded message for Frame ID {row['Frame ID']}: {decoded}")

            # Filter for the required fields
            filtered_data = {
                'Battery_current': decoded.get('Battery_current', 'null'),
                'Battery_Voltage': decoded.get('Battery_Voltage', 'null'),
                'Current_Control_status': decoded.get('Current_Control_status', 'null'),
                'temprature': decoded.get('temprature', 'null')
            }
            return filtered_data
        else:
            return {
                'Battery_current': 'null',
                'Battery_Voltage': 'null',
                'Current_Control_status': 'null',
                'temprature': 'null'
            }
    except ValueError as ve:
        print(f"ValueError: {ve}, row: {row}")
        return {
            'Battery_current': 'null',
            'Battery_Voltage': 'null',
            'Current_Control_status': 'null',
            'temprature': 'null'
        }
    except KeyError as ke:
        print(f"KeyError: {ke}, row: {row}")
        return {
            'Battery_current': 'null',
            'Battery_Voltage': 'null',
            'Current_Control_status': 'null',
            'temprature': 'null'
        }
    except Exception as e:
        print(f"Error decoding row {row.get('Frame ID', 'Unknown')} : {e}, row: {row}")
        return {
            'Battery_current': 'null',
            'Battery_Voltage': 'null',
            'Current_Control_status': 'null',
            'temprature': 'null'
        }

# Apply the decoding function to each row in the CSV file
decoded_data = df_csv.apply(lambda row: decode_can_message(row), axis=1)

# Convert the decoded data to a DataFrame and ensure proper alignment by using .fillna()
decoded_df = pd.json_normalize(decoded_data)

# Ensure all columns from both original and decoded data are present
decoded_df = decoded_df.fillna('null')

# Select only the specified columns for the final output
columns_to_include = ['Battery_current', 'Battery_Voltage', 'Current_Control_status', 'temprature']
final_df = decoded_df[columns_to_include]

# Extract the base name of the CSV file
csv_base_name = os.path.basename(csv_file_path).split('.')[0]

# Generate the output file name with the prefix "extractedcan"
output_file_name = f"extractedcan_{csv_base_name}.csv"

# Specify the directory path to save the file
output_directory = r"E:\KONWERT\CAN\Can_extracted_csv"

# Combine the directory path and the file name
output_csv_file_path = os.path.join(output_directory, output_file_name)

# Save the combined data to a new CSV file
final_df.to_csv(output_csv_file_path, index=False)

# Display the combined dataframe
print(final_df.head())

# Confirm the output file path
print(f"Data saved to: {output_csv_file_path}")
