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

# Extract frame IDs and signal names from the DBC file
dbc_frame_ids = [msg.frame_id for msg in db.messages]
print("Frame IDs in DBC file:", dbc_frame_ids)

# Print signal names for each message in the DBC file
for msg in db.messages:
    print(f"Message: {msg.name} (ID: {msg.frame_id})")
    for signal in msg.signals:
        print(f"  Signal: {signal.name}")

# Read the CSV file
df_csv = pd.read_csv(csv_file_path, skiprows=2, delimiter=';')

# Print the actual column names
print("Columns in CSV file:", df_csv.columns)

# Rename columns to match expected names
df_csv.columns = ['Nr', 'Timestamp', 'Time', 'Type', 'Frame ID', 'Length', 'Data']

print("Columns after renaming:", df_csv.columns)

# Function to decode CAN message using cantools
def decode_can_message(row):
    try:
        message_id = int(row['Frame ID'], 16)
        if message_id in dbc_frame_ids:
            message = db.get_message_by_frame_id(message_id)
            data = bytes.fromhex(row['Data'].replace(' ', ''))
            decoded = message.decode(data)
            return decoded
        else:
            return {}
    except ValueError as ve:
        print(f"ValueError: {ve}, row: {row}")
        return {}
    except KeyError as ke:
        print(f"KeyError: {ke}, row: {row}")
        return {}
    except Exception as e:
        print(f"Error decoding row {row.get('Frame ID', 'Unknown')} : {e}, row: {row}")
        return {}

# Apply the decoding function to each row in the CSV file
decoded_data = df_csv.apply(lambda row: decode_can_message(row), axis=1)

# Convert the decoded data to a DataFrame
decoded_df = pd.json_normalize(decoded_data)

# Combine the original CSV data with the decoded data
df_combined = pd.concat([df_csv, decoded_df], axis=1)

# Define the output directory and file name
output_dir = "E:\\KONWERT\\CAN\\Can_extracted_csv"
os.makedirs(output_dir, exist_ok=True)

# Check if the number of rows exceeds the maximum Excel limit
max_rows_per_sheet = 1048576
output_excel_file_path = os.path.join(output_dir, f"decoded_can_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")

# Write to Excel with multiple sheets if necessary
with pd.ExcelWriter(output_excel_file_path, engine='xlsxwriter') as writer:
    if len(df_combined) > max_rows_per_sheet:
        num_parts = (len(df_combined) // max_rows_per_sheet) + 1
        for i in range(num_parts):
            start_row = i * max_rows_per_sheet
            end_row = start_row + max_rows_per_sheet
            part_df = df_combined.iloc[start_row:end_row]
            part_df.to_excel(writer, sheet_name=f'Part_{i+1}', index=False)
            print(f"Saved part {i+1} to sheet Part_{i+1}")
    else:
        df_combined.to_excel(writer, index=False)
        print(f"Decoded CAN data saved to {output_excel_file_path}")

print(f"Decoded CAN data saved to {output_excel_file_path}")
