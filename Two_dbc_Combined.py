import pandas as pd
import cantools
import os
from tkinter import Tk
from tkinter.filedialog import askopenfilename

# Hide the main Tkinter window
root = Tk()
root.withdraw()

# Show a dialog to select the CSV file
csv_file_path = askopenfilename(title="Select the CSV File", filetypes=[("CSV files", "*.csv")])
if not csv_file_path:
    raise FileNotFoundError("No CSV file selected")

# Show a dialog to select the first DBC file
dbc_file_path_1 = askopenfilename(title="Select the First DBC File", filetypes=[("DBC files", "*.dbc")])
if not dbc_file_path_1:
    raise FileNotFoundError("No DBC file selected")

# Show a dialog to select the second DBC file
dbc_file_path_2 = askopenfilename(title="Select the Second DBC File", filetypes=[("DBC files", "*.dbc")])
if not dbc_file_path_2:
    raise FileNotFoundError("No DBC file selected")

# Load the DBC files
db1 = cantools.database.load_file(dbc_file_path_1)
db2 = cantools.database.load_file(dbc_file_path_2)

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
def decode_can_message(row, db):
    try:
        message_id = int(row['Frame ID'], 16)
        if message_id in [msg.frame_id for msg in db.messages]:
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

# Apply the decoding function to each row in the CSV file using both DBC files
decoded_data_1 = df_csv.apply(lambda row: decode_can_message(row, db1), axis=1)
decoded_data_2 = df_csv.apply(lambda row: decode_can_message(row, db2), axis=1)

# Convert the decoded data to DataFrames and ensure proper alignment by using .fillna()
decoded_df_1 = pd.json_normalize(decoded_data_1).fillna('null')
decoded_df_2 = pd.json_normalize(decoded_data_2).fillna('null')

# Combine the original CSV data with the decoded data from both DBC files, aligning columns correctly
df_combined = pd.concat([df_csv.reset_index(drop=True), decoded_df_1.reset_index(drop=True), decoded_df_2.reset_index(drop=True)], axis=1)

# Ensure all columns from both original and decoded data are present
df_combined = df_combined.fillna('null')

# Extract the base name of the CSV file
csv_base_name = os.path.basename(csv_file_path).split('.')[0]

# Generate the output file name with the prefix "extractedcan"
output_file_name = f"extractedcan_{csv_base_name}.xlsx"

# Specify the directory path to save the file
output_directory = r"E:\KONWERT\CAN\Can_extracted_csv"

# Combine the directory path and the file name
output_excel_file_path = os.path.join(output_directory, output_file_name)

# Save the combined data to a new Excel file
df_combined.to_excel(output_excel_file_path, index=False)

# Display the combined dataframe
print(df_combined.head())

# Confirm the output file path
print(f"Data saved to: {output_excel_file_path}")
