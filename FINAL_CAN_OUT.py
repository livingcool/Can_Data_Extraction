import pandas as pd
import cantools
import os
from datetime import datetime

# Load the DBC file
dbc_file_path = "E:\\KONWERT\\CAN_DBC_FILES\\DBC File for candata\\SEG_Standard_DBC_02.06.23.dbc"
db = cantools.database.load_file(dbc_file_path)

# Read the CSV file and skip the first two rows which seem to contain metadata
csv_file_path = "E:\\KONWERT\\CAN\\candatacsv\\trail3.csv"
df_csv = pd.read_csv(csv_file_path, delimiter=';', skiprows=2)

# Print the actual column names to ensure correct column names
print("Columns in CSV file:", df_csv.columns)

# Rename columns to match expected names
df_csv = df_csv.rename(columns={
    'Id': 'Frame ID',        # Rename 'Id' to 'Frame ID'
    'Data': 'Data'           # The 'Data' column seems to be correctly named
})

# Print the actual column names after renaming
print("Columns after renaming:", df_csv.columns)

# Function to decode CAN message using cantools
def decode_can_message(row):
    try:
        message_id = int(row['Frame ID'], 16)  # Adjust to match the renamed column
        if message_id in [msg.frame_id for msg in db.messages]:
            message = db.get_message_by_frame_id(message_id)
            data = bytes.fromhex(row['Data'].replace(' ', ''))  # Adjust to match the renamed column
            decoded = message.decode(data)
            return decoded
        else:
            return {}
    except ValueError as ve:
        print(f"ValueError: {ve}, row: {row}")
        return {}
    except KeyError:
        print(f"KeyError: 'Frame ID' or 'Data' column is missing in row: {row}")
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

# Generate a unique file name using the current timestamp
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
unique_file_name = f"gen_can_data_{timestamp}.csv"
output_csv_file_path = os.path.join("E:\\KONWERT\\CAN\\Can_extracted_csv", unique_file_name)

# Save the combined data to a new CSV file
df_combined.to_csv(output_csv_file_path, index=False)

# Display the combined dataframe
print(df_combined.head())

# Confirm the output file path
print(f"Data saved to: {output_csv_file_path}")
