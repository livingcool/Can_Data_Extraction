import pandas as pd
import cantools
import os
from tkinter import Tk, filedialog

# Define the corrected data where each list has the same length
data = {
    '102200A1': ['MC_MOTOR_SPEED', 'MC_STATUS_REGEN', 'MC_STATUS_REVERSE', 'MC_STATUS_FWD', 'MC_STATUS_BRK'],
    '102200A3': ['MC_STATUS_BRK', '', '', '', ''],
    '102200A0': ['MC_PH_CURR', '', 'MC_MOTOR_TEMP', '', ''],
    '19FF01D8': ['Battery_Voltage', 'Battery_Current', '', '', ''],
    '19FF01D9': ['State_of_Charge', 'State_of_Health', 'Availablecapacity', '', ''],
    '19FF01DA': ['Fault', 'Warning', '', '', '']
}

# Create a DataFrame from the corrected data
df_static = pd.DataFrame(data)

# Remove empty strings from the DataFrame
df_static = df_static.apply(lambda x: x.mask(x == '').fillna('null'))

# Hide the main Tkinter window
root = Tk()
root.withdraw()

# Show a dialog to select multiple CSV files
csv_file_paths = filedialog.askopenfilenames(title="Select CSV Files", filetypes=[("CSV files", "*.csv")])
if not csv_file_paths:
    raise FileNotFoundError("No CSV files selected")

# Show a dialog to select the first DBC file
dbc_file_path_1 = filedialog.askopenfilename(title="Select the First DBC File", filetypes=[("DBC files", "*.dbc")])
if not dbc_file_path_1:
    raise FileNotFoundError("No DBC file selected")

# Show a dialog to select the second DBC file
dbc_file_path_2 = filedialog.askopenfilename(title="Select the Second DBC File", filetypes=[("DBC files", "*.dbc")])
if not dbc_file_path_2:
    raise FileNotFoundError("No DBC file selected")

# Load the DBC files
db1 = cantools.database.load_file(dbc_file_path_1)
db2 = cantools.database.load_file(dbc_file_path_2)

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

# Function to extract a representative time from each second with data frames
def extract_representative_time(df):
    df['Time'] = pd.to_datetime(df['Time'], format='%H:%M:%S.%f')
    df['second'] = df['Time'].dt.floor('S')
    times_per_second = df.groupby('second')['Time'].first().reset_index(drop=True)
    times_per_second = times_per_second.dt.strftime('%H:%M:%S.%f').str[:-3]  # Format as HH:MM:SS.sss
    return times_per_second

# Aggregate decoded data into a single dictionary for each row
def aggregate_decoded_data(row):
    decoded_data_1 = decode_can_message(row, db1)
    decoded_data_2 = decode_can_message(row, db2)
    return decoded_data_1, decoded_data_2

# Function to calculate average values per second
def calculate_average_values(df, column_prefix):
    df_numeric = df.apply(pd.to_numeric, errors='coerce')
    df_resampled = df_numeric.groupby(df.index // 1000).agg({
        col: 'mean' if col not in ['Battery_Current', 'Battery_Voltage'] else 'max' for col in df_numeric.columns
    })  # Assuming each second has 1000 records
    df_resampled.columns = [f"{column_prefix}_{col}" for col in df_resampled.columns]
    return df_resampled

# Process each selected CSV file
for csv_file_path in csv_file_paths:
    # Read the CSV file and skip the first two rows which seem to contain metadata
    df_csv = pd.read_csv(csv_file_path, delimiter=';', skiprows=2)

    # Manually rename the columns based on their positions
    df_csv.columns = ['Index', 'Timestamp', 'Time', 'Type', 'Frame ID', 'Length', 'Data']

    # Check if necessary columns are present
    if 'Frame ID' not in df_csv.columns or 'Data' not in df_csv.columns:
        print("The required 'Frame ID' or 'Data' columns are missing in the CSV file.")
        print("Current columns in the CSV file:", df_csv.columns)
        raise KeyError("The required 'Frame ID' or 'Data' columns are missing in the CSV file.")

    # Apply the aggregation function to each row in the CSV file
    df_csv[['Aggregated Data_1', 'Aggregated Data_2']] = df_csv.apply(lambda x: pd.Series(aggregate_decoded_data(x)), axis=1)

    # Calculate average values per second for each DBC file's aggregated data
    df_avg_1 = calculate_average_values(pd.json_normalize(df_csv['Aggregated Data_1']), 'dbc1')
    df_avg_2 = calculate_average_values(pd.json_normalize(df_csv['Aggregated Data_2']), 'dbc2')

    # Combine the average dataframes into a single row
    df_combined_avg = pd.concat([df_avg_1, df_avg_2], axis=1).fillna('null')

    # Ensure the length of times_per_second matches df_combined_avg
    times_per_second = extract_representative_time(df_csv)
    times_per_second = times_per_second[:len(df_combined_avg)]

    # Assign the representative time to each second's aggregated data
    df_combined_avg['Time'] = times_per_second.values

    # Function to calculate additional columns
    def calculate_additional_columns(df):
        df['dbc1_MC_PH_CURR'] = pd.to_numeric(df['dbc1_MC_PH_CURR'], errors='coerce')
        df['dbc1_MC_MOTOR_SPEED'] = pd.to_numeric(df['dbc1_MC_MOTOR_SPEED'], errors='coerce')
        df.loc[:, 'motor_current'] = df['dbc1_MC_PH_CURR'] * 0.866  # Convert phase current to DC current
        df.loc[:, 'vehicle_speed'] = df['dbc1_MC_MOTOR_SPEED'] * 0.012551909  # Convert motor speed to vehicle speed
        return df

    # Apply the additional column calculations
    df_combined_avg = calculate_additional_columns(df_combined_avg)

    # Merge the static and dynamic dataframes while ensuring all columns are aligned
    df_combined_final = pd.concat([df_combined_avg, df_static], axis=1)

    # Define columns to keep
    columns_to_keep = [
        'dbc1_MC_MOTOR_SPEED', 'dbc1_MC_PH_CURR', 'motor_current', 'vehicle_speed',
        'dbc1_MC_STATUS_REGEN', 'dbc1_MC_STATUS_REVERSE', 'dbc1_MC_STATUS_FWD', 'dbc1_MC_STATUS_BRK',
        'dbc1_MC_MOTOR_TEMP', 'dbc1_MC_DC_VOLT',
        'dbc2_Battery_Voltage', 'dbc2_Battery_Current',
        'dbc2_State_of_Charge', 'dbc2_State_of_Health', 'dbc2_Availablecapacity', 'dbc2_Fault', 'dbc2_Warning',
        'Time'
    ]

    # Ensure all columns are aligned and reorder
    df_combined_final = df_combined_final.reindex(columns=columns_to_keep, fill_value='null')

    # Extract the base name of the CSV file
    csv_base_name = os.path.basename(csv_file_path).split('.')[0]

    # Generate the output file name with the prefix "extractedcan"
    output_file_name = f"extractedcan_{csv_base_name}.xlsx"

    # Specify the directory path to save the file
    output_directory = r"E:\KONWERT\Can_extracted_csv"

    # Combine the directory path and the file name
    output_excel_file_path = os.path.join(output_directory, output_file_name)

    # Save the selected data to a new Excel file
    df_combined_final.to_excel(output_excel_file_path, index=False)

    # Display the combined dataframe
    print(f"Data for {csv_base_name} saved to: {output_excel_file_path}")
