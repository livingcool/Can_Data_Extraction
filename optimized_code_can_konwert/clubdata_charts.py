import math
import pandas as pd
import os
from tkinter import Tk, filedialog
import matplotlib.pyplot as plt

# Hide the main Tkinter window
root = Tk()
eco_mode = 0.73
boost_mode = 0.92

root.withdraw()

# Show a dialog to select the Excel files
excel_file_paths = filedialog.askopenfilenames(title="Select Excel Files", filetypes=[("Excel files", "*.xlsx")])
if not excel_file_paths:
    raise FileNotFoundError("No Excel files selected")

# Initialize an empty list to store DataFrames from each Excel file
dfs = []

# Read each Excel file into a DataFrame and append to dfs list
for excel_file_path in excel_file_paths:
    try:
        df = pd.read_excel(excel_file_path)
        # Convert all columns to numeric (if possible) to ensure numerical values
        df = df.apply(pd.to_numeric, errors='ignore')
        dfs.append(df)
    except Exception as e:
        print(f"Error processing {excel_file_path}: {e}")

# Check if any DataFrames were loaded
if not dfs:
    raise ValueError("No valid data found in selected Excel files.")

# Concatenate all DataFrames based on 'Time' column if it exists
df_combined = pd.concat(dfs, ignore_index=True)

# Sort combined DataFrame by 'Time' column if it exists
if 'Time' in df_combined.columns:
    df_combined.sort_values(by='Time', inplace=True)
else:
    print("Warning: 'Time' column not found. Sorting may not be performed.")

# Calculate motor power with efficiency of 0.80
efficiency = eco_mode
if 'motor_current' in df_combined.columns and 'dbc1_MC_DC_VOLT' in df_combined.columns and 'dbc1_MC_MOTOR_SPEED' in df_combined.columns:
    # Calculate power in watts
    df_combined['power'] = df_combined['motor_current'] * df_combined['dbc1_MC_DC_VOLT'] * eco_mode
    
    # Calculate torque in Nm (Newton-meters)
    df_combined['torque'] = df_combined['power'] / (2 * math.pi * df_combined['dbc1_MC_MOTOR_SPEED'] / 60)
    
    # Filter out torque values above 35 Nm
    df_combined = df_combined[df_combined['torque'] <= 35]
    
else:
    print("Warning: Columns 'motor_current', 'dbc1_MC_DC_VOLT', or 'dbc1_MC_MOTOR_SPEED' not found. Power and torque calculation skipped.")

# Calculate battery power
if 'dbc2_Battery_Current' in df_combined.columns and 'dbc2_Battery_Voltage' in df_combined.columns:
    df_combined['battery_power'] = df_combined['dbc2_Battery_Current'] * df_combined['dbc2_Battery_Voltage']
else:
    print("Warning: Columns 'dbc2_Battery_Current' or 'dbc2_Battery_Voltage' not found. Battery power calculation skipped.")

# Plotting various parameters
plt.figure(figsize=(18, 12))

# RPM plot
plt.subplot(3, 3, 1)
plt.plot(df_combined['Time'], df_combined['dbc1_MC_MOTOR_SPEED'], color='blue')
plt.title('RPM Over Time')
plt.xlabel('Time')
plt.ylabel('RPM')
plt.grid(True)
plt.gca().ticklabel_format(axis='y', style='plain')

# Annotations for RPM plot
max_rpm_time = df_combined.loc[df_combined['dbc1_MC_MOTOR_SPEED'].idxmax(), 'Time']
max_rpm_value = df_combined['dbc1_MC_MOTOR_SPEED'].max()
plt.annotate(f'Max RPM\n{max_rpm_value:.2f}', xy=(max_rpm_time, max_rpm_value), xytext=(-20, 20), textcoords='offset points',
             arrowprops=dict(arrowstyle='->', color='black'))

min_rpm_time = df_combined.loc[df_combined['dbc1_MC_MOTOR_SPEED'].idxmin(), 'Time']
min_rpm_value = df_combined['dbc1_MC_MOTOR_SPEED'].min()
plt.annotate(f'Min RPM\n{min_rpm_value:.2f}', xy=(min_rpm_time, min_rpm_value), xytext=(-20, -30), textcoords='offset points',
             arrowprops=dict(arrowstyle='->', color='black'))

# Current plot
plt.subplot(3, 3, 2)
plt.plot(df_combined['Time'], df_combined['motor_current'], color='green')
plt.title('Motor Current Over Time')
plt.xlabel('Time')
plt.ylabel('Current (A)')
plt.grid(True)
plt.gca().ticklabel_format(axis='y', style='plain')

# Annotations for Current plot
max_current_time = df_combined.loc[df_combined['motor_current'].idxmax(), 'Time']
max_current_value = df_combined['motor_current'].max()
plt.annotate(f'Max Current\n{max_current_value:.2f}', xy=(max_current_time, max_current_value), xytext=(-20, 20), textcoords='offset points',
             arrowprops=dict(arrowstyle='->', color='black'))

min_current_time = df_combined.loc[df_combined['motor_current'].idxmin(), 'Time']
min_current_value = df_combined['motor_current'].min()
plt.annotate(f'Min Current\n{min_current_value:.2f}', xy=(min_current_time, min_current_value), xytext=(-20, -30), textcoords='offset points',
             arrowprops=dict(arrowstyle='->', color='black'))

# Voltage plot
plt.subplot(3, 3, 3)
plt.plot(df_combined['Time'], df_combined['dbc1_MC_DC_VOLT'], color='red')
plt.title('DC Voltage Over Time')
plt.xlabel('Time')
plt.ylabel('Voltage (V)')
plt.grid(True)
plt.gca().ticklabel_format(axis='y', style='plain')

# Annotations for Voltage plot
max_voltage_time = df_combined.loc[df_combined['dbc1_MC_DC_VOLT'].idxmax(), 'Time']
max_voltage_value = df_combined['dbc1_MC_DC_VOLT'].max()
plt.annotate(f'Max Voltage\n{max_voltage_value:.2f}', xy=(max_voltage_time, max_voltage_value), xytext=(-20, 20), textcoords='offset points',
             arrowprops=dict(arrowstyle='->', color='black'))

min_voltage_time = df_combined.loc[df_combined['dbc1_MC_DC_VOLT'].idxmin(), 'Time']
min_voltage_value = df_combined['dbc1_MC_DC_VOLT'].min()
plt.annotate(f'Min Voltage\n{min_voltage_value:.2f}', xy=(min_voltage_time, min_voltage_value), xytext=(-20, -30), textcoords='offset points',
             arrowprops=dict(arrowstyle='->', color='black'))

# SOC vs Battery Capacity plot
plt.subplot(3, 3, 4)
plt.plot(df_combined['dbc2_Availablecapacity'], df_combined['dbc2_State_of_Charge'], color='purple')
plt.title('SOC vs Battery Capacity')
plt.xlabel('Battery Capacity')
plt.ylabel('SOC (%)')
plt.grid(True)

# Annotations for SOC vs Battery Capacity plot
max_soc_capacity = df_combined.loc[df_combined['dbc2_State_of_Charge'].idxmax(), 'dbc2_Availablecapacity']
max_soc_value = df_combined['dbc2_State_of_Charge'].max()
plt.annotate(f'Max SOC\n{max_soc_value:.2f}', xy=(max_soc_capacity, max_soc_value), xytext=(-20, 20), textcoords='offset points',
             arrowprops=dict(arrowstyle='->', color='black'))

min_soc_capacity = df_combined.loc[df_combined['dbc2_State_of_Charge'].idxmin(), 'dbc2_Availablecapacity']
min_soc_value = df_combined['dbc2_State_of_Charge'].min()
plt.annotate(f'Min SOC\n{min_soc_value:.2f}', xy=(min_soc_capacity, min_soc_value), xytext=(-20, -30), textcoords='offset points',
             arrowprops=dict(arrowstyle='->', color='black'))

# Current vs Voltage plot
plt.subplot(3, 3, 5)
plt.plot(df_combined['dbc2_Battery_Voltage'], df_combined['dbc2_Battery_Current'], color='blue')
plt.title('Current vs Voltage')
plt.xlabel('Voltage (V)')
plt.ylabel('Current (A)')
plt.grid(True)
plt.gca().ticklabel_format(axis='y', style='plain')

# Annotations for Current vs Voltage plot
max_voltage_current = df_combined.loc[df_combined['dbc2_Battery_Current'].idxmax(), 'dbc2_Battery_Voltage']
max_current_value = df_combined['dbc2_Battery_Current'].max()
plt.annotate(f'Max Current\n{max_current_value:.2f}', xy=(max_voltage_current, max_current_value), xytext=(-20, 20), textcoords='offset points',
             arrowprops=dict(arrowstyle='->', color='black'))

min_voltage_current = df_combined.loc[df_combined['dbc2_Battery_Current'].idxmin(), 'dbc2_Battery_Voltage']
min_current_value = df_combined['dbc2_Battery_Current'].min()
plt.annotate(f'Min Current\n{min_current_value:.2f}', xy=(min_voltage_current, min_current_value), xytext=(-20, -30), textcoords='offset points',
             arrowprops=dict(arrowstyle='->', color='black'))

# Motor Power vs Battery Power plot
plt.subplot(3, 3, 6)
plt.plot(df_combined['power'], df_combined['battery_power'], color='green')
plt.title('Motor Power vs Battery Power')
plt.xlabel('Motor Power (W)')
plt.ylabel('Battery Power (W)')
plt.grid(True)
plt.gca().ticklabel_format(axis='y', style='plain')

# Annotations for Motor Power vs Battery Power plot
max_motor_power = df_combined.loc[df_combined['power'].idxmax(), 'power']
max_battery_power = df_combined.loc[df_combined['power'].idxmax(), 'battery_power']
plt.annotate(f'Max Power\n{max_motor_power:.2f}', xy=(max_motor_power, max_battery_power), xytext=(-20, 20), textcoords='offset points',
             arrowprops=dict(arrowstyle='->', color='black'))

min_motor_power = df_combined.loc[df_combined['power'].idxmin(), 'power']
min_battery_power = df_combined.loc[df_combined['power'].idxmin(), 'battery_power']
plt.annotate(f'Min Power\n{min_motor_power:.2f}', xy=(min_motor_power, min_battery_power), xytext=(-20, -30), textcoords='offset points',
             arrowprops=dict(arrowstyle='->', color='black'))

# Torque vs Current plot
plt.subplot(3, 3, 7)
plt.plot(df_combined['motor_current'], df_combined['torque'], color='red')
plt.title('Torque vs Current')
plt.xlabel('Current (A)')
plt.ylabel('Torque (Nm)')
plt.grid(True)
plt.gca().ticklabel_format(axis='y', style='plain')

# Annotations for Torque vs Current plot
max_torque_current = df_combined.loc[df_combined['torque'].idxmax(), 'motor_current']
max_torque_value = df_combined['torque'].max()
plt.annotate(f'Max Torque\n{max_torque_value:.2f}', xy=(max_torque_current, max_torque_value), xytext=(-20, 20), textcoords='offset points',
             arrowprops=dict(arrowstyle='->', color='black'))

min_torque_current = df_combined.loc[df_combined['torque'].idxmin(), 'motor_current']
min_torque_value = df_combined['torque'].min()
plt.annotate(f'Min Torque\n{min_torque_value:.2f}', xy=(min_torque_current, min_torque_value), xytext=(-20, -30), textcoords='offset points',
             arrowprops=dict(arrowstyle='->', color='black'))

# RPM vs Motor Current plot
plt.subplot(3, 3, 8)
plt.plot(df_combined['dbc1_MC_MOTOR_SPEED'], df_combined['motor_current'], color='purple')
plt.title('RPM vs Motor Current')
plt.xlabel('RPM')
plt.ylabel('Current (A)')
plt.grid(True)
plt.gca().ticklabel_format(axis='y', style='plain')

# Annotations for RPM vs Motor Current plot
max_rpm_current = df_combined.loc[df_combined['motor_current'].idxmax(), 'dbc1_MC_MOTOR_SPEED']
max_current_value = df_combined['motor_current'].max()
plt.annotate(f'Max Current\n{max_current_value:.2f}', xy=(max_rpm_current, max_current_value), xytext=(-20, 20), textcoords='offset points',
             arrowprops=dict(arrowstyle='->', color='black'))

min_rpm_current = df_combined.loc[df_combined['motor_current'].idxmin(), 'dbc1_MC_MOTOR_SPEED']
min_current_value = df_combined['motor_current'].min()
plt.annotate(f'Min Current\n{min_current_value:.2f}', xy=(min_rpm_current, min_current_value), xytext=(-20, -30), textcoords='offset points',
             arrowprops=dict(arrowstyle='->', color='black'))

# Power vs RPM plot
plt.subplot(3, 3, 9)
plt.plot(df_combined['dbc1_MC_MOTOR_SPEED'], df_combined['power'], color='orange')
plt.title('Power vs RPM')
plt.xlabel('RPM')
plt.ylabel('Power (W)')
plt.grid(True)
plt.gca().ticklabel_format(axis='y', style='plain')

# Annotations for Power vs RPM plot
max_power_rpm = df_combined.loc[df_combined['power'].idxmax(), 'dbc1_MC_MOTOR_SPEED']
max_power_value = df_combined['power'].max()
plt.annotate(f'Max Power\n{max_power_value:.2f}', xy=(max_power_rpm, max_power_value), xytext=(-20, 20), textcoords='offset points',
             arrowprops=dict(arrowstyle='->', color='black'))

min_power_rpm = df_combined.loc[df_combined['power'].idxmin(), 'dbc1_MC_MOTOR_SPEED']
min_power_value = df_combined['power'].min()
plt.annotate(f'Min Power\n{min_power_value:.2f}', xy=(min_power_rpm, min_power_value), xytext=(-20, -30), textcoords='offset points',
             arrowprops=dict(arrowstyle='->', color='black'))

# Adjust layout and display plots
plt.tight_layout()
plt.show()

# Specify the directory path to save the combined file
output_directory = r"E:\KONWERT\Can_combined_excel"

# Ensure the output directory exists or create it if not
os.makedirs(output_directory, exist_ok=True)

# Generate the output file name
output_file_name = "combined_data_with_power_torque_battery.xlsx"
output_excel_file_path = os.path.join(output_directory, output_file_name)

# Save the combined data to a new Excel file
df_combined.to_excel(output_excel_file_path, index=False)

# Display the saved file path
print(f"Combined data with power, torque, and battery power saved to: {output_excel_file_path}")

print("Processing completed.")
