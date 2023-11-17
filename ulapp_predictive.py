import pandas as pd
import matplotlib.pyplot as plt

# Read the Excel file into a DataFrame
excel_file_path = r"D:\Hist_data\Mabalacat City_5_day_forecast.xlsx"

# Reading the sheets in the excel file into dictionaries
dfs = pd.read_excel(excel_file_path, sheet_name=None)

# Concatenate all DataFrames into a single DataFrame
df = pd.concat(dfs.values(), ignore_index=True)

# Convert the 'Date & Time' column to datetime format
df['Date & Time'] = pd.to_datetime(df['Date & Time'])

# Set the 'Date & Time' column as the index
df.set_index('Date & Time', inplace=True)

# Plotting historical data for Temperature, Humidity, Wind Speed, and Rain (3h)
fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(10, 8), sharex=True)

# Plot Temperature and Humidity
ax1.plot(df.index, df['Temperature (°C)'], label='Temperature (°C)', marker='o')
ax1.plot(df.index, df['Humidity (%)'], label='Humidity (%)', marker='o')
ax1.set_ylabel('Temperature (°C) / Humidity (%)')
ax1.legend()

# Plot Wind Speed and Rain (3h)
ax2.plot(df.index, df['Wind Speed (m/s)'], label='Wind Speed (m/s)', marker='o')
ax2.plot(df.index, df['Rain (3h) (mm)'], label='Rain (3h) (mm)', marker='o')
ax2.set_xlabel('Date & Time')
ax2.set_ylabel('Wind Speed (m/s) / Rain (3h) (mm)')
ax2.legend()

# Plot names
plt.suptitle('Predictive Weather Analysis')
plt.xticks(rotation=45)
plt.tight_layout(rect=[0, 0.03, 1, 0.95])

# Output plot
plt.show()
