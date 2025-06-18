import serial
import pandas as pd
from datetime import datetime
import re
import os
import csv

# --- Configuration ---
COM_PORT = 'COM22'         # Replace with your actual port
BAUD_RATE = 1200
TIMEOUT = 3                # Timeout in seconds

# --- Generate file names ---
timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
CSV_FILE = f'shimadzu_log_{timestamp}.csv'
EXCEL_FILE = f'shimadzu_log_{timestamp}.xlsx'

# --- Initialize containers ---
metadata = {}
header_written = False

# --- Connect to serial port ---
try:
    ser = serial.Serial(COM_PORT, BAUD_RATE, timeout=TIMEOUT)
    print(f"[✓] Connected to {COM_PORT} at {BAUD_RATE} baud.")
except Exception as e:
    print(f"[✗] Failed to open port: {e}")
    exit()

# --- Open CSV and start logging ---
with open(CSV_FILE, 'w', newline='') as csvfile:
    writer = csv.writer(csvfile)
    writer.writerow(['Timestamp', 'Elapsed Time', 'Moisture Content (%)'])  # Header

    print("\n[⏳] Logging... (press Ctrl+C to stop safely)\n")

    try:
        start_timestamp = datetime.now()  # Anchor timestamp once

        while True:
            line = ser.readline().decode('ascii', errors='ignore').strip()
            if not line:
                continue

            # --- Parse metadata once ---
            if not metadata:
                if "TYPE" in line:
                    m = re.search(r'TYPE\s+(\S+)', line)
                    if m: metadata["Instrument Type"] = m.group(1)

                if "SN" in line:
                    m = re.search(r'SN\s+(\S+)', line)
                    if m: metadata["Serial Number"] = m.group(1)

                if "DATE" in line:
                    m = re.search(r'DATE\s+(\S+)', line)
                    if m: metadata["Date"] = m.group(1)

                if "TIME" in line:
                    m = re.search(r'TIME\s+(\S+)', line)
                    if m: metadata["Start Time"] = m.group(1)

                if "TEMP" in line:
                    m = re.search(r'TEMP\s+(\S+)', line)
                    if m: metadata["Temperature"] = m.group(1)

                if "Wet W(g)" in line:
                    m = re.search(r'Wet W\(g\)\s+([\d.]+)', line)
                    if m: metadata["Wet Weight (g)"] = float(m.group(1))

            # --- Parse measurement data ---
            match = re.search(r'(\d{2}:\d{2}:\d{2})\s+([\d.]+)', line)
            if match:
                elapsed_str = match.group(1)  # e.g., '00:01:30'
                moisture = float(match.group(2))

                # Convert elapsed time to timedelta
                h, m, s = map(int, elapsed_str.split(':'))
                elapsed_delta = pd.Timedelta(hours=h, minutes=m, seconds=s)

                # Compute aligned timestamp
                aligned_timestamp = start_timestamp + elapsed_delta
                aligned_str = aligned_timestamp.strftime('%Y/%m/%d_%H:%M:%S')

                # ✅ Print consistent log line
                print(f"{aligned_str}    {line}")

                # Write to CSV
                writer.writerow([aligned_str, elapsed_str, moisture])
                csvfile.flush()

    except KeyboardInterrupt:
        print("\n[!] Logging stopped by user.")

    finally:
        ser.close()
        print("[✓] Serial port closed.")

# --- Convert to Excel with metadata ---
print("[🔁] Converting to Excel...")

# Read the CSV data
df_data = pd.read_csv(CSV_FILE)

# Create metadata DataFrame
df_meta = pd.DataFrame.from_dict(metadata, orient='index', columns=['Value'])

# Write both sheets to Excel
with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl') as writer:
    df_data.to_excel(writer, sheet_name='Measurement Data', index=False)
    df_meta.to_excel(writer, sheet_name='Metadata')

print(f"\n✅ Excel file saved as: {EXCEL_FILE}")
