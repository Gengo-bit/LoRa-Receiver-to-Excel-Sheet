import serial
import openpyxl
from datetime import datetime

ser = serial.Serial('COM10', 115200) 

wb = openpyxl.Workbook()
sheet = wb.active
sheet.title = "LoRa Data"
sheet.append(["Time", "Packet"])  # header

print("Listening to LoRa data...")

try:
    while True:
        line = ser.readline().decode('utf-8', errors='ignore').strip()

        if "," in line:
            parts = line.split(",", 1)  # split on first comma
            timestamp = parts[0].strip()
            packet = parts[1].strip()

            print(f"{timestamp} -> {packet}")
            
            sheet.append([timestamp, packet])
            wb.save("lora_data.xlsx")
        else:
            print(f"[Warning] Skipping malformed line: {line}")

except KeyboardInterrupt:
    print("Logging stopped.")
    ser.close()
