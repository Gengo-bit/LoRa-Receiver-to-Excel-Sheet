import openpyxl

wb = openpyxl.load_workbook("lora_data.xlsx")
sheet = wb.active

last_values = {}
data_losses = {}

for row in sheet.iter_rows(min_row=2, values_only=True):
    timestamp, packet = row

    if packet and ":" in packet:
        address, value = packet.split(":")
        value = int(value)

        if address in last_values:
            expected_value = last_values[address] + 1
            if value > expected_value:
                loss_count = value - expected_value
                print(f"[CORRUPTED DATA!] {loss_count} missing packet(s) for {address} between {last_values[address]} and {value}")
                data_losses[address] = data_losses.get(address, 0) + loss_count

        last_values[address] = value

print("\n--- Data Loss Summary ---")
for address in ["0000", "FFFF", "1111"]:
    losses = data_losses.get(address, 0)
    print(f"{address}: {losses} packet(s) lost")

