import openpyxl

wb = openpyxl.load_workbook("lora_data.xlsx")
sheet = wb.active

last_values = {}
data_losses = {}
packet_counts = {}

for row in sheet.iter_rows(min_row=2, values_only=True):
    timestamp, packet = row

    if packet and ":" in packet:
        address, value = packet.split(":")
        value = int(value)

        # total count
        packet_counts[address] = packet_counts.get(address, 0) + 1

        if address in last_values:
            expected_value = last_values[address] + 1
            if value > expected_value:
                loss_count = value - expected_value
                print(f"[LOSS] {loss_count} missing packet(s) for {address} between {last_values[address]} and {value}")
                data_losses[address] = data_losses.get(address, 0) + loss_count

        last_values[address] = value

print("\n--- Data Loss Summary ---")
for address in ["0000", "FFFF", "1111"]:
    losses = data_losses.get(address, 0)
    total_packets = packet_counts.get(address, 0)
    
    # Avoid divide by zero
    if total_packets + losses > 0:
        loss_percentage = (losses / (total_packets + losses)) * 100
    else:
        loss_percentage = 0.0

    print(f"{address}:")
    print(f"  Total Packets Received: {total_packets}")
    print(f"  Packets Lost: {losses}")
    print(f"  Data Loss Percentage: {loss_percentage:.2f}%\n")
