import openpyxl

wb = openpyxl.load_workbook("lora_data.xlsx")
sheet = wb.active

last_values = {}
data_losses = {}
packet_counts = {}
resets_detected = {}

for row in sheet.iter_rows(min_row=2, values_only=True):
    timestamp, packet = row

    if packet and ":" in packet:
        address, value = packet.split(":")
        value = int(value)

        packet_counts[address] = packet_counts.get(address, 0) + 1

        if address in last_values:
            # reset detected
            if value < last_values[address]:
                print(f"[RESET] {address} reset detected at value {value} (previous was {last_values[address]})")
                resets_detected[address] = resets_detected.get(address, 0) + 1

            # normal loss detection if value increased by more than 1
            elif value > last_values[address] + 1:
                loss_count = value - last_values[address] - 1
                print(f"[LOSS] {loss_count} missing packet(s) for {address} between {last_values[address]} and {value}")
                data_losses[address] = data_losses.get(address, 0) + loss_count

        # update last value
        last_values[address] = value

print("\n--- Data Loss Summary (with Resets) ---")
for address in sorted(packet_counts.keys()):
    losses = data_losses.get(address, 0)
    total_packets = packet_counts.get(address, 0)
    resets = resets_detected.get(address, 0)

    if total_packets + losses > 0:
        loss_percentage = (losses / (total_packets + losses)) * 100
    else:
        loss_percentage = 0.0

    print(f"{address}:")
    print(f"  Total Packets Received: {total_packets}")
    print(f"  Packets Lost: {losses}")
    print(f"  Data Loss Percentage: {loss_percentage:.2f}%")
    print(f"  Resets Detected: {resets}\n")
