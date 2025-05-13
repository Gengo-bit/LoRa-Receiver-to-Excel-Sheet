import openpyxl

wb = openpyxl.load_workbook("lora_data.xlsx")
sheet = wb.active

last_values = {}
data_losses = {}
packet_counts = {}
resets_detected = {}
cumulative_received = {}

for row in sheet.iter_rows(min_row=2, values_only=True):
    timestamp, packet = row

    if packet and ":" in packet:
        address, value = packet.split(":")
        try:
            value = int(value)
        except ValueError:
            continue  # skip if value is not an integer

        packet_counts.setdefault(address, 0)
        data_losses.setdefault(address, 0)
        resets_detected.setdefault(address, 0)
        cumulative_received.setdefault(address, 0)

        if address in last_values:
            last_val = last_values[address]

            if value < last_val:

                print(f"[RESET] {address} reset at value {value} (previous was {last_val})")
                resets_detected[address] += 1
                cumulative_received[address] += last_val + 1  # +1 because 0 is also a valid packet
            elif value > last_val + 1:
                loss_count = value - last_val - 1
                print(f"[LOSS] {loss_count} packet(s) lost for {address} between {last_val} and {value}")
                data_losses[address] += loss_count

        last_values[address] = value
        packet_counts[address] += 1

# add last known value to cumulative count
for address in last_values:
    cumulative_received[address] += last_values[address] + 1  # +1 to include current value

for address in sorted(packet_counts.keys()):
    losses = data_losses[address]
    resets = resets_detected[address]
    total_received = cumulative_received[address]

    if total_received + losses > 0:
        loss_percentage = (losses / (total_received + losses)) * 100
    else:
        loss_percentage = 0.0

    print(f"{address}:")
    print(f"  Total Packets Received (including resets): {total_received}")
    print(f"  Packets Lost: {losses}")
    print(f"  Data Loss Percentage: {loss_percentage:.2f}%")
    print(f"  Resets Detected: {resets}\n")
