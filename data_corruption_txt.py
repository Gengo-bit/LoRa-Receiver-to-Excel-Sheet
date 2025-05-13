import re

with open("lora_data.txt", "r") as file:
    content = file.read()

losses = re.findall(r'\[LOSS\] (\d+) packet', content)
total_lost = sum(int(loss) for loss in losses)

print(f"Total packets lost: {total_lost}")
