with open('build_map.py', 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_block = False
for i in range(len(lines)):
    if 'for c in clinics:' in lines[i]:
        in_block = True
        continue

    if in_block:
        # Add 4 spaces back to all lines inside the loop
        if lines[i].strip():
            lines[i] = "    " + lines[i]

        if 'c["lng"] = None' in lines[i]:
            break

with open('build_map.py', 'w', encoding='utf-8') as f:
    f.writelines(lines)

print("Indentation fixed! Run python build_map.py now.")