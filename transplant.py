import sys

print("Starting the Data Transplant...")

# 1. Read the newly generated index.html (which has the real reviews but no sliders)
try:
    with open('index.html', 'r', encoding='utf-8') as f:
        new_html = f.read()
except FileNotFoundError:
    print("Error: Could not find the new index.html.")
    sys.exit()

# 2. Read your masterpiece index_FIXED.html (which has the sliders but old data)
try:
    with open('index_FIXED.html', 'r', encoding='utf-8') as f:
        fixed_html = f.read()
except FileNotFoundError:
    print("Error: Could not find index_FIXED.html.")
    sys.exit()

# 3. Extract the new data block from the new file
try:
    new_data = new_html.split('const DATA = ')[1].split('const SINGAPORE')[0].strip()
except IndexError:
    print("Error: Could not find the DATA block in the new index.html")
    sys.exit()

# 4. Extract the UI parts from your fixed file
try:
    part_before = fixed_html.split('const DATA = ')[0]
    part_after = 'const SINGAPORE' + fixed_html.split('const SINGAPORE')[1]
except IndexError:
    print("Error: Could not find the DATA block in index_FIXED.html")
    sys.exit()

# 5. Combine them! (UI fixes + New Data)
final_html = part_before + 'const DATA = ' + new_data + '\n\n' + part_after

# 6. Save it as the final index.html
with open('index.html', 'w', encoding='utf-8') as f:
    f.write(final_html)

print("✅ Success! The new index.html now has your UI fixes AND the real reviews.")