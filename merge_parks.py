import pandas as pd

print("Loading files...")
# Load the new Google file
new_df = pd.read_excel("PET_PLACES.xlsx")

# Load the old BAD file
old_df = pd.read_excel("PET_PLACES_BAD.xlsx")

# Filter the old file to keep ONLY Parks and Training
parks_training = old_df[old_df['Category'].isin(['Park', 'Training'])]

print(f"Found {len(parks_training)} Parks and Training locations in the backup.")

# Merge them together. Pandas will automatically align the columns by name.
merged_df = pd.concat([new_df, parks_training], ignore_index=True)

# Save it back to the main file
merged_df.to_excel("PET_PLACES.xlsx", index=False)
print(f"SUCCESS! Merged. Total locations now: {len(merged_df)}")