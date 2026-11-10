import pandas as pd

print("Loading PET_PLACES.xlsx...")
df = pd.read_excel("PET_PLACES.xlsx")

# This removes any duplicate column names, keeping only the first one (the good one)
df = df.loc[:, ~df.columns.duplicated()]

# Save it back
df.to_excel("PET_PLACES.xlsx", index=False)
print("SUCCESS! Duplicate columns removed. Your ratings are safe.")