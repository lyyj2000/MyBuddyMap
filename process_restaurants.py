import pandas as pd
import requests
import json
import time
import os
import sys

# --- CONFIGURATION ---
INPUT_EXCEL = "Pets_places rest.xlsx"           # Your original file
OUTPUT_EXCEL = "Pets_places_rest_updated.xlsx"  # The file build_map.py reads
NEEDS_REVIEW_FILE = "needs_review.csv"
API_ERRORS_FILE = "data_discrepancies.csv"

def main():
    # 1. Ask for API Key
    api_key = os.environ.get("GOOGLE_API_KEY")
    if not api_key:
        api_key = input("Enter your Google API Key: ").strip()
    
    if not os.path.exists(INPUT_EXCEL):
        print(f"❌ Error: Cannot find '{INPUT_EXCEL}' in this folder.")
        sys.exit(1)

    print(f"📖 Reading {INPUT_EXCEL}...")
    df = pd.read_excel(INPUT_EXCEL)
    
    # Create columns for the new data
    place_ids = []
    live_ratings = []
    live_review_counts = [] # <-- CRITICAL FIX: This holds the real numbers
    reviews_data = {}

    matched_count = 0
    unmatched_count = 0
    needs_review_rows = []
    api_error_rows = []

    print("🔍 Searching Google Places API...")
    
    for index, row in df.iterrows():
        query = f"{row.get('Name', '')} {row.get('Address', '')} Singapore"
        
        url = "https://places.googleapis.com/v1/places:searchText"
        headers = {
            "Content-Type": "application/json",
            "X-Goog-Api-Key": api_key,
            # WE MUST ASK FOR userRatingCount or Google gives us 0!
            "X-Goog-FieldMask": "places.id,places.displayName,places.rating,places.userRatingCount,places.reviews"
        }
        payload = {"textQuery": query}

        try:
            response = requests.post(url, headers=headers, json=payload, timeout=10)
            data = response.json()
            
            places = data.get('places', [])
            if places:
                best_place = places[0]
                matched_count += 1
                
                p_id = best_place.get('id', '')
                # Extract the real rating
                p_rating = best_place.get('rating', 0)
                # Extract the REAL total review count
                p_review_count = best_place.get('userRatingCount', 0) 
                # Extract the max 5 reviews for the popup
                p_reviews = best_place.get('reviews', [])
                
                place_ids.append(p_id)
                live_ratings.append(p_rating)
                live_review_counts.append(p_review_count)
                reviews_data[p_id] = p_reviews[:5]
                
                print(f"✅ {row.get('Name', '')[:20]}... -> {p_review_count} reviews")
            else:
                unmatched_count += 1
                needs_review_rows.append(row.to_dict())
                place_ids.append('')
                live_ratings.append(row.get('Google Rating', 0))
                live_review_counts.append(0)
                print(f"⚠️ No match found for: {row.get('Name', '')}")
                
        except Exception as e:
            print(f"❌ API Error for {row.get('Name', '')}: {e}")
            api_error_rows.append(row.to_dict())
            place_ids.append('')
            live_ratings.append(0)
            live_review_counts.append(0)
            
        # Small delay to respect API rate limits
        time.sleep(0.5)

    # 2. Save the updated data back to Excel
    df['Place ID'] = place_ids
    df['Google Rating'] = live_ratings
    df['Review Count'] = live_review_counts # <-- Saving to the column build_map.py looks for
    
    df.to_excel(OUTPUT_EXCEL, index=False)
    
    if needs_review_rows:
        pd.DataFrame(needs_review_rows).to_csv(NEEDS_REVIEW_FILE, index=False)
    if api_error_rows:
        pd.DataFrame(api_error_rows).to_csv(API_ERRORS_FILE, index=False)

    # 3. Save the reviews cache (for the popups)
    with open('reviews_cache.json', 'w', encoding='utf-8') as f:
        json.dump(reviews_data, f, ensure_ascii=False, indent=2)

    print("\n=== Processing Report ===")
    print(f"Total Rows: {len(df)}")
    print(f"Rows Matched: {matched_count}")
    print(f"Rows Unmatched: {unmatched_count}")
    print(f"✅ Success! Saved to {OUTPUT_EXCEL}")

if __name__ == '__main__':
    main()