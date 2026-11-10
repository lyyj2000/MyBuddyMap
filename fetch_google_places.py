import requests
import pandas as pd
import time
import re

# PASTE YOUR API KEY HERE
API_KEY = "AIzaSyBQpDyF91lv1xy9E9d9lHcG_o7YV1qHD4Q"

SEARCH_QUERIES = [
    "pet groomer", "pet grooming", "dog groomer", 
    "pet cafe", "dog cafe", "pet friendly cafe",
    "pet friendly restaurant", "dog friendly restaurant", "pet friendly dining"
]

def extract_postal_code(address):
    match = re.search(r'Singapore\s(\d{6})', address)
    return match.group(1) if match else ""

def find_social_links(website_url):
    fb, ig = "", ""
    if not website_url:
        return fb, ig
    try:
        headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'}
        res = requests.get(website_url, timeout=5, headers=headers)
        fb_match = re.search(r'(https?://(?:www\.)?facebook\.com/[a-zA-Z0-9_\.-]+)', res.text)
        ig_match = re.search(r'(https?://(?:www\.)?instagram\.com/[a-zA-Z0-9_\.-]+)', res.text)
        if fb_match: fb = fb_match.group(1)
        if ig_match: ig = ig_match.group(1)
    except Exception:
        pass
    return fb, ig

def search_places(query):
    url = "https://places.googleapis.com/v1/places:searchText"
    headers = {
        "Content-Type": "application/json",
        "X-Goog-Api-Key": API_KEY,
        "X-Goog-FieldMask": "places.id,places.displayName,places.formattedAddress,places.rating,places.userRatingCount,places.websiteUri,places.nationalPhoneNumber,places.location,nextPageToken"
    }
    
    # Strictly limit the search to Singapore's GPS boundaries
    payload = {
        "textQuery": query,
        "languageCode": "en",
        "locationRestriction": {
            "rectangle": {
                "low": {"latitude": 1.15, "longitude": 103.60},
                "high": {"latitude": 1.48, "longitude": 104.10}
            }
        }
    }
    
    all_results = []
    
    # Loop to get multiple pages of results
    while True:
        response = requests.post(url, headers=headers, json=payload, timeout=10)
        if response.status_code != 200:
            print(f"API Error: {response.text}")
            break
            
        data = response.json()
        all_results.extend(data.get('places', []))
        
        # Check if there is a next page
        if 'nextPageToken' in data:
            payload['pageToken'] = data['nextPageToken']
            time.sleep(1) # Wait 1 second between pages
        else:
            break
            
    return all_results

def main():
    all_places = {}
    
    for query in SEARCH_QUERIES:
        print(f"Searching for: {query}...")
        places = search_places(query)
        print(f"  -> Found {len(places)} results for this search.")
        
        for place in places:
            place_id = place.get('id')
            if not place_id:
                continue
                
            if place_id not in all_places:
                if "groom" in query: category = "Groomer"
                elif "cafe" in query: category = "Cafe"
                else: category = "Dine"
                
                name = place.get('displayName', {}).get('text', '')
                address = place.get('formattedAddress', '')
                postal = extract_postal_code(address)
                website = place.get('websiteUri', '')
                
                print(f"    -> Processing: {name}...")
                fb, ig = find_social_links(website)
                
                all_places[place_id] = {
                    "Place ID": place_id,
                    "Name": name,
                    "Category": category,
                    "Address": address,
                    "Postal Code": postal,
                    "Phone": place.get('nationalPhoneNumber', ''),
                    "Operating Hours": "", 
                    "Rating": place.get('rating', ''),
                    "Review Count": place.get('userRatingCount', ''),
                    "Website": website,
                    "Facebook URL": fb,
                    "IG URL": ig,
                    "Latitude": place.get('location', {}).get('latitude', ''),
                    "Longitude": place.get('location', {}).get('longitude', '')
                }
                
    print(f"\nTotal unique places found: {len(all_places)}")
    if not all_places: return
        
    df = pd.DataFrame(list(all_places.values()))
    df = df[["Place ID", "Name", "Category", "Address", "Postal Code", "Phone", "Operating Hours", "Rating", "Review Count", "Website", "Facebook URL", "IG URL", "Latitude", "Longitude"]]
    df.to_excel("PET_PLACES_GOOGLE.xlsx", index=False)
    print(f"SUCCESS! Saved {len(df)} records to PET_PLACES_GOOGLE.xlsx")

if __name__ == "__main__":
    main()