import sys
import os

print("Applying Map UI Fixes...")

if not os.path.exists('index.html'):
    print("Error: Could not find index.html.")
    sys.exit()

with open('index.html', 'r', encoding='utf-8') as f:
    html = f.read()

# 1. Search Button
html = html.replace('onclick="map.setView(map.getCenter(), map.getZoom())"', 'onclick="updateMapCentre(map.getCenter())"')

# 2. Popup CSS
if '.leaflet-popup-content-wrapper' not in html:
    css = "\n.leaflet-popup-content-wrapper { border-radius: 16px !important; padding: 0; overflow: hidden; }\n.leaflet-popup-content { max-height: 320px !important; overflow-y: auto !important; margin: 0 !important; }\n"
    html = html.replace('</style>', css + '</style>')

# 3. Popup Cut-off
html = html.replace('marker.bindPopup(() => popupHTML(c, statusFor(c)));', 'marker.bindPopup(() => popupHTML(c, statusFor(c)), { maxHeight: 320, autoPanPaddingTopLeft: [20, 60], autoPanPaddingBottomRight: [20, 20] });')

# 4. Reviews
html = html.replace(
    '''let ratingHTML = c.rating > 0 ? `<div class="rating">⭐ ${parseFloat(c.rating).toFixed(1)} ${c.review_count > 0 ? `<span>(${c.review_count} reviews)</span>` : ''}</div>` : "";''',
    '''let showReviewCount = false;
      if (c.category === "Vet" || ["Asian", "Western", "Cafe", "Dine"].includes(c.category)) {
          showReviewCount = true;
      }
      let rc = Number(c.review_count) || 0; 
      let reviewStr = (showReviewCount && rc > 0) ? `<span>(${rc} reviews)</span>` : '';
      let ratingHTML = c.rating > 0 ? `<div class="rating">⭐ ${parseFloat(c.rating).toFixed(1)} ${reviewStr}</div>` : "";'''
)

# 5. Slider HTML
html = html.replace(
    '''<div class="toggle-item">
          <span>⭐ Top Rated (4.5★ & Above)</span>
          <input type="checkbox" id="filter-top">
      </div>''',
    '''<div class="radius-box" style="margin-top: 12px;">
          <h4>Minimum Rating <span id="ratingVal">3★ & Above</span></h4>
          <input id="rating-slider" type="range" min="3" max="5" value="3" step="0.5">
          <div class="radius-labels">
              <span>3★</span><span>4★</span><span>5★</span>
          </div>
      </div>'''
)

# 6. Slider JavaScript
html = html.replace('let filterTopRated = false;', 'let minRating = 3;')
html = html.replace(
    'document.getElementById("filter-top").addEventListener("change", e => { filterTopRated = e.target.checked; updateVisibility(); });',
    '''document.getElementById("rating-slider").addEventListener("input", e => {
    minRating = parseFloat(e.target.value);
    document.getElementById("ratingVal").textContent = (minRating > 3) ? minRating.toFixed(1) + "★ & Above" : "3★ & Above";
    updateVisibility();
});'''
)
html = html.replace('if (show && filterTopRated && c.rating < 4.5) show = false;', 'if (show && c.rating > 0 && c.rating < minRating) show = false;')
html = html.replace('filterTopRated = false;\n    document.getElementById("filter-top").checked = false;', 'minRating = 3;\n    document.getElementById("rating-slider").value = 3;\n    document.getElementById("ratingVal").textContent = "3★ & Above";')

# 7. StatusFor and InWindow
old_status = '''function statusFor(c) {
  if (c.ecat === "E1_CONFIRMED_24_7") return { key:"open", label:"OPEN NOW · 24/7 EMERGENCY" };
  const { day, mins } = singaporeNow();
  const iv = c.schedule[day];
  if (iv === "appointment") return { key:"appt", label:"BY APPOINTMENT ONLY" };
  if (iv === null || iv === undefined) return { key:"unknown", label:"HOURS UNCERTAIN" };
  if (!iv.length) return { key:"closed", label:"CLOSED NOW" };
  return iv.some(([a,b]) => a <= mins && mins < b) ? { key:"open",  label:"OPEN NOW" } : { key:"closed", label:"CLOSED NOW" };
}'''
new_status = '''function inWindow(mins, a, b) {
  if (a <= b) return a <= mins && mins < b;
  return mins >= a || mins < b;
}

function statusFor(c) {
  const { day, mins } = singaporeNow();
  const iv = c.schedule[day];
  const is247 = c.ecat === "E1_CONFIRMED_24_7";

  if (iv === "appointment") return { key:"appt", label:"BY APPOINTMENT ONLY" };
  if (iv === null || iv === undefined) return { key:"unknown", label:"HOURS UNCERTAIN" };
  if (!iv.length) {
    return is247
      ? { key:"unknown", label:"24/7 STATUS CONFLICT — CALL FIRST" }
      : { key:"closed",  label:"CLOSED NOW" };
  }
  const open = iv.some(([a, b]) => inWindow(mins, a, b));
  if (open) return { key:"open", label: is247 ? "OPEN NOW · 24/7 EMERGENCY" : "OPEN NOW" };
  return { key:"closed", label:"CLOSED NOW" };
}'''
html = html.replace(old_status, new_status)

# 8. UpdateVisibility (Blinking/Search)
old_vis = '''function updateVisibility() {
  const centre = centreCircle.getLatLng(); const radius = centreCircle.getRadius();
  const q = searchText.trim().toLowerCase(); const hasSearch = q.length > 0;
  let visible = 0;

  markers.forEach(m => {
    const c = m.clinic, st = m.status;
    const isEmergency = c.ecat === "E1_CONFIRMED_24_7";
    
    let show = true;
    
    if (isEmergency) {
        show = true;
    } else {
        if (hasSearch) {
            let match = c.name.toLowerCase().includes(q) || c.id.toLowerCase().includes(q) || c.category.toLowerCase().includes(q);
            if (!match && /^\\d{6}$/.test(q)) match = (c.postal === q);
            show = match;
        } else {
            if (currentCategory === "All") {
            } else if (currentCategory === "Park") {
                if (c.category !== "Park") show = false;
            } else {
                if (c.category === "Park") show = false; 
                else if (c.category !== currentCategory) {
                    if (currentCategory === "Dine" && ["Asian", "Western", "Cafe", "Dine"].includes(c.category)) {
                    } else {
                        show = false;
                    }
                }
            }
            if (show && currentSpecies !== "all" && currentCategory === "Vet" && c.category === "Vet") {
              const sp = c.species || [];
              if (sp.length > 0 && !sp.includes(currentSpecies)) show = false;
            }
        }
        
        if (show && filterTopRated && c.rating < 4.5) show = false;

        if (show && !hasSearch) {
          show = map.distance(centre, m.marker.getLatLng()) <= radius;
        }
    }

    if (show) { 
        if (!layer.hasLayer(m.marker)) layer.addLayer(m.marker); 
        visible++; 
    }
    else if (layer.hasLayer(m.marker)) layer.removeLayer(m.marker);
  });
  
  let catName = "locations";
  if (currentCategory === "Vet") catName = "vetted clinics";
  if (currentCategory === "Groomer") catName = "groomers";
  if (currentCategory === "Mall") catName = "pet malls";
  if (currentCategory === "Training") catName = "training schools";
  if (currentCategory === "Dine") catName = "pet dining spots";

  const radiusKm = radius / 1000;
  document.getElementById("bottom-text").innerHTML = `Showing <b>${visible}</b> ${catName} within your ${radiusKm} km search radius.`;
}

document.getElementById("radius").addEventListener("input", e => {
  const km = parseFloat(e.target.value);
  centreCircle.setRadius(km * 1000);
  document.getElementById("radiusVal").textContent = km + " km";
  updateVisibility();
});'''
new_vis = '''function updateVisibility() {
  const centre = centreCircle.getLatLng(); const radius = centreCircle.getRadius();
  const q = searchText.trim().toLowerCase(); const hasSearch = q.length > 0;
  let visible = 0;
  
  document.body.classList.toggle('is-searching', hasSearch);
  
  markers.forEach(m => {
    const c = m.clinic, st = m.status;
    const isEmergency = c.ecat === "E1_CONFIRMED_24_7";
    
    let show = true;
    
    if (hasSearch) {
        let match = c.name.toLowerCase().includes(q) || c.id.toLowerCase().includes(q) || c.category.toLowerCase().includes(q);
        if (!match && /^\\d{6}$/.test(q)) match = (c.postal === q);
        show = match;
    } else {
        if (isEmergency) {
            show = true;
        } else {
            if (currentCategory === "All") {
            } else if (currentCategory === "Park") {
                if (c.category !== "Park") show = false;
            } else {
                if (c.category === "Park") show = false; 
                else if (c.category !== currentCategory) {
                    if (currentCategory === "Dine" && ["Asian", "Western", "Cafe", "Dine"].includes(c.category)) {
                    } else {
                        show = false;
                    }
                }
            }
            if (show && currentSpecies !== "all" && currentCategory === "Vet" && c.category === "Vet") {
              const sp = c.species || [];
              if (sp.length > 0 && !sp.includes(currentSpecies)) show = false;
            }
            if (show && c.rating > 0 && c.rating < minRating) show = false;
            if (show) {
              show = map.distance(centre, m.marker.getLatLng()) <= radius;
            }
        }
    }

    if (show) { 
        if (!layer.hasLayer(m.marker)) layer.addLayer(m.marker); 
        visible++; 
    }
    else if (layer.hasLayer(m.marker)) layer.removeLayer(m.marker);
  });
  
  let catName = "locations";
  if (currentCategory === "Vet") catName = "vetted clinics";
  if (currentCategory === "Groomer") catName = "groomers";
  if (currentCategory === "Mall") catName = "pet malls";
  if (currentCategory === "Training") catName = "training schools";
  if (currentCategory === "Dine") catName = "pet dining spots";

  const radiusKm = radius / 1000;
  document.getElementById("bottom-text").innerHTML = `Showing <b>${visible}</b> ${catName} within your ${radiusKm} km search radius.`;
}

document.getElementById("radius").addEventListener("input", e => {
  const km = parseFloat(e.target.value);
  centreCircle.setRadius(km * 1000);
  document.getElementById("radiusVal").textContent = km + " km";
  updateVisibility();
});'''
html = html.replace(old_vis, new_vis)

# 9. Blinking CSS
if 'body.is-searching' not in html:
    blink_css = "\nbody.is-searching .emergency-pin .pin-img { animation: none !important; filter: drop-shadow(0px 4px 6px rgba(0,0,0,0.3)) !important; }\n"
    html = html.replace('@keyframes pulse-emergency', blink_css + '@keyframes pulse-emergency')

with open('index.html', 'w', encoding='utf-8') as f:
    f.write(html)

print("✅ Success! All fixes have been applied to index.html.")