import folium
import pandas as pd

# Load the data from Excel (assumes latitude and longitude columns are already created)
file_path = r"C:\Users\RFrampton\OneDrive - Cowan Motor Group\Desktop\BCA Work.xlsx"
scotland_data = pd.read_excel(file_path, sheet_name="Scotland")

# Initialize the map centered around an average location
center_lat = scotland_data['CollLat'].mean()
center_lon = scotland_data['CollLon'].mean()
map_osm = folium.Map(location=[center_lat, center_lon], zoom_start=6)

# Add markers for each pickup and delivery point
for _, row in scotland_data.iterrows():
    # Marker for CollPostCode
    folium.Marker(
        location=[row['CollLat'], row['CollLon']],
        popup=f"Pickup: {row['CollPostCode']}",
        icon=folium.Icon(color="blue", icon="cloud")
    ).add_to(map_osm)

    # Marker for DelPostCode
    folium.Marker(
        location=[row['DelLat'], row['DelLon']],
        popup=f"Delivery: {row['DelPostCode']}",
        icon=folium.Icon(color="red", icon="info-sign")
    ).add_to(map_osm)

    # Draw a line between CollPostCode and DelPostCode
    folium.PolyLine(
        locations=[(row['CollLat'], row['CollLon']), (row['DelLat'], row['DelLon'])],
        color="green"
    ).add_to(map_osm)

# Save the map as an HTML file
map_path = r"C:\Users\RFrampton\OneDrive - Cowan Motor Group\Desktop\interactive_map.html"
map_osm.save(map_path)

print(f"Interactive map created and saved to {map_path}")
