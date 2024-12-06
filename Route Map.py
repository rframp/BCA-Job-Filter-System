# Map Generation Script

import pandas as pd
import folium
import random
from datetime import datetime

# Load the saved routes data from the Excel file
file_path = r"C:\Users\RFrampton\OneDrive - Cowan Motor Group\Desktop\BCA Work.xlsx"
routes_data = pd.read_excel(file_path, sheet_name="Optimal Scotland Routes")

# Remove rows where Latitude or Longitude is NaN
routes_data = routes_data.dropna(subset=['Latitude', 'Longitude'])

# Define the starting location (Livingstone)
starting_location = (55.899819685016475, -3.5198384054833203)

# Initialize the map centered around the starting location
map_center = starting_location
m = folium.Map(location=map_center, zoom_start=8)

# Function to generate a random color
def random_color():
    return "#{:06x}".format(random.randint(0, 0xFFFFFF))

# Add a marker for the depot (starting location)
folium.Marker(
    starting_location,
    popup="Depot (Starting Location)",
    icon=folium.Icon(color="red", icon="info-sign")
).add_to(m)

# Group the data by RouteID and add them to the map as complete routes
for route_id, group in routes_data.groupby("RouteID"):
    color = random_color()  # Generate a new color for each route

    # Create round-trip coordinates for each route
    route_coordinates = [starting_location] + list(zip(group['Latitude'], group['Longitude'])) + [starting_location]
    total_distance = group['Round Trip Distance (miles)'].iloc[0]

    # Add line to map to represent the round-trip route
    folium.PolyLine(route_coordinates, color=color, weight=5, opacity=0.7,
                    tooltip=f"Route {route_id} - Total Distance: {total_distance} miles").add_to(m)
    
    # Add markers for each job in the route
    for _, row in group.iterrows():
        folium.Marker(
            (row['Latitude'], row['Longitude']),
            popup=f"Job: {row['JobNumber']}\nRoute: {route_id}",
            icon=folium.Icon(color="blue"),
        ).add_to(m)

# Save the map as an HTML file
timestamp = datetime.now().strftime("%d%m%Y_%H%M%S")
m.save(fr"C:\Users\RFrampton\OneDrive - Cowan Motor Group\Desktop\Optimised Routes\optimised_routes_map_{timestamp}.html")
print(f"Map saved as 'optimised_routes_map_{timestamp}.html'")
