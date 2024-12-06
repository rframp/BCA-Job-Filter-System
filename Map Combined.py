import pandas as pd
import folium
import random
from datetime import datetime

# Load the saved routes data from the Excel file
file_path = r"C:\Users\RFrampton\OneDrive - Cowan Motor Group\Desktop\BCA Work.xlsx"
routes_data = pd.read_excel(file_path, sheet_name="Optimal Scotland Routes")

# Remove rows where CollLatitude, CollLongitude, DelLatitude, or DelLongitude are NaN
coll_data = routes_data.dropna(subset=['CollLatitude', 'CollLongitude'])
del_data = routes_data.dropna(subset=['DelLatitude', 'DelLongitude'])

# Define the starting location (Livingstone) and delivery endpoint
starting_location = (55.899819685016475, -3.5198384054833203)
end_location = (52.26700759136509, -0.7527653741274775)

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

# Add a marker for the delivery endpoint
folium.Marker(
    end_location,
    popup="Delivery Endpoint",
    icon=folium.Icon(color="green", icon="flag")
).add_to(m)

# Plot Collection Routes
for route_id, group in coll_data.groupby("RouteID"):
    color = random_color()  # Generate a new color for each route

    # Create round-trip coordinates for each collection route
    route_coordinates = [starting_location] + list(zip(group['CollLatitude'], group['CollLongitude'])) + [starting_location]
    total_distance = group['Round Trip Distance (miles)'].iloc[0]

    # Add line to map to represent the round-trip collection route
    folium.PolyLine(route_coordinates, color=color, weight=5, opacity=0.7,
                    tooltip=f"Collection Route {route_id} - Total Distance: {total_distance} miles").add_to(m)
    
    # Add markers for each job in the collection route
    for _, row in group.iterrows():
        folium.Marker(
            (row['CollLatitude'], row['CollLongitude']),
            popup=f"Job: {row['JobNumber']}\nCollection Route: {route_id}",
            icon=folium.Icon(color="blue"),
        ).add_to(m)

# Plot Delivery Routes
for route_id, group in del_data.groupby("RouteID"):
    color = random_color()  # Generate a new color for each route

    # Create one-way coordinates for each delivery route (start to end location)
    route_coordinates = [starting_location] + list(zip(group['DelLatitude'], group['DelLongitude'])) + [end_location]
    total_distance = group['One-Way Trip Distance (miles)'].iloc[0]

    # Add line to map to represent the one-way delivery route
    folium.PolyLine(route_coordinates, color=color, weight=5, opacity=0.7,
                    tooltip=f"Delivery Route {route_id} - Total Distance: {total_distance} miles").add_to(m)
    
    # Add markers for each job in the delivery route
    for _, row in group.iterrows():
        folium.Marker(
            (row['DelLatitude'], row['DelLongitude']),
            popup=f"Job: {row['JobNumber']}\nDelivery Route: {route_id}",
            icon=folium.Icon(color="purple"),
        ).add_to(m)

# Save the map as an HTML file
timestamp = datetime.now().strftime("%d%m%Y_%H%M%S")
output_path = fr"C:\Users\RFrampton\OneDrive - Cowan Motor Group\Desktop\Optimised Routes\optimised_routes_map_{timestamp}.html"
m.save(output_path)
print(f"Map saved as '{output_path}'")
