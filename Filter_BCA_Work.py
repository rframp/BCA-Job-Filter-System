import pandas as pd
import streamlit as st
import re
import io
import os
import base64
import time
import folium
from geopy.geocoders import Nominatim
from geopy.exc import GeocoderTimedOut
from concurrent.futures import ThreadPoolExecutor, as_completed
from streamlit_folium import st_folium


st.set_page_config(layout="wide")  # Enable wide mode for the app

st.markdown(
    """
    <style>
    .center-content {
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
    }
    .center-table {
        margin: auto;
        width: 80%; /* Adjust width as needed */
    }
    </style>
    """,
    unsafe_allow_html=True
)

def debug_map_state(event):
    print(f"{event}:")
    print(f" - st.session_state['map_center']: {st.session_state.get('map_center')}")
    print(f" - st.session_state['map_zoom']: {st.session_state.get('map_zoom')}")


def get_image_as_base64(image_path):
    with open(image_path, "rb") as image_file:
        return base64.b64encode(image_file.read()).decode("utf-8")
    

# Function to geocode postcodes
@st.cache_data(show_spinner=True)
def geocode_postcode(postcode):
    # Initialize geolocator
    geolocator = Nominatim(user_agent="streamlit_geocoder", timeout=10)
    try:
        location = geolocator.geocode(postcode)
        if location:
            return location.latitude, location.longitude
    except GeocoderTimedOut:
        time.sleep(1)
        return geocode_postcode(postcode)  # Retry on timeout
    return None, None

# Function to geocode postcodes in parallel
def parallel_geocode(postcodes):
    results = []
    with ThreadPoolExecutor(max_workers=10) as executor:  # Adjust workers for performance
        future_to_postcode = {executor.submit(geocode_postcode, postcode): postcode for postcode in postcodes}
        for future in as_completed(future_to_postcode):
            postcode = future_to_postcode[future]
            try:
                results.append((postcode, future.result()))
            except Exception as e:
                st.warning(f"Failed to geocode {postcode}: {e}")
                results.append((postcode, (None, None)))
    return results

# Function to process geocoding for the dataset with progress bar
@st.cache_data(show_spinner=False)
def process_geocoding(data):
    # Deduplicate postcodes to minimize requests
    unique_postcodes = pd.concat([data['CollPostCode'], data['DelPostCode']]).dropna().unique()
    total_postcodes = len(unique_postcodes)

    # Initialize progress bar
    progress_bar = st.progress(0)

    # Placeholder for results
    geocoded_results = []

    # Geocode postcodes with progress updates
    for idx, postcode in enumerate(unique_postcodes):
        # Perform geocoding (replace this with the actual geocoding logic)
        lat, lon = geocode_postcode(postcode)  # Replace `geocode_postcode` with your actual geocoding function
        geocoded_results.append((postcode, (lat, lon)))

        # Update progress bar
        progress_bar.progress(int(((idx + 1) / total_postcodes) * 100))

    # Map geocoding results to a dictionary
    geocode_dict = {postcode: coords for postcode, coords in geocoded_results}

    # Add latitude and longitude columns to the dataset
    data['CollLat'] = data['CollPostCode'].map(lambda x: geocode_dict.get(x, (None, None))[0])
    data['CollLon'] = data['CollPostCode'].map(lambda x: geocode_dict.get(x, (None, None))[1])
    data['DelLat'] = data['DelPostCode'].map(lambda x: geocode_dict.get(x, (None, None))[0])
    data['DelLon'] = data['DelPostCode'].map(lambda x: geocode_dict.get(x, (None, None))[1])

    # Replace missing latitude/longitude with "N/A"
    data.fillna({"CollLat": "N/A", "CollLon": "N/A", "DelLat": "N/A", "DelLon": "N/A"}, inplace=True)

    return data


def create_folium_map(data):
    """
    Generate a Folium map with markers for collection and delivery points.
    Combines markers with the same location into a single marker and adds job count.
    """
    # Define a default center and zoom for the map
    map_center = [data["CollLat"].mean(), data["CollLon"].mean()]
    map_zoom = 9

    # Create the map
    folium_map = folium.Map(location=map_center, zoom_start=map_zoom)

    # Group collection points by location
    coll_grouped = data.groupby(["CollLat", "CollLon"])["JobNumber"].apply(list).reset_index()
    for _, row in coll_grouped.iterrows():
        if row["CollLat"] != "N/A" and row["CollLon"] != "N/A":
            job_numbers = row["JobNumber"]
            job_count = len(job_numbers)
            job_list = ", ".join(job_numbers)
            popup_content = f"""
                <b>Total Jobs:</b> {job_count}<br>
                <b>Job Numbers:</b> {job_list}<br>
                <b>Type:</b> Collection
            """
            folium.Marker(
                location=[row["CollLat"], row["CollLon"]],
                popup=popup_content,
                tooltip=f"Collection Point - {job_count} Jobs",
                icon=folium.Icon(color="blue", icon="info-sign"),
            ).add_to(folium_map)

    # Group delivery points by location
    del_grouped = data.groupby(["DelLat", "DelLon"])["JobNumber"].apply(list).reset_index()
    for _, row in del_grouped.iterrows():
        if row["DelLat"] != "N/A" and row["DelLon"] != "N/A":
            job_numbers = row["JobNumber"]
            job_count = len(job_numbers)
            job_list = ", ".join(job_numbers)
            popup_content = f"""
                <b>Total Jobs:</b> {job_count}<br>
                <b>Job Numbers:</b> {job_list}<br>
                <b>Type:</b> Delivery
            """
            folium.Marker(
                location=[row["DelLat"], row["DelLon"]],
                popup=popup_content,
                tooltip=f"Delivery Point - {job_count} Jobs",
                icon=folium.Icon(color="green", icon="flag"),
            ).add_to(folium_map)

    return folium_map
    
def main():
    
    # Paths to images
    x_image_path = "images/X_New.png"
    bca_image_path = "images/BCA_New.png"
    cmg_image_path = "images/CMG_New.png"

    # Convert images to Base64
    x_image_base64 = get_image_as_base64(x_image_path)
    bca_image_base64 = get_image_as_base64(bca_image_path)
    cmg_image_base64 = get_image_as_base64(cmg_image_path)

    # Inject custom CSS for scrollbar and layout
    st.markdown(
        f"""
        <style>
        /* Custom scrollbar for all scrollable content */
        ::-webkit-scrollbar {{
            width: 15px; /* Width of the vertical scrollbar */
            height: 15px; /* Height of the horizontal scrollbar */
        }}

        /* Scrollbar track */
        ::-webkit-scrollbar-track {{
            background: #555; /* Light grey background for the track */
        }}

        /* Scrollbar thumb (the draggable handle) */
        ::-webkit-scrollbar-thumb {{
            background: #555; /* Default color of the scrollbar thumb */
            border-radius: 10px; /* Rounded corners */
            border: 3px solid #e0e0e0; /* Add a border to make it distinct */
        }}

        /* Scrollbar thumb hover (when the mouse hovers over the thumb) */
        ::-webkit-scrollbar-thumb:hover {{
            background: #555 !important; /* Dark grey for better visibility when hovered */
            border: 3px solid #e0e0e0; /* Maintain the border color */
        }}

        /* Scrollbar thumb active (when clicked or dragged) */
        ::-webkit-scrollbar-thumb:active {{
            background: #333 !important; /* Even darker grey when actively dragging */
            border: 3px solid #e0e0e0; /* Maintain the border color */
        }}

        /* Scrollbar corner (intersection of horizontal and vertical scrollbars) */
        ::-webkit-scrollbar-corner {{
            background: #e0e0e0; /* Matches the track color */
        }}

        /* Custom image alignment section */
        .image-container {{
            display: flex;
            justify-content: space-between; /* Adjust spacing between images */
            align-items: center;
            margin-bottom: 20px; /* Add space below the images */
            position: relative;
        }}

        /* Individual image positions */
        .bca-img {{
            flex-grow: 0 !important; /* Prevent growing */
            flex-shrink: 0 !important; /* Prevent shrinking */
            margin-left: 600px !important;
            width: 275px !important;
            height: 125px !important; /* Explicitly define size */
        }}
        .x-img {{
            flex-grow: 0 !important; /* Prevent growing */
            flex-shrink: 0 !important; /* Prevent shrinking */
            width: 95px !important; /* Explicitly define size */
            height: 95px !important; /* Explicitly define size */
        }}
        .cmg-img {{
            flex-grow: 0 !important; /* Prevent growing */
            flex-shrink: 0 !important; /* Prevent shrinking */
            margin-right:600px !important;
            width: 315px !important; /* Explicitly define size */
            height: 145px !important; /* Explicitly define size */
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )

    # HTML for displaying images
    image_html = f"""
    <div class="image-container">
        <img src="data:image/png;base64,{bca_image_base64}" alt="BCA" class="bca-img">
        <img src="data:image/png;base64,{x_image_base64}" alt="X" class="x-img">
        <img src="data:image/png;base64,{cmg_image_base64}" alt="CMG" class="cmg-img">
    </div>
    """

    # Inject the image HTML into the Streamlit app
    st.markdown(image_html, unsafe_allow_html=True)

    st.title("BCA Filtering Tool")
    
    uploaded_file = st.file_uploader("Upload your Excel or CSV file", type=["xlsx", "csv"])
    if uploaded_file:
        # Read the file
        if uploaded_file.name.endswith('.xlsx'):
            data = pd.read_excel(uploaded_file)
        else:
            data = pd.read_csv(uploaded_file)

        # Fix mixed-type columns
        for column in data.columns:
            if data[column].dtype == 'object':  # Mixed-type columns are usually 'object'
                data[column] = data[column].astype(str)  # Convert everything to string

        # **Fix Comma Handling in Columns (e.g., JobNumber)**
        data['JobNumber'] = data['JobNumber'].astype(str).str.replace(',', '')   # Remove commas from numbers
        if 'CustRef3' in data.columns:
            # Convert the column to string to handle mixed types
            data['CustRef3'] = data['CustRef3'].astype(str)
            data['CustRef3'] = data['CustRef3'].replace({'nan': 'N/A', 'None': 'N/A', '': 'N/A'}).fillna('N/A')

        # **Fix Date Columns (Optional)**
        date_columns = ['StartDate', 'EndDate', 'AgreedDate']
        for col in date_columns:
            if col in data.columns:
                data[col] = pd.to_datetime(data[col], errors='coerce').dt.strftime('%d/%m/%Y')

        # Clean and standardize postcodes
        data['CollPostCode'] = data['CollPostCode'].str.strip().str.upper()
        data['DelPostCode'] = data['DelPostCode'].str.strip().str.upper()

        with st.spinner("Geocoding postcodes... This may take some time."):
            # Deduplicate postcodes to minimize requests
            unique_postcodes = pd.concat([data['CollPostCode'], data['DelPostCode']]).dropna().unique()
            total_postcodes = len(unique_postcodes)

            # Initialize progress bar and timer display
            progress_bar = st.progress(0)
            timer_placeholder = st.empty()

            # Placeholder for geocoded results
            geocoded_results = []

            # Start timing
            start_time = time.time()

            # Geocode postcodes with progress updates
            for idx, postcode in enumerate(unique_postcodes):
                # Start timing for this iteration
                iteration_start_time = time.time()

                # Simulate geocoding (replace this with your geocoding logic)
                lat, lon = geocode_postcode(postcode)  # Replace with your actual geocoding function
                geocoded_results.append((postcode, (lat, lon)))

                # Calculate progress and update progress bar
                progress = int(((idx + 1) / total_postcodes) * 100)
                progress_bar.progress(progress)

                # Calculate average time per postcode dynamically
                elapsed_time = time.time() - start_time
                avg_time_per_postcode = elapsed_time / (idx + 1)

                # Estimate total and remaining time
                estimated_total_time = avg_time_per_postcode * total_postcodes
                remaining_time = estimated_total_time - elapsed_time

                # Update timer display dynamically
                timer_placeholder.write(f"Estimated time remaining: {remaining_time:.2f} seconds")

            # Map geocoding results to a dictionary
            geocode_dict = {postcode: coords for postcode, coords in geocoded_results}

            # Add latitude and longitude columns to the dataset
            data['CollLat'] = data['CollPostCode'].map(lambda x: geocode_dict.get(x, (None, None))[0])
            data['CollLon'] = data['CollPostCode'].map(lambda x: geocode_dict.get(x, (None, None))[1])
            data['DelLat'] = data['DelPostCode'].map(lambda x: geocode_dict.get(x, (None, None))[0])
            data['DelLon'] = data['DelPostCode'].map(lambda x: geocode_dict.get(x, (None, None))[1])
        
        progress_bar.empty()
        timer_placeholder.empty()

        # Ensure the geocoded data is valid
        if data[['CollLat', 'CollLon', 'DelLat', 'DelLon']].isna().all().all():
            st.error("No valid geolocation data found. Please check your file.")

        # Function to extract letters before the first number
        def extract_outward_code(postcode):
            # Match letters before the first number
            match = re.match(r"([A-Z]+)", postcode)
            if match:
                return match.group(1)  # Extract letters before the first number
            return None

        # Extract outward codes for collection and delivery postcodes
        data['CollOutwardCode'] = data['CollPostCode'].apply(extract_outward_code)
        data['DelOutwardCode'] = data['DelPostCode'].apply(extract_outward_code)
        # **Map Postcodes to Regions**
        postcode_to_region = {
            # Scotland
            "AB": "Scotland", "DD": "Scotland", "KW": "Scotland", "DG": "Scotland", "KY": "Scotland",
            "EH": "Scotland", "ML": "Scotland", "FK": "Scotland", "PA": "Scotland", "G": "Scotland",
            "PH": "Scotland", "TD": "Scotland", "IV": "Scotland",
            
            # Northern Ireland
            "BT": "Northern Ireland",
            
            # North East
            "DH": "North East", "NE": "North East", "DL": "North East", "SR": "North East",
            "HG": "North East", "TS": "North East", "HU": "North East", "WF": "North East",
            "LS": "North East", "YO": "North East",
            
            # North West
            "BB": "North West", "L": "North West", "BD": "North West", "LA": "North West",
            "BL": "North West", "M": "North West", "CA": "North West", "OL": "North West",
            "CH": "North West", "PR": "North West", "CW": "North West", "SK": "North West",
            "FY": "North West", "WA": "North West", "HD": "North West", "WN": "North West",
            "HX": "North West",
            
            # East Midlands
            "CB": "East Midlands", "LN": "East Midlands", "CO": "East Midlands", "NG": "East Midlands",
            "DE": "East Midlands", "NR": "East Midlands", "DN": "East Midlands", "PE": "East Midlands",
            "IP": "East Midlands", "S": "East Midlands", "LE": "East Midlands", "SS": "East Midlands",
            
            # West Midlands
            "B": "West Midlands", "ST": "West Midlands", "CV": "West Midlands", "TF": "West Midlands",
            "DY": "West Midlands", "WR": "West Midlands", "HR": "West Midlands", "WS": "West Midlands",
            "NN": "West Midlands", "WV": "West Midlands",
            
            # Wales
            "CF": "Wales", "NP": "Wales", "LD": "Wales", "SA": "Wales", "LL": "Wales", "SY": "Wales",
            
            # South West
            "BA": "South West", "PL": "South West", "BH": "South West", "SN": "South West",
            "BS": "South West", "SP": "South West", "DT": "South West", "TA": "South West",
            "EX": "South West", "TQ": "South West", "TR": "South West", "GL": "South West",
            
            # South East
            "AL": "South East", "OX": "South East", "BN": "South East", "PO": "South East",
            "CM": "South East", "RG": "South East", "CT": "South East", "RH": "South East",
            "GU": "South East", "SG": "South East", "HP": "South East", "SL": "South East",
            "LU": "South East", "SO": "South East", "ME": "South East", "SS": "South East",
            "MK": "South East", "TN": "South East",
            
            # Greater London
            "BR": "Greater London", "NW": "Greater London", "CR": "Greater London", "RM": "Greater London",
            "DA": "Greater London", "SE": "Greater London", "SM": "Greater London", "EC": "Greater London",
            "SW": "Greater London", "EN": "Greater London", "TW": "Greater London", "HA": "Greater London",
            "UB": "Greater London", "IG": "Greater London", "W": "Greater London", "KT": "Greater London",
            "WC": "Greater London", "WD": "Greater London"
        }

        # Create region columns for collection and delivery postcodes
        data['CollRegion'] = data['CollOutwardCode'].map(postcode_to_region).fillna("Unknown")
        data['DelRegion'] = data['DelOutwardCode'].map(postcode_to_region).fillna("Unknown")


        # Add a toggle checkbox to show/hide the dataset preview
        if st.checkbox("Show Dataset Preview", value=False): 
            st.subheader("Dataset Preview")
            st.dataframe(data, width=1800, height=400)  # Adjust dimensions if needed

        # Sidebar filters
        st.sidebar.title("Filter Options")

        # Collection region filter
        coll_regions = data['CollRegion'].unique().tolist()
        selected_coll_regions = st.sidebar.multiselect("Filter by Collection Region", options=coll_regions)
        if selected_coll_regions:
            data = data[data['CollRegion'].isin(selected_coll_regions)]

        # Delivery region filter
        del_regions = data['DelRegion'].unique().tolist()
        selected_del_regions = st.sidebar.multiselect("Filter by Delivery Region", options=del_regions)
        if selected_del_regions:
            data = data[data['DelRegion'].isin(selected_del_regions)]

        # Other filters
        filters = {}
        filter_order = ['Distance', 'StartDate', 'EndDate', 'AgreedDate', 'CustName', 'DelType']
        for column in filter_order:
            if column not in data.columns:
                continue

            if column == 'Distance':
                min_val, max_val = int(data[column].min()), int(data[column].max())
                if min_val == max_val:
                    # Handle case where min_val == max_val
                    st.sidebar.write(f"Only one value available for {column}: {min_val}")
                else:
                    # Add the slider when min_val and max_val are different
                    selected_range = st.sidebar.slider(
                        f"Filter {column}",
                        min_value=min_val,
                        max_value=max_val,
                        value=(min_val, max_val)
                    )
                    filters[column] = selected_range
            else:
                # Handle non-numeric columns with a multiselect
                unique_values = data[column].unique().tolist()
                selected_values = st.sidebar.multiselect(f"Filter {column}", options=unique_values)
                if selected_values:
                    filters[column] = selected_values

        # Apply other filters
        for column, filter_value in filters.items():
            if isinstance(filter_value, tuple):  # Slider filter
                data = data[(data[column] >= filter_value[0]) & (data[column] <= filter_value[1])]
            elif isinstance(filter_value, list):  # Dropdown filter
                data = data[data[column].isin(filter_value)]

        # Display filtered dataset
        st.markdown('<div class="center-content">', unsafe_allow_html=True)
        st.subheader("Filtered Dataset")
        st.dataframe(data, width=1800, height=800)
        st.markdown('</div>', unsafe_allow_html=True)

        # Export filtered data
        if st.button("Export Filtered Data"):
            try:
                buffer = io.BytesIO()
                data.to_csv(buffer, index=False)
                buffer.seek(0)
                st.download_button("Download Filtered Data", data=buffer, file_name="filtered_data.csv", mime="text/csv")
            except Exception as e:
                st.error(f"An error occurred while exporting the file: {e}")
                
        # Generate the map
        folium_map = create_folium_map(data)

        # Render the map
        st_folium(folium_map, width=1800, height=900)

if __name__ == "__main__":
    main()
