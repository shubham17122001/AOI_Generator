import os
import pandas as pd
import simplekml
import zipfile
import streamlit as st
from shapely.geometry import Polygon
import numpy as np
import base64
import geopandas as gpd  # Import geopandas

# Function to convert km to degrees
def km_to_degrees(lat, km, is_latitude=True):
    """Convert km to degrees for latitude or longitude."""
    if is_latitude:
        return km / 111  # 1 degree of latitude is approximately 111 km
    else:
        # Longitude distance varies with latitude, so we adjust based on cosine of latitude
        return km / (111 * abs(np.cos(np.radians(lat))))  # Convert km to degrees of longitude

# Function to create a rectangle around the center point
def create_rectangle(lat, lon, width_km, height_km):
    # Calculate the degree offsets for the rectangle dimensions
    lat_offset = km_to_degrees(lat, height_km / 2, is_latitude=True)
    lon_offset = km_to_degrees(lat, width_km / 2, is_latitude=False)

    # Define the 4 corners of the rectangle (N, S, E, W)
    coordinates = [
        (lon - lon_offset, lat + lat_offset),  # Top-left
        (lon + lon_offset, lat + lat_offset),  # Top-right
        (lon + lon_offset, lat - lat_offset),  # Bottom-right
        (lon - lon_offset, lat - lat_offset),  # Bottom-left
    ]

    # Create a Polygon object for the rectangle
    return Polygon(coordinates)

# Function to generate KML file
def generate_kml(df, aoi_width_km, aoi_height_km, file_name):
    kml = simplekml.Kml()

    # Create a directory for saving the files (if it doesn't exist)
    save_dir = "./AOI_Files"
    if not os.path.exists(save_dir):
        os.makedirs(save_dir)

    # Create a list to store the polygon geometries for shapefile
    polygons = []

    # Loop through each center point to create AOIs
    for _, row in df.iterrows():
        code = row['CODE']
        lat = row['CENTER LAT']
        lon = row['CENTER LONG']

        # Create the rectangle (AOI)
        rectangle = create_rectangle(lat, lon, aoi_width_km, aoi_height_km)

        # Convert the rectangle's coordinates into a list for KML
        coords = [(lon, lat) for lon, lat in rectangle.exterior.coords]

        # Add the rectangle to KML with the code as the name
        pol = kml.newpolygon(name=code, outerboundaryis=coords)
        pol.style.linestyle.color = simplekml.Color.blue
        pol.style.linestyle.width = 4
        pol.style.polystyle.color = simplekml.Color.changealphaint(100, simplekml.Color.blue)

        # Mark the center point on the KML (as a placemark)
        point_name = f"Center: {code},\nLat: {lat}, Lon: {lon}"
        point = kml.newpoint(name=point_name, coords=[(lon, lat)])

        # Create an Icon object for the center point marker
        point.iconstyle.icon = simplekml.Icon()
        point.iconstyle.icon.href = 'http://maps.google.com/mapfiles/kml/pal2/icon10.png'
        point.style.iconstyle.scale = 1

        # Append the polygon to the list for shapefile creation
        polygons.append({'CODE': code, 'geometry': rectangle})

    # Save the KML file in the specified directory with dynamic file name
    kml_path = os.path.join(save_dir, f"{file_name}_AOIs_with_centers.kml")
    kml.save(kml_path)

    return kml_path, polygons

# Function to create KMZ file from KML
def create_kmz(kml_path):
    # Save the KMZ file in the specified directory
    kmz_path = os.path.join(os.path.dirname(kml_path), f"{os.path.basename(kml_path).replace('.kml', '.kmz')}")
    with zipfile.ZipFile(kmz_path, 'w', zipfile.ZIP_DEFLATED) as kmz:
        kmz.write(kml_path, os.path.basename(kml_path))
    return kmz_path

# Function to generate shapefile
def generate_shapefile(polygons, file_name):
    # Create a GeoDataFrame from the polygons
    gdf = gpd.GeoDataFrame(polygons)
    gdf.set_index('CODE', inplace=True)

    # Define the CRS (Coordinate Reference System), WGS84 (EPSG:4326)
    gdf.crs = 'EPSG:4326'

    # Save the shapefile
    shapefile_dir = './AOI_Files'
    shapefile_path = os.path.join(shapefile_dir, file_name)
    gdf.to_file(shapefile_path + '.shp')

    return shapefile_path

# Streamlit app UI

# Function to convert image to base64
def image_to_base64(image_path):
    with open(image_path, "rb") as img_file:
        return base64.b64encode(img_file.read()).decode("utf-8")

# Convert image to base64 (ensure the correct path)
image_base64 = image_to_base64("C:/Users/RSDSOffice/Downloads/Antrix_logo.png")  # Update the path to your image

# HTML to embed the image above the title
st.markdown(f"""
    <style>
        .top-center-logo {{
            position: absolute;
            top: 30px;  /* Adjusted to give space between logo and title */
            left: 50%;
            transform: translateX(-50%);
        }}
        .title {{
            font-size: 30px;
            font-weight: bold;
            white-space: nowrap;
            text-align: center;
            margin-top: 120px;  /* Adjusted to give space between logo and title */
        }}
    </style>
    <div class="top-center-logo">
        <img src="data:image/png;base64,{image_base64}" width="120">
    </div>
    <div class="title">AOI Generator & KMZ/Shapefile Download</div>
""", unsafe_allow_html=True)

st.write("This app allows you to upload an Excel file, specify AOI dimensions, generate and download KMZ and Shapefile.")

# File Upload Section
uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx"])

if uploaded_file:
    # Read the Excel file
    df = pd.read_excel(uploaded_file, skiprows=1)

    # Extracting the file name without extension for KMZ/KML file naming
    file_name = uploaded_file.name.split('.')[0]

    # Display all data from the Excel file
    st.subheader(f"Data from {uploaded_file.name}")
    st.write(df)

    # User input for AOI size
    st.sidebar.header("AOI Dimensions")
    aoi_width_km = st.sidebar.number_input("Enter AOI Width (in km):", min_value=1, max_value=100, value=8)
    aoi_height_km = st.sidebar.number_input("Enter AOI Height (in km):", min_value=1, max_value=100, value=8)
    
    st.sidebar.text("Set AOI dimensions and click the button to generate the KMZ and Shapefile.")

    # Button to generate KMZ and Shapefile
    if st.sidebar.button("Generate Files"):
        with st.spinner('Generating KMZ and Shapefile...'):
            # Generate KML file and polygon data
            kml_path, polygons = generate_kml(df, aoi_width_km, aoi_height_km, file_name)

            # Create KMZ file
            kmz_path = create_kmz(kml_path)

            # Generate shapefile
            shapefile_path = generate_shapefile(polygons, file_name)

            # Provide download button for KMZ file
            with open(kmz_path, "rb") as kmz_file:
                st.sidebar.download_button("Download KMZ", kmz_file, file_name=f"{file_name}_AOIs_with_centers.kmz")

            # Provide download button for shapefile
            shapefile_zip_path = shapefile_path + '.zip'
            with zipfile.ZipFile(shapefile_zip_path, 'w') as shapefile_zip:
                for ext in ['shp', 'shx', 'dbf', 'prj']:
                    shapefile_zip.write(f"{shapefile_path}.{ext}", os.path.basename(f"{shapefile_path}.{ext}"))

            with open(shapefile_zip_path, "rb") as shapefile_zip_file:
                st.sidebar.download_button("Download Shapefile", shapefile_zip_file, file_name=f"{file_name}_AOIs.zip")

            st.sidebar.success(f"KMZ and Shapefile have been generated!")
