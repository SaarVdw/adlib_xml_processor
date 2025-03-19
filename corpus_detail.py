import xml.etree.ElementTree as ET
import pandas as pd

# Load and parse the XML file
export_adlib = r"C:\Users\SaarVandeweghe\PycharmProjects\registratiecheck\export_20240917.xml"

# Function to parse XML and extract required fields
def parse_xml(export_adlib):
    tree = ET.parse(export_adlib)
    root = tree.getroot()

    #list to store all records
    data = []

    # Loop over each 'record' element
    for record in root.findall('recordList/record'):
        record_data = {}

        # Extract the required fields
        record_data['priref'] = record.findtext('priref')

    return data

# Load the XML file and parse it
xml_file = export_adlib
parsed_data = parse_xml(xml_file)

df = pd.DataFrame(parsed_data)

print(df.head())

