import xml.etree.ElementTree as ET
import pandas as pd


def format_name(name):
    if "," in name:
        parts = name.split(", ")
        return " ".join(parts[1:] + [parts[0]])
    return name


def parse_xml(xml_file):
    tree = ET.parse(xml_file)
    root = tree.find(".//recordList")  # Ensuring we start at the correct root

    records = []

    for record in root.findall("record"):
        priref = record.find("priref").text if record.find("priref") is not None else ""

        # Extract names
        names = [format_name(name.text) for name in record.findall("naam") if name.text]

        # Extract statuses
        statuses = [status.text for status in record.findall("status") if status.text]

        # Extract qualifiers with their occurrences
        qualifier_dict = {}
        for qual in record.findall("kwalificatie"):
            if qual.text and qual.attrib.get("lang") == "nl-NL" and "occurrence" in qual.attrib:
                occurrence = int(qual.attrib["occurrence"]) - 1  # Convert to zero-based index
                qualifier_dict[occurrence] = qual.text

        # Separate names into "huidige toeschrijving" and "verworpen toeschrijving"
        huidige_toeschrijving_parts = []
        verworpen_toeschrijving_parts = []

        for i, name in enumerate(names):
            if i < len(statuses) and statuses[i] == "verworpen":
                verworpen_toeschrijving_parts.append(name)
            else:  # Default to "huidig" if not explicitly rejected
                if i in qualifier_dict:
                    huidige_toeschrijving_parts.append(f"{qualifier_dict[i]} {name}")
                else:
                    huidige_toeschrijving_parts.append(name)

        huidige_toeschrijving = " ".join(huidige_toeschrijving_parts)
        verworpen_toeschrijving = ", ".join(verworpen_toeschrijving_parts)

        records.append([priref, huidige_toeschrijving, verworpen_toeschrijving])

    return records


def save_to_excel(data, output_file):
    df = pd.DataFrame(data, columns=["priref", "huidige toeschrijving", "verworpen toeschrijving"])
    df.to_excel(output_file, index=False)


# Usage example
xml_file = "Project_RUB_vervaardigers.xml"  # Replace with your XML file path
output_file = "formatted_artists.xlsx"
data = parse_xml(xml_file)
save_to_excel(data, output_file)
print(f"Data saved to {output_file}")
