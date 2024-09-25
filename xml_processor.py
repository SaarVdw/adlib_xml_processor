import tkinter
import xml.etree.ElementTree as ET
import pandas as pd
import os
from tkinter import Tk, Label, Button, filedialog, messagebox, Image, PhotoImage
from shutil import copyfile
from PIL import ImageTk, Image
from pefile import sizeof_type

#kleuren importeren
donkerblauw = '#%02x%02x%02x' % (11, 38, 62)
middenblauw = '#%02x%02x%02x' % (133, 193, 218)
lichtblauw = '#%02x%02x%02x' % (229, 240, 245)


# Function to parse XML and extract required fields
def parse_xml(xml_file):
    tree = ET.parse(xml_file)
    root = tree.getroot()

    # List to store all records
    data = []

    # Loop over each 'record' element
    for record in root.findall('recordList/record'):
        record_data = {}

        # Extract the required fields
        record_data['priref'] = record.findtext('priref')
        record_data['afbeeldingsnummer'] = record.findtext('afbeeldingsnummer')

        # Extract objectcategorie in Dutch (nl-NL)
        objectcategorie_elements = record.findall("objectcategorie[@lang='nl-NL']")
        record_data['objectcategorie'] = objectcategorie_elements[0].text if objectcategorie_elements else None

        # Other fields
        record_data['benaming_kunstwerk'] = record.findtext('benaming_kunstwerk')
        record_data['titel_engels'] = record.findtext('titel_engels')

        # Extract collectienaam - take the last value or "veiling" if empty
        collectienaam_elements = record.findall('collectienaam')
        if collectienaam_elements:
            record_data['collectienaam'] = collectienaam_elements[-1].text  # Last value
        else:
            record_data['collectienaam'] = 'veiling'  # Default value

        # Extract genre in Dutch (nl-NL)
        genre_elements = record.findall("genre[@lang='nl-NL']")
        record_data['genre'] = genre_elements[0].text if genre_elements else None

        # Extract keywords_nl in Dutch (nl-NL) (RKD_algemene_trefwoorden)
        keywords_elements = record.findall("RKD_algemene_trefwoorden[@lang='nl-NL']")
        record_data['keywords_nl'] = ', '.join([kw.text for kw in keywords_elements]) if keywords_elements else None

        # Other fields that are not language-specific
        record_data['datering_engels'] = record.findtext('datering_engels')
        record_data['eenheid'] = record.findtext('eenheid')
        record_data['drager'] = record.findtext("drager[@lang='nl-NL']")  # Extract drager in Dutch
        record_data['iconclass'] = record.findtext('iconclass_code')

        # Extract material in Dutch (nl-NL)
        material_elements = record.findall("materiaal[@lang='nl-NL']")
        record_data['materiaal'] = ', '.join([mat.text for mat in material_elements]) if material_elements else None

        record_data['project'] = ', '.join([proj.text for proj in record.findall("project[@lang='nl-NL']")])
        record_data['archiefreferentie'] = record.findtext('archiefreferentie')

        # Append the record data to the list
        data.append(record_data)

    return data

def process_file(xml_file):
    # Parse the XML file and proceed with processing
    parsed_data = parse_xml(xml_file)

    # Convert parsed data into a DataFrame for analysis
    df = pd.DataFrame(parsed_data)

    df['priref'] = df['priref'].astype(str)

    # Adding data from external files
    # Collectienamen
    collectienamen = r"C:\Users\SaarVandeweghe\PycharmProjects\registratiecheck\collectienamen_CRLB_20240912.xlsx"
    collectienamen_df = pd.read_excel(collectienamen)

    # Merging dataframes
    merged_df = pd.merge(df, collectienamen_df, on='collectienaam', how='left')

    # Toeschrijvingen
    toeschrijvingen = r"C:\Users\SaarVandeweghe\PycharmProjects\registratiecheck\toeschrijvingen.xlsx"
    toeschrijvingen_df = pd.read_excel(toeschrijvingen)

    toeschrijvingen_df['priref'] = toeschrijvingen_df['priref'].astype(str)

    # Merging full df
    full_df = pd.merge(left=merged_df, right=toeschrijvingen_df, on='priref', how='left')

    # Renaming columns if necessary
    full_df.rename(columns={
        'priref': 'kunstwerknummer',
        'keywords_nl': 'iconografie',
        'objectcategorie_nl': 'objectcategorie',
        'benaming_kunstwerk': 'titel',
        'materiaal_nl': 'materiaal',
        'genre_nl': 'genre',
        'toeschrijving': 'huidige toeschrijving',
        'corpusdeel': 'CRLB',
        'corpusnummer': '#',
        'iconclass_code': 'iconclass',
        'plaats': 'plaats collectie',
        'type': 'type collectie',
    }, inplace=True)

    # Reorganizing columns
    full_df = full_df.reindex(columns=['kunstwerknummer', 'CRLB', '#', 'afbeeldingsnummer', 'objectcategorie', 'huidige toeschrijving', 'verworpen toeschrijving', 'titel', 'titel_engels', 'collectienaam', 'type collectie', 'plaats collectie', 'genre', 'iconografie', 'iconclass', 'datering_engels', 'eenheid', 'drager', 'materiaal', 'project', 'archiefreferentie'])

    # Step 3: Extract all terms from the 'iconografie' column by splitting on commas and flattening
    all_iconografie_terms = full_df['iconografie'].dropna().str.split(',').explode().str.strip()

    # Step 4: Get the count of each unique term
    iconografie_counts = all_iconografie_terms.value_counts().reset_index()
    iconografie_counts.columns = ['unieke termen', 'aantal']

    # Step 5: Save both the full_df and iconografie_counts to the same Excel file
    output_excel_file = 'filtered_xml_data_analysis.xlsx'

    with pd.ExcelWriter(output_excel_file, engine='openpyxl', mode='w') as writer:
        # Write the main DataFrame to the first sheet
        full_df.to_excel(writer, sheet_name='Filtered Data', index=False)

        # Add a new sheet with the unique iconografie terms and their counts
        iconografie_counts.to_excel(writer, sheet_name='Iconografie', index=False)

    messagebox.showinfo("Success", f"De data is omgezet en werd opgeslagen in {output_excel_file}.")

def open_file_dialog():
    # Step 1: Open a file dialog for XML selection
    xml_file = filedialog.askopenfilename(
        title="Select XML File",
        filetypes=(("XML files", "*.xml"), ("All files", "*.*"))
    )

    if xml_file:
        # Step 2: Process the selected file
        process_file(xml_file)
    else:
        messagebox.showerror("Error", "No file selected.")

def download_instructions():
    # Path to the instructions/sample XML file
    instructions_file = r"C:\Users\SaarVandeweghe\PycharmProjects\registratiecheck\definitie_export_CRLB.xml"

    # Ask the user where they want to save the file
    save_path = filedialog.asksaveasfilename(
        title="Save Instructions",
        defaultextension=".xml",
        filetypes=[("XML files", "*.xml"), ("All files", "*.*")]
    )

    if save_path:
        # Copy the instructions file to the chosen location
        try:
            copyfile(instructions_file, save_path)
            messagebox.showinfo("Success", "Instructions file saved successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save file: {e}")
    else:
        messagebox.showerror("Error", "Save location not selected.")

def resource_path(relative_path):
    """ Get the absolute path to a resource, works for both development and PyInstaller bundle """
    try:
        # PyInstaller creates a temp folder and stores the path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def create_gui():
    # Create the main window
    root = Tk()
    root.title("XML Processing Tool")
    root.geometry("600x500")
    root.configure(bg=donkerblauw)

    # Welcome message
    welcome_label = Label(root, text="XML Processing Tool", font=("Poppins ExtraBold", 20), bg=donkerblauw, fg=lichtblauw)
    welcome_label.pack(pady=20)

    # Explanation message
    explanation_label = Label(root, text="Zet snel een XML export uit Adlib om naar een leesbaar excel-bestand. Gebruik het definitiebestand zodat je de juiste velden exporteert.",font="Poppins", wraplength=450,bg=donkerblauw, fg=middenblauw)
    explanation_label.pack(pady=10)

    # File selection button
    select_button = Button(root, text="Select XML file", command=open_file_dialog, font=("Poppins SemiBold", 12), width=20)
    select_button.pack(pady=20)

    # Download instructions button
    download_button = Button(root, text="Download definitiebestand", command=download_instructions, font=("Poppins", 12), width=30, bg=middenblauw, fg=donkerblauw)
    download_button.pack(pady=10)

    # Load the image using the resource_path function
    image_path = resource_path("Rubenshuis_Logo_2_Middenblauw_2905_2400.png")
    image = Image.open(image_path)
    image = ImageTk.PhotoImage(image)

    panel = Label(root, image = image, borderwidth=0,compound="center",highlightthickness = 0,padx=0,pady=0)
    panel.image = image
    panel.pack(side = "left", expand = "yes")

    # Start the Tkinter loop
    root.mainloop()

if __name__ == "__main__":
    create_gui()

