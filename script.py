import xml.etree.ElementTree as ET
import pandas as pd
import html
import re

def extract_links(text):
    # Find all links in the string using a regular expression
    urls = re.findall(r'(https?://\S+)', text)
    # Return the first link found and the subsequent text
    if urls:
        link = urls[0]
        # Check if the link is at the beginning of the text
        link_index = text.find(link)
        if link_index == 0:
            remaining_text = text[len(link):].strip()
        else:
            remaining_text = text[:link_index].strip()
        return link, remaining_text
    else:
        return None, text.strip()


# Function to extract information from each Placemark
def extract_placemark_data(placemark):
    data = {}
    for child in placemark:
        tag = child.tag.split('}')[1]  # Remove the namespace
        text = child.text.strip() if child.text else None
        
        if tag == 'description':
            text = html.unescape(text)  # Decode HTML entities
            text = text.replace('<div>', '').replace('</div>', '')  # Remove <div> tags
            text = text.replace('<b>', '').replace('</b>', '')  # Remove <b> tags
            program_start = text.find('Program: ')
            company_start = text.find('Company: ')
            location_start = text.find('Location: ')
            more_info_start = text.find('More Info: ')
            
            data['Program'] = text[program_start + len('Program: '):company_start].strip()
            data['Company'] = text[company_start + len('Company: '):location_start].strip()
            data['Location'] = text[location_start + len('Location: '):more_info_start].strip()
            link, remaining_text = extract_links(text[more_info_start + len('More Info: '):])
            # Check if the remaining text is a link
            if remaining_text and not remaining_text.startswith('https://'):
                data['More Info'] = link
                data['U.S. Representative'] = remaining_text
            else:
                data['More Info'] = None
                data['U.S. Representative'] = None
        elif tag not in ['styleUrl', 'ExtendedData', 'Point']:
            data[tag] = text
        
    return data

# Path to the .kml file
file_path = './Artemis Partners.kml'

# Parsing the .kml file
tree = ET.parse(file_path)
root = tree.getroot()

# Extracting information from each Placemark
placemarks_data = []
for placemark in root.findall('.//{http://www.opengis.net/kml/2.2}Placemark'):
    placemarks_data.append(extract_placemark_data(placemark))

# Creating DataFrame with the extracted data
df = pd.DataFrame(placemarks_data)

# Saving the data to an Excel file
excel_file_path = 'artemis_data.xlsx'
df.to_excel(excel_file_path, index=False)

print(f'Data saved successfully to {excel_file_path}')
