
# Activate virtual env manually before running this script:
# source /Users/marlonkegel/Desktop/incite/post-neoliberalism/literature/.venv/bin/activate

from pyzotero import zotero
import pandas as pd
import os
import re

# Zotero configuration
library_id = '[ADD LIBRARY ID]'
library_type = 'group'
api_key = '[ADD API KEY]'
excel_filename = "/Users/marlonkegel/Desktop/incite/post-neoliberalism/literature/literature_RI.xlsx"

# Initialize Zotero client
zot = zotero.Zotero(library_id, library_type, api_key)

def create_initial_excel_file(filename):
    """
    Create a new Excel file with initial sheets and structure
    """
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        # Create ZoteroData sheet with predefined columns
        df_zotero_data = pd.DataFrame(columns=[
            'Zotero Key', 'Author(s)', 'Year', 'Title', 'Type',
            'Theme', 'Abstract', 'Publication/Publisher',
            'Tags', 'URL', 'Notes'
        ])
        df_zotero_data.to_excel(writer, sheet_name='ZoteroData', index=False)

        # Create Literature sheet with predefined columns
        df_literature = pd.DataFrame(columns=[
            'Zotero Key', 'Theme', 'Notes'
        ])
        df_literature.to_excel(writer, sheet_name='Literature', index=False)

    print("Created new Excel file.")

def sync_zotero_to_excel(excel_filename):
    """
    Sync Zotero library to Excel ZoteroData sheet
    """
# Fetch all items from Zotero
items = zot.everything(zot.items())
zotero_data_list = []  # Initialize the list before the loop

for item in items:
    d = item['data']
    # Skip attachments and notes
    if d.get('itemType', '') in ['attachment', 'note']:
        continue
    
    # Attempt to retrieve a normal URL field
    parent_url = d.get('url', '')

    # If no URL, check for a child attachment titled "Google Books Link"
    if not parent_url:
        child_attachments = zot.children(d.get('key'), itemType='attachment')
        for att in child_attachments:
            if att['data'].get('title', '') == 'Google Books Link':
                parent_url = att['data'].get('url', '')
                break

    # Process authors
    authors = '; '.join([
        f"{c.get('lastName', '')}, {c.get('firstName', '')}"
        for c in d.get('creators', [])
        if c.get('creatorType', '') == 'author' or 'contributor'
    ])

    # Process publication
    pub = (
        d.get('publicationTitle', '') or
        d.get('blogTitle', '') or
        d.get('websiteTitle', '') or
        d.get('publisher', '') or
        d.get('institution', '')
    )

    # Process automatic & manual tags (combine both)
    auto_tags = [t['tag'] for t in d.get('tags', []) if t.get('type', 1) == 1]
    manual_tags = [t['tag'] for t in d.get('tags', []) if t.get('type', 0) == 0]
    combined_tags = auto_tags + manual_tags
    tags_str = '; '.join(combined_tags)

    # Extract year from 'date' if possible
    from dateutil.parser import parse
    date_str = d.get('date', '')
    try:
        parsed_date = parse(date_str, fuzzy=True)
        year_str = parsed_date.year
    except (ValueError, TypeError):
        year_str = ''

    # Append row to list (leave Theme and Notes as empty strings)
    zotero_data_list.append({
        'Zotero Key': d.get('key', ''),
        'Author(s)': authors,
        'Year': year_str,
        'Title': d.get('title', ''),
        'Type': d.get('itemType', ''),
        'Abstract': d.get('abstractNote', ''),
        'Publication/Publisher': pub,
        'Tags': tags_str,
        'URL': parent_url
    })
# Create DataFrame with Zotero data
df_zotero = pd.DataFrame(zotero_data_list)

# Sort by last name of the first author
#  - isolate the first author (split by ';')
#  - then take the last name (split by ',')
df_zotero['SortKey'] = df_zotero['Author(s)'].apply(
    lambda x: x.split(';')[0].split(',')[0].strip() if isinstance(x, str) and x else ''
)
df_zotero.sort_values(by='SortKey', inplace=True)
df_zotero.drop(columns=['SortKey'], inplace=True)

# Replace ZoteroData sheet in the Excel file
with pd.ExcelWriter(excel_filename, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    df_zotero.to_excel(writer, sheet_name='ZoteroData', index=False)

print("Zotero library synced to Excel successfully.")

def main():
    """
    Main function to manage the workflow:
    - Checks if the Excel file exists and creates it if not.
    - Syncs the Zotero library data to the Excel file.
    """
    # Check if Excel file exists, create if not
    if not os.path.exists(excel_filename):
        create_initial_excel_file(excel_filename)
    
    # Only sync from Zotero --> Excel
    sync_zotero_to_excel(excel_filename)

if __name__ == "__main__":
    main()
