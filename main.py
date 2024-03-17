import pandas as pd
import requests
from bs4 import BeautifulSoup
import re
from concurrent.futures import ThreadPoolExecutor

# Read Excel File
excel_file_path = r'C:\Users\anton\OneDrive\Desktop\Programming\linklist.xlsx'
output_file_path = r'C:\Users\anton\OneDrive\Desktop\Programming\linklistoutput.xlsx'

print("Reading Excel file...")
df = pd.read_excel(excel_file_path, header=None, names=['Link'])

# Check DataFrame
print("DataFrame after reading Excel file:")
print(df.head())


# Function to process a single link
# Function to process a single link
def process_link(link):
    idx = df.index[df['Link'] == link][0] + 1  # Get the index of the link in the DataFrame
    print(f"Processing link: {link} (row {idx})")
    status = 'Error'  # Default status
    product_group = None  # Default value for product_group
    product_category = None  # Default value for product_category
    product_subcategory = None  # Default value for product_subcategory
    group_id = None  # Default value for group_id
    category_id = None  # Default value for category_id

    try:
        # Make GET Request
        print(f"Sending GET request to: {link}")
        response = requests.get(link)

        # Check Response Status
        if response.status_code == 200:
            status = 'Active'
            print(f"Response status for link {link}: {status}")
            # Parse HTML content
            soup = BeautifulSoup(response.content, 'html.parser')

            # Extract relevant information
            script_tag = soup.find('script', string=lambda text: text and 'philips.context' in text)
            if script_tag:
                script_content = script_tag.string
                try:
                    # Extracting required data using regex
                    product_group = re.search(r"productGroup:\s*'([^']*)'", script_content).group(1)
                    product_category = re.search(r"productCategory:\s*'([^']*)'", script_content).group(1)
                    product_subcategory = re.search(r"productSubCategory:\s*'([^']*)'", script_content).group(1)
                    group_id = re.search(r"groupId:\s*'([^']*)'", script_content).group(1)
                    category_id = re.search(r"categoryId:\s*'([^']*)'", script_content).group(1)
                except AttributeError:
                    # If the extraction fails, try a different set of regular expressions
                    product_group = re.search(r"connectedGroup:\s*'([^']*)'", script_content).group(1)
                    product_category = re.search(r"connectedProductCategory:\s*'([^']*)'", script_content).group(1)
                    product_subcategory = re.search(r"connectedProductSubCategory:\s*'([^']*)'", script_content).group(
                        1)
                    group_id = re.search(r"connectedGroupId:\s*'([^']*)'", script_content).group(1)
                    category_id = re.search(r"connectedProductCategoryId:\s*'([^']*)'", script_content).group(1)
            else:
                status = 'Inactive'
                print(f"No 'philips.context' script found for link {link}")
    except Exception as e:
        print(f"Error processing link {link}: {e}")

    return link, status, product_group, product_category, product_subcategory, group_id, category_id


# Concurrently process links
print("Verifying links...")
with ThreadPoolExecutor(max_workers=24) as executor:  # Increase max_workers value for higher concurrency
    results = executor.map(process_link, df['Link'])

# Update DataFrame with results
for idx, result in enumerate(results, start=1):
    link, status, product_group, product_category, product_subcategory, group_id, category_id = result
    print(f"Updating DataFrame for link {link} (row {idx})...")
    df.loc[df['Link'] == link, 'Status'] = status
    df.loc[df['Link'] == link, 'Product Group'] = product_group
    df.loc[df['Link'] == link, 'Product Category'] = product_category
    df.loc[df['Link'] == link, 'Product Subcategory'] = product_subcategory
    df.loc[df['Link'] == link, 'Group ID'] = group_id
    df.loc[df['Link'] == link, 'Category ID'] = category_id

# Save DataFrame
print("Saving DataFrame to Excel file...")
df.to_excel(output_file_path, index=False)
print(f"Output Excel file saved at: {output_file_path}")
