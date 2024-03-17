import pandas as pd
import requests
from bs4 import BeautifulSoup
import re

# Read Excel File
excel_file_path = r'C:\Users\anton\OneDrive\Desktop\Programming\linklist.xlsx'
output_file_path = r'C:\Users\anton\OneDrive\Desktop\Programming\linklistoutput.xlsx'

print("Reading Excel file...")
df = pd.read_excel(excel_file_path, header=None, names=['Link'])

# Check DataFrame
print("DataFrame after reading Excel file:")
print(df.head())

# Iterate through Links (limit to first 10 links for testing)
print("Verifying links (limiting to first 10 links for testing)...")
for index, row in df.head(10).iterrows():
    link = row['Link']

    # Make GET Request
    response = requests.get(link)

    # Check Response Status
    if response.status_code == 200:
        status = 'Active'
        # Record Result
        df.loc[index, 'Status'] = status
        print(f"Link {index + 1}: {status}")
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
                product_subcategory = re.search(r"connectedProductSubCategory:\s*'([^']*)'", script_content).group(1)
                group_id = re.search(r"connectedGroupId:\s*'([^']*)'", script_content).group(1)
                category_id = re.search(r"connectedProductCategoryId:\s*'([^']*)'", script_content).group(1)
            except Exception as e:
                print(f"Error extracting data for Link {index + 1}: {e}")

            # Append extracted data to adjacent columns
            df.at[index, 'Product Group'] = product_group
            df.at[index, 'Product Category'] = product_category
            df.at[index, 'Product Subcategory'] = product_subcategory
            df.at[index, 'Group ID'] = group_id
            df.at[index, 'Category ID'] = category_id

            print(f"Extracted Data for Link {index + 1}:")
            print(f"Product Group: {product_group}")
            print(f"Product Category: {product_category}")
            print(f"Product Subcategory: {product_subcategory}")
            print(f"Group ID: {group_id}")
            print(f"Category ID: {category_id}")

    else:
        status = 'Inactive'

# Save DataFrame
print("Saving DataFrame to Excel file...")
df.to_excel(output_file_path, index=False)
print(f"Output Excel file saved at: {output_file_path}")
