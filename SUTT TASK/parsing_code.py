import pandas as pd
import json

def parse_mess_menu(excel_path):
    # Read the Excel file into a DataFrame without headers
    df = pd.read_excel(excel_path, sheet_name='Sheet1', header=None)
    
    # Identify separator rows (rows where all cells are day names)
    day_names = {'SATURDAY', 'SUNDAY', 'MONDAY', 'TUESDAY', 'WEDNESDAY', 'THURSDAY', 'FRIDAY'}
    separator_indices = []
    for idx, row in df.iterrows():
        # Check if all non-null cells in the row are day names
        cells = [str(cell).strip().upper() for cell in row if pd.notna(cell)]
        if all(cell in day_names for cell in cells) and len(cells) > 0:
            separator_indices.append(idx)
    
    # Find the rows where each meal starts
    breakfast_start = lunch_start = dinner_start = None
    for idx, row in df.iterrows():
        for cell in row:
            cell_str = str(cell).strip().upper()
            if cell_str == 'BREAKFAST':
                breakfast_start = idx
            elif 'LUNCH' in cell_str:
                lunch_start = idx
            elif 'DINNER' in cell_str:
                dinner_start = idx
    
    # Function to find the next separator index after a given row
    def find_next_separator(start_row, separators):
        for s in sorted(separators):
            if s > start_row:
                return s
        return len(df)  # If no separator found, use the end of the DataFrame
    
    # Calculate end indices for each meal
    breakfast_end = find_next_separator(breakfast_start, separator_indices)
    lunch_end = find_next_separator(lunch_start, separator_indices)
    dinner_end = find_next_separator(dinner_start, separator_indices)
    
    # Generate row ranges for each meal
    breakfast_rows = range(breakfast_start + 1, breakfast_end)
    lunch_rows = range(lunch_start + 1, lunch_end)
    dinner_rows = range(dinner_start + 1, dinner_end)
    
    # Process each column (date) to collect meal items
    menu = {}
    for col in df.columns:
        # Extract and parse the date from row 1
        date_cell = df.iloc[1, col]
        try:
            date = pd.to_datetime(date_cell, errors='coerce').strftime('%Y-%m-%d')
            if not date:
                continue
        except:
            continue
        
        # Collect Breakfast items
        breakfast_items = []
        for row in breakfast_rows:
            item = df.iloc[row, col]
            if pd.notna(item) and '********' not in str(item):
                breakfast_items.append(str(item).strip())
        
        # Collect Lunch items
        lunch_items = []
        for row in lunch_rows:
            item = df.iloc[row, col]
            if pd.notna(item) and '********' not in str(item):
                lunch_items.append(str(item).strip())
        
        # Collect Dinner items
        dinner_items = []
        for row in dinner_rows:
            item = df.iloc[row, col]
            if pd.notna(item) and '********' not in str(item):
                dinner_items.append(str(item).strip())
        
        # Add to the menu dictionary
        menu[date] = {
            'Breakfast': breakfast_items,
            'Lunch': lunch_items,
            'Dinner': dinner_items
        }
    
    return menu

# Parse the menu and save to JSON
menu_data = parse_mess_menu('Mess Menu Sample.xlsx')
with open('menu.json', 'w') as json_file:
    json.dump(menu_data, json_file, indent=2)

print("JSON file generated successfully.")