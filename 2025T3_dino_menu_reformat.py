import pandas as pd
import os
import re
import pyperclip

def parse_dietary_tags(text):
    tags = {}
    if 'GF' in text:
        tags['gluten_free'] = 'true'
    if 'DF' in text:
        tags['dairy_free'] = 'true'
    if 'V ' in text or 'V,' in text:
        tags['veg'] = 'true'
    if 'Vegan' in text:
        tags['vegan'] = 'true'
    if 'Soy' in text:
        tags['soy'] = 'true'
    if 'override' in text.lower() or 'Sandwich Bar' in text:
        tags['override'] = 'true'
    return tags

def is_title_case(text):
    return text.endswith(':') or (len(text) > 3 and sum(1 for c in text if c.isupper()) / len(text) > 0.7)

def format_food_item(name, meal_type, is_special=False, special_type=None, week_number=None, day=None):
    # Skip items that are just a dash
    if name.strip() == '-':
        return ''
    
    # Remove trailing brackets content and escape single quotes
    clean_name = re.sub(r'\([^)]*\)$', '', name.strip().rstrip('.'))
    clean_name = clean_name.replace("'", "\\'")
    
    # Remove newline characters and replace with spaces
    clean_name = re.sub(r'\n+', ' ', clean_name)
    
    clean_name = re.sub(r'\s+,', ',', clean_name)
    clean_name = re.sub(r',(?=\S)', ', ', clean_name)
    # Replace ampersands with exactly one space before and after them
    clean_name = re.sub(r'\s*&\s*', ' & ', clean_name)
    # Normalize multiple spaces to single spaces
    clean_name = re.sub(r'\s+', ' ', clean_name)
    clean_name = clean_name.strip()
   
    if not clean_name:
        return ''
    
    # Check for more than 5 consecutive spaces (after normalization, this should be rare)
    if re.search(r' {6,}', clean_name):
        print(f"Warning: More than 5 consecutive spaces found in Week {week_number}, Day {day}: {clean_name}")
    
    tags = parse_dietary_tags(name)
    if is_title_case(clean_name):
        tags['title'] = 'true'
    
    tags[special_type if is_special else meal_type] = 'true'
    tags_str = ', '.join([f'{k}: {v}' for k, v in tags.items()])
    return f"FoodItem(text: '{clean_name}', {tags_str}), "

def find_meal_rows(df):
    """Find the row indices for BREAKFAST, LUNCH, and DINNER"""
    meal_rows = {}
    
    # Look through the first column for meal indicators
    for idx, cell in enumerate(df.iloc[:, 0]):
        if pd.notna(cell) and isinstance(cell, str):
            cell_upper = cell.upper().strip()
            if 'BREAKFAST' in cell_upper:
                meal_rows['breakfast'] = idx
            elif 'LUNCH' in cell_upper:
                meal_rows['lunch'] = idx
            elif 'DINNER' in cell_upper:
                meal_rows['dinner'] = idx
    
    return meal_rows

def get_day_columns(df):
    """Identify which columns correspond to which days"""
    days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
    day_columns = {}
    
    # Check header row and first few rows for day names
    for col_idx, col in enumerate(df.columns):
        col_str = str(col).strip()
        for day in days:
            if day.lower() in col_str.lower():
                day_columns[day] = col_idx
                break
    
    # If not found in headers, check first few rows
    if not day_columns:
        for row_idx in range(min(5, len(df))):
            for col_idx, cell in enumerate(df.iloc[row_idx]):
                if pd.notna(cell) and isinstance(cell, str):
                    cell_str = cell.strip()
                    for day in days:
                        if day.lower() in cell_str.lower():
                            day_columns[day] = col_idx
                            break
    
    return day_columns

def extract_meal_items(df, meal_row, day_columns):
    """Extract food items for a specific meal across all days"""
    meal_data = {}
    
    # Find the end of this meal section (next meal or end of data)
    next_meal_row = len(df)
    for idx in range(meal_row + 1, len(df)):
        cell = df.iloc[idx, 0]
        if pd.notna(cell) and isinstance(cell, str):
            cell_upper = cell.upper().strip()
            if any(meal in cell_upper for meal in ['BREAKFAST', 'LUNCH', 'DINNER']):
                next_meal_row = idx
                break
    
    # Extract items for each day
    for day, col_idx in day_columns.items():
        items = []
        # Look through rows from meal_row+1 to next_meal_row
        for row_idx in range(meal_row + 1, next_meal_row):
            if col_idx < len(df.columns):
                cell = df.iloc[row_idx, col_idx]
                if pd.notna(cell) and str(cell).strip():
                    items.append(str(cell).strip())
        meal_data[day] = items
    
    return meal_data

def extract_menu_for_days(input_file, sheet_name, week_number):
    df = pd.read_excel(input_file, sheet_name=sheet_name, header=None)
    
    # Find meal section rows
    meal_rows = find_meal_rows(df)
    
    # Find day columns
    day_columns = get_day_columns(df)
    
    if not meal_rows:
        print(f"Warning: Could not find meal sections in sheet {sheet_name}")
        return ""
    
    if not day_columns:
        print(f"Warning: Could not find day columns in sheet {sheet_name}")
        return ""
    
    print(f"Sheet {sheet_name}: Found meals at rows {meal_rows}")
    print(f"Sheet {sheet_name}: Found days at columns {day_columns}")
    
    days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
    all_output = []
    
    for day in days:
        if day not in day_columns:
            print(f"Warning: {day} not found in sheet {sheet_name}")
            continue
            
        output = [f"final {day.lower()}W{week_number} = ["]
        
        for meal_name in ['breakfast', 'lunch', 'dinner']:
            if meal_name not in meal_rows:
                print(f"Warning: {meal_name} not found in sheet {sheet_name}")
                continue
                
            output.append(f"    // {meal_name.capitalize()}")
            
            # Extract items for this meal and day
            meal_data = extract_meal_items(df, meal_rows[meal_name], day_columns)
            items = meal_data.get(day, [])
            
            for i, item in enumerate(items):
                if not item.strip():
                    continue
                    
                is_special = False
                special_type = None
                
                # Mark items as brunch only if they contain 'BRUNCH'
                if 'BRUNCH' in item.upper():
                    is_special = True
                    special_type = 'brunch'
                
                # Mark last non-empty item of Dinner as dessert
                if meal_name == 'dinner' and i == len(items) - 1 and item.strip():
                    is_special = True
                    special_type = 'dessert'
                
                formatted_item = format_food_item(item, meal_name, is_special, special_type, week_number, day)
                if formatted_item:
                    output.append(f"    {formatted_item}")
        
        output.append("];")
        all_output.append("\n".join(output))
    
    return "\n\n".join(all_output)

# Main execution
input_file = '2025T3.xlsx'

try:
    final_output = [
        extract_menu_for_days(input_file, sheet, week_number)
        for week_number, sheet in enumerate(['GH W1', 'GH W2', 'GH W3'], start=1)
    ]
    
    result = "\n\n".join(final_output)
    pyperclip.copy(result)
    print("All sheets' output copied to clipboard.")
    print("\nFirst few lines of output:")
    print("\n".join(result.split('\n')[:20]) + "...")
    
except Exception as e:
    print(f"Error processing file: {e}")
    import traceback
    traceback.print_exc()