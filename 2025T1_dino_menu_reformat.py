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
    clean_name = re.sub(r'\([^)]*\)$', '', name.strip().rstrip('.').replace("'", "\\'"))  # Remove trailing brackets content
    clean_name = re.sub(r'\s+,', ',', clean_name)
    clean_name = re.sub(r',(?=\S)', ', ', clean_name)

    # Replace ampersands with exactly one space before and after them
    clean_name = re.sub(r'\s*&\s*', ' & ', clean_name)

    clean_name = clean_name.strip()  # Ensure spaces before and after the name are removed
    
    if not clean_name:
        return ''

    # Check for more than 5 consecutive spaces
    if re.search(r' {6,}', clean_name):
        print(f"Warning: More than 5 consecutive spaces found in Week {week_number}, Day {day}: {clean_name}")

    tags = parse_dietary_tags(name)
    if is_title_case(clean_name):
        tags['title'] = 'true'
    tags[special_type if is_special else meal_type] = 'true'
    tags_str = ', '.join([f'{k}: {v}' for k, v in tags.items()])

    return f"FoodItem(text: '{clean_name}', {tags_str}), "

def extract_menu_for_days(input_file, sheet_name, week_number):
    df = pd.read_excel(input_file, sheet_name=sheet_name)
    df = df.loc[:, ~df.columns.str.contains('^Unnamed')].reset_index(drop=True)

    days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
    all_output = []

    for day in days:
        day_data = df[df.columns[df.columns.str.contains(day, case=False)]].dropna().values.flatten()
        output = [f"final {day.lower()}W{week_number} = ["]

        for idx, meal in enumerate(['Breakfast', 'Lunch', 'Dinner']):
            output.append(f"    // {meal}")
            items = [item.strip().rstrip('.') for item in day_data[idx].split('\n') if item.strip()] if idx < len(day_data) else []

            for i, item_clean in enumerate(items):
                is_special = False
                special_type = None
                
                # Mark last non-empty item of Breakfast as brunch
                if meal == 'Breakfast' and i == len(items) - 1 and item_clean:
                    is_special = True
                    special_type = 'brunch'
                
                # Mark last non-empty item of Dinner as dessert
                if meal == 'Dinner' and i == len(items) - 1 and item_clean:
                    is_special = True
                    special_type = 'dessert'

                if is_special and not item_clean:
                    continue  # Skip empty items intended for deletion

                formatted_item = format_food_item(item_clean, meal.lower(), is_special, special_type, week_number, day)
                if formatted_item:
                    output.append(f"    {formatted_item}")

        output.append("];")
        all_output.append("\n".join(output))

    return "\n\n".join(all_output)

input_file = '2025T1.xlsx'
final_output = [
    extract_menu_for_days(input_file, sheet, week_number)
    for week_number, sheet in enumerate(['GH W1', 'GH W2', 'GH W3'], start=1)
]

pyperclip.copy("\n\n".join(final_output))
print("All sheets' output copied to clipboard.")