from pyairtable.formulas import match
from pyairtable import Table
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from datetime import datetime
from dotenv import load_dotenv
import streamlit as st
from io import BytesIO
import logging
import ast
import json
import openai
import os

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


class AirTableError(Exception):
    """Custom exception for AirTable operations"""
    pass

class ToMakeSheetGenerator():
    def __init__(self, external_api_key=None):
        """Initialize AirTable connection with API key and table definitions"""
        try:
            load_dotenv()
            self.api_key = external_api_key or st.secrets["AIRTABLE_API_KEY"]
            self.base_id = 'appEe646yuQexwHJo'
            
            # Initialize tables
            self.clientserving_table = Table(self.api_key, self.base_id, 'tblVwpvUmsTS2Se51')
            self.ingredients_table = Table(self.api_key, self.base_id, 'tblPhcO06ce4VcAPD')
            
            # Cache for conversion factors to avoid repeated API calls
            self.conversion_cache = {}
            
        except Exception as e:
            logger.error(f"Failed to initialize AirTable connection: {str(e)}")
            raise AirTableError(f"Failed to initialize AirTable connection: {str(e)}")

    def get_clientservings_data(self):
        """Get all client servings data with Modified Recipe Details and component columns"""
        # Field IDs for ClientServings table
        CUSTOMER_NAME = 'fldDs6QXE6uYKIlRk'
        DELIVERY_DATE = 'fld6YLCoS7XCFK04G'
        DISH = 'fldmqHv4aXJxuJ8E2'
        MODIFIED_RECIPE_DETAILS = 'fld7dTxR7ImOtzO4k'
        STARCH = 'fldZKmvAeBTY9tRkR'  # Starch ingredients
        PROTEIN = 'fldksE3QaxIIHzAIi'  # Protein/Meat ingredients  
        SAUCE = 'fldOb13TV0bymcyF6'  # Sauce ingredients
        GARNISH = 'fldgrNa89SJKSPtwY'  # Garnish ingredients
        VEGGIE = 'fldHaAZr6fiWZbNaf'  # Veggie ingredients
        MEAL = 'fld00H3SpKNqTbhC0'
        MEAL_FROM_CLIENT = "fld00H3SpKNqTbhC0"
        
        fields_to_return = [
            CUSTOMER_NAME, DELIVERY_DATE, DISH,
            MODIFIED_RECIPE_DETAILS,
            STARCH, PROTEIN, SAUCE, GARNISH, VEGGIE, MEAL, MEAL_FROM_CLIENT
        ]
        
        try:
            # Get all client servings from the view
            # Note: Modified Recipe Details will be included by default since we're not using field filtering
            records = self.clientserving_table.all(view='viw4WN1XsjMnHwMkt',fields=fields_to_return)
            logger.info(f"Retrieved {len(records)} records from ClientServings table")
            return records
        except Exception as e:
            logger.error(f"Failed to get client servings data: {str(e)}")
            raise AirTableError(f"Failed to get client servings data: {str(e)}")

    def get_ingredient_conversion_factor(self, ingredient_name):
        """Get the raw/cooked conversion factor for an ingredient by name with caching"""
        # Check cache first
        if ingredient_name in self.conversion_cache:
            return self.conversion_cache[ingredient_name]
            
        try:
            # Field IDs for Ingredients table
            INGREDIENT_ID = 'flduR79GRxTbyfyKe'
            RAW_COOKED_CONVERSION = 'fld8VHWjNgPuj5szl' #Cooked/Raw Conversion

            formula = dict()
            formula[INGREDIENT_ID] = ingredient_name
            formula = match(formula)
            # Search for ingredient by name
            records = self.ingredients_table.all(formula=formula)
            
            if records:
                conversion_factor = records[0]['fields'].get("Cooked/Raw Conversion", 1.0)
                result = float(conversion_factor) if conversion_factor else 1.0
                
            else:
                logger.warning(f"Ingredient '{ingredient_name}' not found in ingredients table")
                result = 1.0
            
            # Cache the result
            self.conversion_cache[ingredient_name] = result
            return result
                
        except Exception as e:
            logger.warning(f"Error getting conversion factor for '{ingredient_name}': {str(e)}")
            # Cache the default value
            self.conversion_cache[ingredient_name] = 1.0
            return 1.0

    def is_breakfast_meat(self, meal_type):
        """Determine if a meat ingredient is for breakfast based on meal type from MEAL field"""
        if not meal_type:
            return False
        return 'breakfast' in meal_type.lower()
    

    def parse_sauce_multiplier(self, sauce_field):
        """Parse sauce multiplier from sauce field (e.g., 'Tomato Sauce (2 x sauce)' -> 2)"""
        import re
        
        # Look for pattern like "(2 x sauce)", "(3 x sauce)", etc.
        multiplier_match = re.search(r'\((\d+)\s*x\s*sauce\)', sauce_field, re.IGNORECASE)
        if multiplier_match:
            return int(multiplier_match.group(1))
        
        # If no multiplier found, default to 1
        return 1

    def parse_modified_recipe_details(self, modified_recipe_text):
        """Parse the Modified Recipe Details field which contains ingredient names and gram amounts"""
        if not modified_recipe_text:
            return {}
            
        try:
            # Try to parse as JSON first
            if isinstance(modified_recipe_text, str):
                # Handle both JSON and Python dict string formats
                if modified_recipe_text.startswith('{'):
                    try:
                        return json.loads(modified_recipe_text)
                    except json.JSONDecodeError:
                        # Try parsing as Python literal
                        return ast.literal_eval(modified_recipe_text)
                else:
                    return {}
            elif isinstance(modified_recipe_text, dict):
                return modified_recipe_text
            else:
                return {}
        except (ValueError, SyntaxError) as e:
            logger.warning(f"Could not parse Modified Recipe Details: {modified_recipe_text}, Error: {e}")
            return {}

    def identify_fruits_with_openai(self, ingredient_names):
        """Use OpenAI to identify which ingredients are fruits"""
        try:
            # Get OpenAI API key from environment or streamlit secrets
            openai_api_key = st.secrets.get("OPENAI_API_KEY") if hasattr(st, 'secrets') else None
            if not openai_api_key:
                openai_api_key = os.getenv("OPENAI_API_KEY")
            
            if not openai_api_key:
                logger.warning("OpenAI API key not found, skipping fruit identification")
                return set()
            
            openai.api_key = openai_api_key
            
            # Create prompt for fruit identification
            prompt = f"""
            Given the following list of food ingredients, identify which ones are fruits.
            Return only the fruit names as a comma-separated list, nothing else. Dont change any ingredient names.
            
            Ingredients: {', '.join(ingredient_names)}
            
            Fruits:"""
            
            response = openai.ChatCompletion.create(
                model="gpt-o4-mini",
                messages=[
                    {"role": "system", "content": "You are a helpful assistant that identifies fruits from food ingredient lists. Only respond with the fruit names separated by commas."},
                    {"role": "user", "content": prompt}
                ],
                max_tokens=100,
                temperature=0
            )
            
            fruit_response = response.choices[0].message.content.strip()
            fruits = {fruit.strip() for fruit in fruit_response.split(',') if fruit.strip()}
            
            logger.info(f"Identified fruits: {fruits}")
            return fruits
            
        except Exception as e:
            logger.warning(f"Error identifying fruits with OpenAI: {str(e)}")
            return set()
    

    def group_ingredients_by_component(self, client_servings):
        """Group all ingredients by component type from client servings using Modified Recipe Details"""
        ingredient_summary = {
            'Starch': {},
            'Breakfast Meat': {},  # Separate section for breakfast meats
            'Meat': {},  # Lunch/dinner meats
            'Sauce': {},
            'Garnish': {},  # Will store garnish combinations and their meal counts
            'Veggie': {},
            'Snack': {}  # New section for snacks
        }
        
        # Special handling for garnish combinations
        garnish_combinations = {}  # Will store unique combinations and their meal counts
        
        # Special handling for snacks
        snack_dishes = {}  # Will store snack dishes and their total grams
        snack_ingredients = set()  # Track all ingredients that come from snacks
        
        
        for serving in client_servings:
            fields = serving['fields']
            
            # Get meal type from MEAL field
            meal_field = fields.get('Meal Portion (from Linked OrderItem)', '')
            
            if isinstance(meal_field, list):
                meal_type = meal_field[0] if meal_field else ''
            else:
                meal_type = meal_field if meal_field else ''
            
            if 'pre-set' in meal_type.lower() or 'add on' in meal_type.lower():
                meal_type = fields.get('Meal Portion (from Linked OrderItem)', '')[0]

            is_snack = 'snack' in meal_type.lower()
            
            # Parse Modified Recipe Details to get actual ingredients and their gram amounts
            modified_recipe = self.parse_modified_recipe_details(fields.get('Modified Recipe Details', ''))
            
            if not modified_recipe:
                continue
                
            # Get ingredient names from each component column to determine component mapping
            starch_names = fields.get('Starch', '').replace('\n', '').split(', ')
            meat_names = fields.get('Meat', '').replace('\n', '').split(', ')
            sauce_field = fields.get('Sauce', '')
            sauce_names = sauce_field.replace('\n', '').replace(' (2 x sauce)', '').replace(' (3 x sauce)', '').replace(' (4 x sauce)', '').split(', ')
            garnish_names = fields.get('Garnish', '').replace('\n', '').split(', ')
            veggie_names = fields.get('Veggies', '').replace('\n', '').split(', ')
            
            # Parse sauce multiplier from the original sauce field
            sauce_multiplier = self.parse_sauce_multiplier(sauce_field)
            
            # Create a mapping of ingredient name to component type
            component_mapping = {}
            for name in starch_names:
                if name != '':
                    component_mapping[name.strip()] = 'Starch'
            for name in meat_names:
                if name != '':
                    component_mapping[name.strip()] = 'Meat'
            for name in sauce_names:
                if name != '':
                    component_mapping[name.strip()] = 'Sauce'
            for name in garnish_names:
                if name != '':
                    component_mapping[name.strip()] = 'Garnish'
            for name in veggie_names:
                if name != '':
                    component_mapping[name.strip()] = 'Veggie'
            
            # Get the dish name
            dish_name = fields.get('Dish', ['Unknown'])[0] if fields.get('Dish') else 'Unknown'

            # Process each ingredient in Modified Recipe Details
            total_grams = 0
            for ingredient_name, grams in modified_recipe.items():
                
                # Remove ID prefix from ingredient name (e.g. "16137 Hummus" -> "Hummus")
                clean_ingredient_name = ' '.join(ingredient_name.split(' ')[1:]).strip()
                
                # Track ingredients that come from snacks
                if is_snack:
                    if dish_name not in snack_dishes:
                        snack_dishes[dish_name] = 0
                    snack_dishes[dish_name] += grams
                    continue
                
                # Determine component type based on mapping
                component_type = component_mapping.get(clean_ingredient_name)
                
                # Debug: Log component mapping results
                # if clean_ingredient_name == 'Carrot':
                #     logger.info(f"Processing ingredient: '{clean_ingredient_name}' -> component: '{component_type}', grams: {grams}")
                
                if not component_type:
                    logger.warning(f"Could not determine component type for ingredient: '{clean_ingredient_name}'")
                    logger.info(f"Available component mappings: {list(component_mapping.keys())}")
                    continue
                
                if component_type == 'Garnish':
                    # For garnish, we'll handle it separately
                    continue
                
                grams_float = float(grams)
                
                # Apply raw/cooked conversion factor
                conversion_factor = self.get_ingredient_conversion_factor(ingredient_name.strip())
                final_grams = grams_float / conversion_factor if conversion_factor != 0 else grams_float
                
                total_grams += final_grams
                
                # Special handling for meat - exclude snack meats and group by meal type
                if component_type == 'Meat':
                    
                    # Determine if this is breakfast meat
                    is_breakfast = self.is_breakfast_meat(meal_type)
                    target_section = 'Breakfast Meat' if is_breakfast else 'Meat'
                    
                    if clean_ingredient_name not in ingredient_summary[target_section]:
                        ingredient_summary[target_section][clean_ingredient_name] = {
                            'total_grams': 0
                        }
                    
                    ingredient_summary[target_section][clean_ingredient_name]['total_grams'] += final_grams
                    
                else:
                    # Special handling for sauce - count quantities instead of grams
                    if component_type == 'Sauce':
                        if clean_ingredient_name not in ingredient_summary[component_type]:
                            ingredient_summary[component_type][clean_ingredient_name] = {
                                'total_count': 0
                            }
                        
                        # Add the sauce multiplier count (from "(2 x sauce)" notation)
                        ingredient_summary[component_type][clean_ingredient_name]['total_count'] += sauce_multiplier
                    
                    # Handle all other ingredient types without grouping
                    else:
                        # Skip ingredients from snacks for components other than garnish and sauce
                        if component_type not in ['Garnish', 'Sauce'] and is_snack:
                            continue
                            
                        if clean_ingredient_name not in ingredient_summary[component_type]:
                            ingredient_summary[component_type][clean_ingredient_name] = {
                                'total_grams': 0
                            }
                        
                        ingredient_summary[component_type][clean_ingredient_name]['total_grams'] += final_grams
                            
            # Handle garnish combinations
            if garnish_names and any(name.strip() for name in garnish_names):
                if dish_name not in garnish_combinations:
                    garnish_combinations[dish_name] = {
                        'garnishes': set(),
                        'meal_count': 0
                    }
                garnish_combinations[dish_name]['garnishes'].update(name.strip() for name in garnish_names if name.strip())
                # if dish_name == 'Poke Bowl':
                #     print(f"Garnishes for {serving['id']}: {garnish_names}")
                #     print(f"Garnishes for {dish_name}: {garnish_combinations[dish_name]['garnishes']}")
                garnish_combinations[dish_name]['meal_count'] += 1
           
            
            
        
        # Add garnish combinations to the summary
        ingredient_summary['Garnish'] = {
            dish: {
                'garnish_combo': ', '.join(sorted(garnish_data['garnishes'])),
                'meal_count': garnish_data['meal_count']
            }
            for dish, garnish_data in garnish_combinations.items()
        }
        
        # Add snack dishes to the summary
        ingredient_summary['Snack'] = {
            dish: {'total_grams': grams}
            for dish, grams in snack_dishes.items()
        }
        
        # Consolidate duplicate ingredients between Breakfast Meat and Meat sections
        breakfast_meats = ingredient_summary['Breakfast Meat']
        lunch_dinner_meats = ingredient_summary['Meat']
        
        # Find duplicates and consolidate them into Breakfast Meat
        duplicates_to_remove = []
        for ingredient_name in lunch_dinner_meats:
            if ingredient_name in breakfast_meats:
                # Add the lunch/dinner meat quantity to breakfast meat
                breakfast_meats[ingredient_name]['total_grams'] += lunch_dinner_meats[ingredient_name]['total_grams']
                # Mark for removal from lunch/dinner meat
                duplicates_to_remove.append(ingredient_name)
                logger.info(f"Consolidated duplicate ingredient '{ingredient_name}' into Breakfast Meat section")
        
        # Remove duplicates from lunch/dinner meat section
        for ingredient_name in duplicates_to_remove:
            del lunch_dinner_meats[ingredient_name]
        
        # Debug: Log final summary
        for component, ingredients in ingredient_summary.items():
            if ingredients:
                logger.info(f"{component} section has {len(ingredients)} items: {list(ingredients.keys())}")
            else:
                logger.info(f"{component} section is empty")
        return ingredient_summary

    def create_to_make_sheet_excel(self, ingredient_summary):
        """Create Excel file with to-make sheet formatted for printing"""
        try:
            # current_date = datetime.now().strftime("%Y%m%d_%H%M%S")
            output = BytesIO()
            
            # Create workbook and worksheet
            wb = Workbook()
            ws = wb.active
            ws.title = "To-Make Sheet"
            
            # Set up print formatting
            ws.page_setup.paperSize = ws.PAPERSIZE_LETTER
            ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
            ws.page_margins.left = 0.5
            ws.page_margins.right = 0.5
            ws.page_margins.top = 0.75
            ws.page_margins.bottom = 0.75
            ws.page_margins.header = 0.3
            ws.page_margins.footer = 0.3
            
            # Define styles
            header_font = Font(bold=True, size=14)
            component_header_font = Font(bold=True, size=12, color="FFFFFF")
            ingredient_font = Font(size=11)
            detail_font = Font(size=9)
            
            # Define fills
            component_fills = {
                'Starch': PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid"),
                'Breakfast Meat': PatternFill(start_color="FF6B35", end_color="FF6B35", fill_type="solid"),  # Orange for breakfast meat
                'Meat': PatternFill(start_color="E74C3C", end_color="E74C3C", fill_type="solid"),  
                'Sauce': PatternFill(start_color="F39C12", end_color="F39C12", fill_type="solid"),
                'Garnish': PatternFill(start_color="27AE60", end_color="27AE60", fill_type="solid"),
                'Veggie': PatternFill(start_color="8E44AD", end_color="8E44AD", fill_type="solid"),
                'Snack': PatternFill(start_color="8E44AD", end_color="8E44AD", fill_type="solid")  # Using same purple as Veggie
            }
            
            # Define borders
            thin_border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
            
            # Add main header
            """ ws.merge_cells('A1:K1')
            header_cell = ws['A1']
            header_cell.value = f"TO-MAKE SHEET - Created on{datetime.now().strftime('%B %d, %Y')}"
            header_cell.font = header_font
            header_cell.alignment = Alignment(horizontal='center', vertical='center')
            header_cell.border = thin_border """
            
            # Initialize row counters for left and right columns
            left_row = 1
            right_row = 1
            
            # Define section order and their positions
            left_sections = ['Veggie', 'Garnish']
            right_sections = ['Starch', 'Meat']
            bottom_sections = ['Sauce', 'Snack']  # New sections for bottom row
            
            # Process left column sections
            for component_type in left_sections:
                if component_type not in ingredient_summary or not ingredient_summary[component_type]:
                    continue
                    
                ingredients = ingredient_summary[component_type]
                
                # Column headers for this component
                headers = ['Ingredient', 'Preparation', 'Uncooked (g)','uncooked (lbs)'] if component_type != 'Garnish' else ['Garnishes', '# of Meals']
                for col, header in enumerate(headers, 1):
                    cell = ws.cell(row=left_row, column=col)
                    if col == 1:  # First column (Ingredient)
                        cell.value = f"{component_type.capitalize()}"
                    else:
                        cell.value = header
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    cell.font = Font(bold=True, size=10, color="FFFFFF")
                    cell.border = thin_border
                    cell.fill = component_fills[component_type]
                left_row += 1
                
                if component_type == 'Garnish':
                    # Sort garnish combinations by meal count (descending)
                    sorted_items = sorted(ingredients.items(), 
                                        key=lambda x: (x[0].split()[-1], -x[1]['meal_count']), 
                                        reverse=True)
                    
                    for item_name, data in sorted_items:
                        # Garnish row
                        ws.cell(row=left_row, column=1, value=data['garnish_combo']).font = ingredient_font
                        ws.cell(row=left_row, column=2, value=data['meal_count']).font = ingredient_font
                        
                        # Apply borders to all cells in the row
                        for col in range(1, len(headers) + 1):
                            ws.cell(row=left_row, column=col).border = thin_border
                            ws.cell(row=left_row, column=col).alignment = Alignment(wrap_text=True, vertical='top')
                        
                        left_row += 1
                else:
                    # Sort ingredients: first by name (for grouping), then by total grams (descending)
                    sorted_items = sorted(ingredients.items(), 
                                        key=lambda x: (x[0].split()[-1], -x[1]['total_grams']))
                    
                    for item_name, data in sorted_items:
                        # Ingredient row
                        ws.cell(row=left_row, column=1, value=item_name).font = ingredient_font
                        ws.cell(row=left_row, column=3, value=round(data['total_grams'], 1)).font = ingredient_font
                        ws.cell(row=left_row, column=4, value=round(data['total_grams']/453.592, 2)).font = ingredient_font
                        
                        # Apply borders to all cells in the row
                        for col in range(1, len(headers) + 1):
                            ws.cell(row=left_row, column=col).border = thin_border
                            ws.cell(row=left_row, column=col).alignment = Alignment(wrap_text=True, vertical='top')
                        
                        left_row += 1
                
                # Add spacing between components
                left_row += 1
            
            right_section_starting_column = 6
            # Process right column sections
            for component_type in right_sections:
                if component_type == 'Meat':
                    # Special handling for Meat section with sub-sections
                    breakfast_meats = ingredient_summary.get('Breakfast Meat', {})
                    lunch_dinner_meats = ingredient_summary.get('Meat', {})
                    
                    if not breakfast_meats and not lunch_dinner_meats:
                        continue
                    
                    # Main Meat header
                    headers = ['Ingredient', 'Preparation', 'Uncooked (g)','uncooked (lbs)']
                    
                    for col, header in enumerate(headers, 1):
                        cell = ws.cell(row=right_row, column=right_section_starting_column + col - 1)
                        if col == 1:  # First column (Ingredient)
                            cell.value = "Meat"
                        else:
                            cell.value = header
                        cell.alignment = Alignment(horizontal='center', vertical='center')
                        cell.font = Font(bold=True, size=10, color="FFFFFF")
                        cell.border = thin_border
                        cell.fill = component_fills['Meat']
                    right_row += 1
                    
                    # Breakfast Meat sub-section
                    if breakfast_meats:
                        # Breakfast sub-header
                        ws.cell(row=right_row, column=right_section_starting_column, value="Breakfast:").font = Font(bold=True, size=10)
                        for col in range(right_section_starting_column, right_section_starting_column + len(headers)):
                            ws.cell(row=right_row, column=col).border = thin_border
                            ws.cell(row=right_row, column=col).fill = PatternFill(start_color="FFE6CC", end_color="FFE6CC", fill_type="solid")
                        right_row += 1
                        
                        # Sort breakfast meats: first by name (for grouping), then by total grams (descending)
                        sorted_breakfast = sorted(breakfast_meats.items(), 
                                                key=lambda x: (x[0].split()[-1], -x[1]['total_grams']))
                        
                        for item_name, data in sorted_breakfast:
                            ws.cell(row=right_row, column=right_section_starting_column, value=item_name).font = ingredient_font
                            ws.cell(row=right_row, column=right_section_starting_column + 2, value=round(data['total_grams'], 1)).font = ingredient_font
                            ws.cell(row=right_row, column=right_section_starting_column + 3, value=round(data['total_grams']/453.592, 2)).font = ingredient_font
                            
                            for col in range(right_section_starting_column, right_section_starting_column + len(headers)):
                                ws.cell(row=right_row, column=col).border = thin_border
                                ws.cell(row=right_row, column=col).alignment = Alignment(wrap_text=True, vertical='top')
                            
                            right_row += 1
                    
                    # Lunch/Dinner Meat sub-section
                    if lunch_dinner_meats:
                        # Lunch/Dinner sub-header
                        ws.cell(row=right_row, column=right_section_starting_column, value="Lunch/Dinner:").font = Font(bold=True, size=10)
                        for col in range(right_section_starting_column, right_section_starting_column + len(headers)):
                            ws.cell(row=right_row, column=col).border = thin_border
                            ws.cell(row=right_row, column=col).fill = PatternFill(start_color="FFE6CC", end_color="FFE6CC", fill_type="solid")
                        right_row += 1
                        
                        # Sort lunch/dinner meats: first by first word, then by last word
                        sorted_lunch_dinner = sorted(lunch_dinner_meats.items(), 
                                                   key=lambda x: (x[0].split()[0], x[0].split()[-1]))
                        
                        for item_name, data in sorted_lunch_dinner:
                            ws.cell(row=right_row, column=right_section_starting_column, value=item_name).font = ingredient_font
                            ws.cell(row=right_row, column=right_section_starting_column + 2, value=round(data['total_grams'], 1)).font = ingredient_font
                            ws.cell(row=right_row, column=right_section_starting_column + 3, value=round(data['total_grams']/453.592, 2)).font = ingredient_font
                            
                            for col in range(right_section_starting_column, right_section_starting_column + len(headers)):
                                ws.cell(row=right_row, column=col).border = thin_border
                                ws.cell(row=right_row, column=col).alignment = Alignment(wrap_text=True, vertical='top')
                            
                            right_row += 1
                
                else:
                    # Handle other sections normally (like Starch)
                    if component_type not in ingredient_summary or not ingredient_summary[component_type]:
                        continue
                        
                    ingredients = ingredient_summary[component_type]
                    
                    # Column headers for this component
                    headers = ['Ingredient', 'Preparation', 'Uncooked (g)','uncooked (lbs)']
                    
                    for col, header in enumerate(headers, 1):
                        cell = ws.cell(row=right_row, column=col + right_section_starting_column - 1)  # Start from column G (7)
                        if col == 1:  # First column (Ingredient)
                            cell.value = f"{component_type.capitalize()}"
                        else:
                            cell.value = header
                        cell.alignment = Alignment(horizontal='center', vertical='center')
                        cell.font = Font(bold=True, size=10, color="FFFFFF")
                        cell.border = thin_border
                        cell.fill = component_fills[component_type]
                    right_row += 1
                    
                    # Sort ingredients: first by name (for grouping), then by total grams (descending)
                    sorted_items = sorted(ingredients.items(), 
                                        key=lambda x: (x[0].split()[-1], -x[1]['total_grams']))
                    
                    for item_name, data in sorted_items:
                        # Ingredient row
                        ws.cell(row=right_row, column=right_section_starting_column, value=item_name).font = ingredient_font
                        ws.cell(row=right_row, column=right_section_starting_column + 2, value=round(data['total_grams'], 1)).font = ingredient_font
                        ws.cell(row=right_row, column=right_section_starting_column + 3, value=round(data['total_grams']/453.592, 2)).font = ingredient_font
                        # Apply borders to all cells in the row
                        for col in range(right_section_starting_column, right_section_starting_column + len(headers)):
                            ws.cell(row=right_row, column=col).border = thin_border
                            ws.cell(row=right_row, column=col).alignment = Alignment(wrap_text=True, vertical='top')
                        
                        right_row += 1
                
                # Add spacing between components
                right_row += 1

            # Process bottom sections (Sauce and Snack)
            bottom_row = right_row + 2 
            
            # Process Sauce section
            if 'Sauce' in ingredient_summary and ingredient_summary['Sauce']:
                sauce_headers = ['Sauce to Make', '# of Servings']
                start_col = right_section_starting_column  # Start from column G
                
                for col, header in enumerate(sauce_headers, 1):
                    cell = ws.cell(row=bottom_row, column=col + start_col - 1)
                    if col == 1:  # First column
                        cell.value = "Sauce"
                    else:
                        cell.value = header
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    cell.font = Font(bold=True, size=10, color="FFFFFF")
                    cell.border = thin_border
                    cell.fill = component_fills['Sauce']
                
                # Sort sauces: first by name (for grouping), then by total count (descending)
                sorted_items = sorted(ingredient_summary['Sauce'].items(), 
                                    key=lambda x: (x[0].split()[-1], -x[1]['total_count']))
                
                for item_name, data in sorted_items:
                    # Row
                    ws.cell(row=bottom_row + 1, column=start_col, value=item_name).font = ingredient_font
                    ws.cell(row=bottom_row + 1, column=start_col + 1, value=data['total_count']).font = ingredient_font
                    
                    # Apply borders to all cells in the row
                    for col in range(start_col, start_col + len(sauce_headers)):
                        ws.cell(row=bottom_row + 1, column=col).border = thin_border
                        ws.cell(row=bottom_row + 1, column=col).alignment = Alignment(wrap_text=True, vertical='top')
                    
                    bottom_row += 1
            
            # Process Snack section (starting at the same row as Sauce)
            if 'Snack' in ingredient_summary and ingredient_summary['Snack']:
                snack_headers = ['Snack', 'Total Grams (g)','Total Grams (lbs)']
                start_col = right_section_starting_column + 3 
                
                for col, header in enumerate(snack_headers, 1):
                    cell = ws.cell(row=bottom_row - len(ingredient_summary['Sauce'].items()) if 'Sauce' in ingredient_summary else bottom_row, 
                                 column=col + start_col - 1)
                    if col == 1:  # First column
                        cell.value = "Snack"
                    else:
                        cell.value = header
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    cell.font = Font(bold=True, size=10, color="FFFFFF")
                    cell.border = thin_border
                    cell.fill = component_fills['Snack']
                
                # Sort snacks: first by name (for grouping), then by total grams (descending)
                sorted_items = sorted(ingredient_summary['Snack'].items(), 
                                    key=lambda x: (x[0].split()[-1], -x[1]['total_grams']))
                
                for item_name, data in sorted_items:
                    # Row
                    ws.cell(row=bottom_row - len(ingredient_summary['Sauce'].items()) + 1 if 'Sauce' in ingredient_summary else bottom_row + 1, 
                           column=start_col, value=item_name).font = ingredient_font
                    ws.cell(row=bottom_row - len(ingredient_summary['Sauce'].items()) + 1 if 'Sauce' in ingredient_summary else bottom_row + 1, 
                           column=start_col + 1, value=round(data['total_grams'], 1)).font = ingredient_font
                    ws.cell(row=bottom_row - len(ingredient_summary['Sauce'].items()) + 1 if 'Sauce' in ingredient_summary else bottom_row + 1, 
                           column=start_col + 2, value=round(data['total_grams']/453.592, 2)).font = ingredient_font
                    # Apply borders to all cells in the row
                    for col in range(start_col, start_col + len(snack_headers)):
                        ws.cell(row=bottom_row - len(ingredient_summary['Sauce'].items()) + 1 if 'Sauce' in ingredient_summary else bottom_row + 1, 
                               column=col).border = thin_border
                        ws.cell(row=bottom_row - len(ingredient_summary['Sauce'].items()) + 1 if 'Sauce' in ingredient_summary else bottom_row + 1, 
                               column=col).alignment = Alignment(wrap_text=True, vertical='top')
                    
                    bottom_row += 1
            
            # Set row heights
            max_row = max(left_row, right_row, bottom_row)
            for row in range(1, max_row):
                ws.row_dimensions[row].height = 20
            
            # Save to BytesIO
            wb.save(output)
            output.seek(0)
            
            logger.info(f"Successfully created to-make sheet Excel file")
            return output
            
        except Exception as e:
            logger.error(f"Error creating to-make sheet Excel: {str(e)}")
            raise AirTableError(f"Error creating to-make sheet Excel: {str(e)}")

    def generate_to_make_sheet(self):
        """Main method to generate the to-make sheet"""
        try:
            logger.info("Starting to-make sheet generation")
            
            # Get client servings data
            client_servings = self.get_clientservings_data()
            if not client_servings:
                raise AirTableError("No client servings data found")
                
            logger.info(f"Found {len(client_servings)} client servings")
            
            # Group ingredients by component
            ingredient_summary = self.group_ingredients_by_component(client_servings)
            
            # Log summary
            for component, ingredients in ingredient_summary.items():
                logger.info(f"{component}: {len(ingredients)} unique ingredients")
            
            # move Granola with Nuts from starch to snack if any
            if 'Granola with Nuts' in ingredient_summary['Starch']:
                ingredient_summary['Snack']['Granola with Nuts'] = ingredient_summary['Starch']['Granola with Nuts']
                del ingredient_summary['Starch']['Granola with Nuts']

            # move any eggs from lunch/dinner meat to breakfast meat
            eggs_to_move = []
            for meat in ingredient_summary['Meat']:
                if 'egg' in meat.lower() or 'eggs' in meat.lower():
                    eggs_to_move.append(meat)
            
            for meat in eggs_to_move:
                ingredient_summary['Breakfast Meat'][meat] = ingredient_summary['Meat'][meat]
                del ingredient_summary['Meat'][meat]
            
            # move fruits from breakfast meat to snack section
            if ingredient_summary['Veggie']:
                
                veggies_ingredients = list(ingredient_summary['Veggie'].keys())
                
                # Identify fruits using OpenAI
                fruits = self.identify_fruits_with_openai(veggies_ingredients)
                
                # Move fruits from breakfast meat to snack
                fruits_to_move = []
                for ingredient in ingredient_summary['Veggie']:
                    if ingredient in fruits:
                        fruits_to_move.append(ingredient)
                
                for fruit in fruits_to_move:
                    ingredient_summary['Snack'][fruit] = ingredient_summary['Veggie'][fruit]
                    del ingredient_summary['Veggie'][fruit]
            
            # Create Excel file
            excel_file = self.create_to_make_sheet_excel(ingredient_summary)
            print(f"cached ingredients: {self.conversion_cache}")
            return excel_file
            
        except Exception as e:
            logger.error(f"Error generating to-make sheet: {str(e)}")
            raise AirTableError(f"Error generating to-make sheet: {str(e)}")

def new_to_make_generator():
    """Factory function to create a new ToMakeSheetGenerator instance"""
    try:
        return ToMakeSheetGenerator()
    except AirTableError as e:
        logger.critical(f"Failed to create ToMakeSheetGenerator instance: {str(e)}")
        raise

if __name__ == "__main__":
    try:
        generator = new_to_make_generator()
        excel_file = generator.generate_to_make_sheet()
        
        # Write the BytesIO content to a file
        filename = f'to_make_sheet_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
        with open(filename, 'wb') as f:
            f.write(excel_file.getvalue())
        
        logger.info(f"Successfully generated to-make sheet: {filename}")
        
    except AirTableError as e:
        logger.critical(f"Application error: {str(e)}")
    except Exception as e:
        logger.critical(f"Unexpected error: {str(e)}")