from pyairtable import Table
from pyairtable.formulas import match
import os
from functools import cache
from openpyxl import Workbook
from openpyxl.styles import PatternFill
import pandas as pd
from datetime import datetime
from dotenv import load_dotenv
import streamlit as st
from io import BytesIO
import logging
import re

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class AirTableError(Exception):
    """Custom exception for AirTable operations"""
    pass

class AirTable():
    def __init__(self, external_api_key=None):
        """Initialize AirTable connection with API key and table definitions"""
        try:
            # Set API key (use external_api_key if provided, otherwise use hardcoded key)
            load_dotenv()
            # Get the API key from environment variables or the passed argument
            self.api_key = external_api_key or st.secrets["AIRTABLE_API_KEY"]
            self.base_id = 'appEe646yuQexwHJo'  # Base ID for Airtable
            
            # Initialize tables
            self.clientserving_table = Table(self.api_key, self.base_id, 'tblVwpvUmsTS2Se51')
            self.dishes_table = Table(self.api_key, self.base_id, 'tblvTGgCq6k5iQBnL')
            self.ingredients_table = Table(self.api_key, self.base_id, 'tblPhcO06ce4VcAPD')
        except Exception as e:
            logger.error(f"Failed to initialize AirTable connection: {str(e)}")
            raise AirTableError(f"Failed to initialize AirTable connection: {str(e)}")

    def get_clientservings_one_dish(self, dish_id):
        """Get all client servings for a specific dish"""
        # Field IDs - keeping original names as requested
        CUSTOMER_NAME = 'fldDs6QXE6uYKIlRk'
        MEAT = 'fldksE3QaxIIHzAIi'
        DELIVERY_DATE = 'fld6YLCoS7XCFK04G'
        DISH = 'fldmqHv4aXJxuJ8E2'
        QUANTITY = 'fldfwdu2UKbcTve4a'
        ALL_DELETIONS = 'fldzkTaNYfIBKxF11'
        SAUCE = 'fldOb13TV0bymcyF6'
        STARCH = 'fldZKmvAeBTY9tRkR'
        VEGGIES_G = 'fldnPpqigL4HmN4jV'
        GARNISH_G = 'fldpzKEyEiFA5vKIw'
        LINKED_ORDERITEM = 'fld6tE8nDjVxSyEtb'
        MEAT_G = 'fld4V0qMFWsCJ6VU7'
        SAUCE_G = 'fldiudWDCfVwDGfrK'
        STARCH_G = 'fldoyv8xZjZwQ9Loh'
        VEGGIES = 'fldHaAZr6fiWZbNaf'
        GARNISH = 'fldgrNa89SJKSPtwY'
        REVIEW_NEEDED = 'fldzf5wNyvEvNQZ6N'
        NUTRITION_NOTES_FROM_LINKED_ORDERITEM = 'fldq6CuqwGFg1v98H'
        EXPLANATION = 'fldaugEJU01LZhlX7'
        LAST_MODIFIED = 'fldV6RFrEYC7YrJT6'
        UPDATED_NUTRITION_INFO = 'fldqV6jLeYa57gWly'
        MEAL_PORTION_FROM_LINKED_ORDERITEM = 'fld00H3SpKNqTbhC0'
        MEAL_STICKER = 'fldeUJtuijUAbklCQ'
        DISH_ID = 'fldhrw7U0pV4D9Cad'
        POSITION_ID = 'fldRWwXRTzUflOPgk'
        NEW_INGREDIENTS = 'fldGJYhPkoWXgCw6W'
        
        fields_to_return = [CUSTOMER_NAME, MEAT, DELIVERY_DATE, DISH, QUANTITY, ALL_DELETIONS, SAUCE,
                           STARCH, VEGGIES_G, GARNISH_G, LINKED_ORDERITEM, MEAT_G, SAUCE_G, STARCH_G, 
                           VEGGIES, GARNISH, NUTRITION_NOTES_FROM_LINKED_ORDERITEM, 
                           MEAL_PORTION_FROM_LINKED_ORDERITEM, MEAL_STICKER, DISH_ID, POSITION_ID, NEW_INGREDIENTS]
        
        try:
            filter_fields = {DISH_ID: dish_id}
            formula = match(filter_fields)
            return self.clientserving_table.all(formula=formula, fields=fields_to_return,view='viwgt50kLisz8jx7b')
        except Exception as e:
            logger.error(f"Failed to get client servings for dish ID {dish_id}: {str(e)}")
            raise AirTableError(f"Failed to get client servings for dish ID {dish_id}: {str(e)}")

    def get_dish_squarespace_name(self, dish_id):
        """Get default ingredients for a specific dish"""
        try:
            DISH_ID = 'fldQbBplmx4oOHhR4'
            filter_fields = {DISH_ID: dish_id}
            formula = match(filter_fields)
            
            dish_ingredients_records = self.dishes_table.all(formula=formula)
            name = ''
            
            for dish_ingredient in dish_ingredients_records:
                tmp_name = dish_ingredient['fields'].get('SquareSpace Product Name', None)
                if tmp_name:
                    name = tmp_name
            return name
        except Exception as e:
            logger.error(f"Failed to get SquareSpace Product Name for dish ID {dish_id}: {str(e)}")
            raise AirTableError(f"Failed to get SquareSpace Product Name for dish ID {dish_id}: {str(e)}")

    def get_dish_default_ingredients(self, dish_id):
            """Get default ingredients for a specific dish"""
            try:
                DISH_ID = 'fldQbBplmx4oOHhR4'
                filter_fields = {DISH_ID: dish_id}
                formula = match(filter_fields)
                
                dish_ingredients_records = self.dishes_table.all(formula=formula)
                dish_all_ingredients = []
                
                for dish_ingredient in dish_ingredients_records:
                    ingredients = dish_ingredient['fields'].get('Ingredient', [])
                    if ingredients:
                        dish_all_ingredients.append(ingredients[0])
                        
                return dish_all_ingredients
            except Exception as e:
                logger.error(f"Failed to get default ingredients for dish ID {dish_id}: {str(e)}")
                raise AirTableError(f"Failed to get default ingredients for dish ID {dish_id}: {str(e)}")
    def format_output_order_ingredients(self, deleted_ingredients, new_ingredients):
        """Format output for ordered ingredients with deletions and additions"""
        components_output = {
            "Meat": [], 
            "Sauce": [],
            "Starch": [], 
            "Veggies": [], 
            "Garnish": []
        }

        if len(deleted_ingredients) == 0 and len(new_ingredients) == 0:
            return components_output
            
        try:
            # Process deleted ingredients
            for ingredient_id in deleted_ingredients:
                ingredient_details = self.get_ingredient_details_by_rec_id(ingredient_id)
                if ingredient_details:
                    ingredient_name = ingredient_details['Ingredient Name']
                    ingredient_component = ingredient_details['Component']
                    if ingredient_component in components_output:
                        components_output[ingredient_component].append("NO " + ingredient_name.upper())
                        
            # Process new ingredients
            for ingredient_id in new_ingredients:
                ingredient_details = self.get_ingredient_details_by_rec_id(ingredient_id)
                if ingredient_details:
                    ingredient_name = ingredient_details['Ingredient Name']
                    ingredient_component = ingredient_details['Component']
                    if ingredient_component in components_output:
                        components_output[ingredient_component].append(ingredient_name.upper())
                        
            return components_output
        except Exception as e:
            logger.error(f"Error formatting output for ordered ingredients: {str(e)}")
            raise AirTableError(f"Error formatting output for ordered ingredients: {str(e)}")

    def format_output_default_ingredients(self, default_ingredients):
        """Format output for default ingredients"""
        components_output = {
            "Meat": [], 
            "Sauce": [],
            "Starch": [], 
            "Veggies": [], 
            "Garnish": []
        }

        try:
            for ingredient in default_ingredients:
                if not ingredient:
                    continue
                    
                ingredient_details = self.get_ingredient_details_by_rec_id(ingredient)
                if ingredient_details:
                    ingredient_name = ingredient_details['Ingredient Name']
                    ingredient_component = ingredient_details['Component']
                    if ingredient_component in components_output:
                        components_output[ingredient_component].append(ingredient_name)
                        
            return components_output
        except Exception as e:
            logger.error(f"Error formatting output for default ingredients: {str(e)}")
            raise AirTableError(f"Error formatting output for default ingredients: {str(e)}")

    def get_ingredient_details_by_rec_id(self, rec_id):
        """Get ingredient details by record ID"""
        fields_to_return = ['Ingredient Name', 'Component']
        result = {}
        
        try:
            ingredient = self.ingredients_table.get(rec_id)
            for field in fields_to_return:
                result[field] = ingredient['fields'][field]
            return result
        except Exception as e:
            logger.warning(f"Ingredient not found with ID {rec_id}: {str(e)}")
            return None

    def one_dish_output(self, dish_id):
        """Generate output for one dish including all client servings"""
        try:
            outputs = []
            default_ingredients = self.get_dish_default_ingredients(dish_id)
            dish_product_name = self.get_dish_squarespace_name(dish_id)
            ordered_clientservings = self.get_clientservings_one_dish(dish_id)
            
            default_ingredients_formatted = self.format_output_default_ingredients(default_ingredients)
            default_ingredients_formatted[' '] = dish_id
            default_ingredients_formatted['Position'] = dish_product_name
            
            outputs.append(default_ingredients_formatted)
            
            for client_serving in ordered_clientservings:
                deleted_ingredients = client_serving['fields'].get('All Deletions', [])
                new_ingredients = client_serving['fields'].get('New Ingredients', [])
                
                # Get deleted ingredient names
                deleted_ingredients_names = []
                for ingredient_id in deleted_ingredients:
                    ingredient_details = self.get_ingredient_details_by_rec_id(ingredient_id)
                    if ingredient_details:
                        deleted_ingredients_names.append(ingredient_details['Ingredient Name'])
                
                # Build output dictionary
                output = {
                    ' ': dish_id,
                    'Position': client_serving['fields'].get('Position Id', 0),
                    'Delivery Date': client_serving['fields']['Delivery Date'],
                    'Client': str(client_serving['fields']['Customer Name'][0]),
                    'Allergies': client_serving['fields'].get('Nutrition Notes (from Linked OrderItem)', [""])[0],
                    'Meal': client_serving['fields'].get('Meal Portion (from Linked OrderItem)', [''])[0].strip(),
                    'Sticker': client_serving['fields'].get('Meal Sticker (from Linked OrderItem)', [''])[0],
                    #'Dish': client_serving['fields'].get('Dish', [''])[0],
                    'All Deletions': deleted_ingredients_names,
                    'Portions': 1
                }
                
                # Get components output
                components_output = self.format_output_order_ingredients(deleted_ingredients, new_ingredients)
                
                # Add component amounts
                components_output['Meat'].append(
                    '-' if client_serving['fields']['Meat (g)'] == 0 else round(client_serving['fields']['Meat (g)'], 1))
                
                # Handle Sauce with special case for "n x sauce" format
                sauce_value = client_serving['fields']['Sauce (g)']
                sauce_text = client_serving['fields'].get('Sauce', '')
                
                if sauce_value == 0:
                    components_output['Sauce'].append('-')
                else:
                    sauce_output = str(round(sauce_value, 1))
                    # Extract any content in parentheses if it exists
                    parentheses_pattern = re.search(r'\((.*?)\)', sauce_text)
                    if parentheses_pattern:
                        sauce_output += f" ({parentheses_pattern.group(1)})"
                    components_output['Sauce'].append(sauce_output)
                
                components_output['Starch'].append(
                    '-' if client_serving['fields']['Starch (g)'] == 0 else round(client_serving['fields']['Starch (g)'], 1))
                components_output['Veggies'].append(
                    '-' if client_serving['fields']['Veggies (g)'] == 0 else round(client_serving['fields']['Veggies (g)'], 1))
                components_output['Garnish'].append(
                    '-' if client_serving['fields']['Garnish (g)'] == 0 else round(client_serving['fields']['Garnish (g)'], 1))
                
                output.update(components_output)
                outputs.append(output)
                
            return outputs
        except Exception as e:
            logger.error(f"Error generating output for dish ID {dish_id}: {str(e)}")
            raise AirTableError(f"Error generating output for dish ID {dish_id}: {str(e)}")

    def generate_formatted_clientservings_onedish(self, clientservings):
        """Format client servings data for one dish into a dataframe"""
        try:
            def flatten(item):
                if isinstance(item, list):
                    return ', '.join(map(str, item))
                return item

            flattened_data = [{k: flatten(v) for k, v in entry.items()} for entry in clientservings]
            df = pd.DataFrame(flattened_data)
            return df
        except Exception as e:
            logger.error(f"Error formatting client servings data: {str(e)}")
            raise AirTableError(f"Error formatting client servings data: {str(e)}")

    def consolidated_all_dishes_output(self):
        """Consolidate output for all dishes"""
        try:
            all_clientservings = self.clientserving_table.all(fields=['Dish ID (from Linked OrderItem)'],view='viwgt50kLisz8jx7b')
            if len(all_clientservings) == 0:
                raise AirTableError("No clientservings found in the source view. Please check the view and try again.")
            all_dishes = set()
            all_output = pd.DataFrame()
            
            # Collect all unique dish IDs
            for client_serving in all_clientservings:
                dish_ids = client_serving['fields'].get('Dish ID (from Linked OrderItem)', [])
                if dish_ids:
                    all_dishes.add(dish_ids[0])
            all_dishes = sorted(all_dishes)
            # Process each dish
            for dish_id in all_dishes:
                result = self.one_dish_output(dish_id)
                result_df = self.generate_formatted_clientservings_onedish(result)
                all_output = pd.concat([all_output, result_df], axis=0, ignore_index=True)
                
            return all_output
        except Exception as e:
            logger.error(f"Error consolidating all dishes output: {str(e)}")
            raise AirTableError(f"Error consolidating all dishes output: {str(e)}")

    def generate_clientservings_excel(self, formatted_clientservings):
        """Generate Excel file with formatted client servings data"""
        try:
            current_date = datetime.now().strftime("%Y%m%d")
            output = BytesIO()
            logger.info('Generating Excel file for formatted client servings')
            # Select and reorder columns
            formatted_clientservings = formatted_clientservings[[
                'Position', 'Client', 'Meal','Portions','Delivery Date', 
                'All Deletions', 'Sauce', 'Garnish', 'Starch', 'Veggies', 
                'Meat', 'Allergies'
            ]]
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                formatted_clientservings.to_excel(writer, index=False, sheet_name='Sheet1')

                # Access the workbook and sheet
                workbook = writer.book
                sheet = workbook['Sheet1']

                # Define fill colors
                blue_fill = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")

                # Apply row colors based on NaN index
                for i, row in enumerate(sheet.iter_rows(min_row=2, max_row=sheet.max_row, min_col=1, max_col=sheet.max_column)):
                    if pd.isna(formatted_clientservings.loc[i, 'Delivery Date']):
                        for cell in row:
                            cell.fill = blue_fill

            output.seek(0)
            return output
        except Exception as e:
            logger.error(f"Error generating Excel file: {str(e)}")
            raise AirTableError(f"Error generating Excel file: {str(e)}")


def new_database_access():
    """Factory function to create a new AirTable instance"""
    try:
        return AirTable()
    except AirTableError as e:
        logger.critical(f"Failed to create AirTable instance: {str(e)}")
        raise


if __name__ == "__main__":
    try:
        ac = new_database_access()
        all_output = ac.consolidated_all_dishes_output()
        file = ac.generate_clientservings_excel(all_output)
        
        # Write the BytesIO content to a file
        with open('clientservings_excel_output_'+datetime.now().strftime("%Y%m%d")+'.xlsx', 'wb') as f:
            f.write(file.getvalue())
        logger.info("Successfully generated Excel file")
    except AirTableError as e:
        logger.critical(f"Application error: {str(e)}")
    except Exception as e:
        logger.critical(f"Unexpected error: {str(e)}")