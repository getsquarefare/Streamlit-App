from pyairtable.api.table import Table  # Use direct table import
from pyairtable.formulas import match
import os
from dotenv import load_dotenv
from functools import cache
import streamlit as st
from exceptions import AirTableError, AirtableDataError  # Import from new exceptions file
import logging

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class AirTableError(Exception):
    """Custom exception for AirTable operations"""
    pass

class AirTable():
    def __init__(self, ex_api_key=None):
        # Load environment variables from the .env file
        load_dotenv()
        self.conversion_cache = {}
        
        # Get the API key from environment variables or the passed argument
        self.api_key = ex_api_key or st.secrets["AIRTABLE_API_KEY"]
        self.base_id = "appEe646yuQexwHJo"
        
        # Initialize tables
        self.ingredients_table = Table(self.api_key, self.base_id, 'tblPhcO06ce4VcAPD')
        self.client_table = Table(self.api_key, self.base_id, 'tbl63hIXZYUYY774v')
        self.subscription_table = Table(self.api_key, self.base_id, 'tblk2MjS25RBH6r9F')
        self.allergies_diet_table = Table(self.api_key, self.base_id, 'tblaR6iBdiGPVEEkL')
        self.dishes_table = Table(self.api_key, self.base_id, 'tblvTGgCq6k5iQBnL')
        self.variants_rule_table = Table(self.api_key, self.base_id, 'tblfdLImqSdAfHYI7')
        self.open_orders_table = Table(self.api_key, self.base_id, 'tblxT3Pg9Qh0BVZhM')
        self.clientserving_table = Table(self.api_key, self.base_id, 'tblVwpvUmsTS2Se51')
        self.grocery_table = Table(self.api_key, self.base_id, 'tblVndbQyR3yHwoL5')
        self.portion_algo_constraints_table = Table(self.api_key, self.base_id, 'tbl3jZNKowrO1IPAm')
        self.shopify_product_table = Table(self.api_key, self.base_id, 'tblZqBM26nx9QW1mN')
        self.shopify_variants_table = Table(self.api_key, self.base_id, 'tblonWG8wVPVA9w82')

    def get_ingredient_details_by_rcd_id(self, id):
        ingredient = self.ingredients_table.get(id)['fields']
        return ingredient

    def get_ingredient_details_by_name(self, ingredient_id, component):
        INGREDIENT_ID_FIELD = 'flduR79GRxTbyfyKe'
        COMPONENT_FIELD = 'fldpvnK23aS5nEHLI'
        formula = dict()
        formula[INGREDIENT_ID_FIELD] = ingredient_id
        formula[COMPONENT_FIELD] = component
        formula = match(formula)
        ingredients = self.ingredients_table.all(formula=formula)
        if ingredients:
            return ingredients['fields']
        else:
            return None

    def get_ingredient_details_by_recId(self, recId):
        formula = dict()
        result = dict()
        formula['id'] = recId
        fields_to_return = ['Ingredient ID', 'Ingredient Name',
                            'Component', 'Grams', 'Energy (kcal)', 
                            'Carbohydrate, total (g)', 'Protein (g)', 
                            'Fat, Total (g)', 'Dietary Fiber (g)','Sodium (mg)','Calcium (mg)', 'Phosphorus, P (mg)','Fatty acids, total saturated (g)']
        ingredients = self.ingredients_table.get(recId)
        if ingredients:
            for field in fields_to_return:
                result[field] = ingredients['fields'].get(field,0)
            return result
        else:
            return None

    def get_allergy_by_id(self, record_id):
        try:
            record = self.allergies_diet_table.get(record_id)
            return record['fields']
        except Exception as e:
            # print(f"An error occurred: {e}")
            return None

    # new, return a list of tuples
    def get_allergy_by_client_id(self, client_id):
        fields_here = dict()
        if client_id is not None:
            fields_here['Client'] = client_id
        formula = match(fields_here)

        allergies = self.allergies_diet_table.all(formula=formula)

        if allergies:
            return [(allergy['id'], allergy['fields'])for allergy in allergies]
        else:
            return None

    # Open Orders
    # #changed
    # @cache

    def get_all_open_orders_for_portioning(self):
        SHOPIFY_ID = 'fldXVHeLiy8npzVnb'
        DELETIONS = 'fldwgVkboOme5380s'
        QUANTITY = 'fldvkwFMlBOW5um2y'
        SELECTED_PROTEIN = 'fldL1BpT5B6dPdh32'
        FINAL_INGREDIENTS = 'fldZhqltOLU0tiEWI'
        TO_MATCH_CLIENT_NUTRITION = 'fldjEgeRh2bGxajXT'
        DISH_ID = 'fldLOvWuvg6X9Odvw'
        PORITON_RESULT = 'fldadHgYOukaCrC6v',
        SKIP_PORTION = "fldIKtQS5bIEr1iNU",
        INDEX = 'flddofDLsRpVLe14s'
        
        open_orders = self.open_orders_table.all(
            view='viwrZHgdsYWnAMhtX',
            fields=[INDEX,
                    TO_MATCH_CLIENT_NUTRITION,
                    SHOPIFY_ID,
                    DISH_ID,
                    FINAL_INGREDIENTS,
                    SELECTED_PROTEIN,
                    QUANTITY,
                    DELETIONS,
                    PORITON_RESULT,
                    SKIP_PORTION],
                    formula="{Portion Result (in ClientServings)} = BLANK()"
                    )
        
        return open_orders

    # Subscription Orders (Intermediary Table)
    # @cache
    def get_subscription_orders(self):
        return self.subscription_table.all()

    # inside the dish, the ingredient is recorded as the id
    def get_dish_value(self):
        dishes = self.dishes_table.first()
        # print(dishes['fields'])

    def get_ingredient_sample(self):
        ingredient = self.ingredients_table.first()
        # print(ingredient)

    '''
    input: dish_id
    output: [{'id': 589, 'Dish Name': 'Lemongrass Bowl', 'Component (from Ingredient)': ['Garnish'], 'Ingredient ID': '11677 Fried Shallot', 'NDB': '11677', 'Ingredient Name': 'Fried Shallot', 'Grams': 100, 'Kcal': 7.92, 'Protein (g)': 0.275, 'Fat, Total (g)': 0.011000000000000001, 'Dietary Fiber (g)': 0.35200000000000004, 'Carbohydrate, total (g)': 1.848}, {'id': 589, 'Dish Name': 'Lemongrass Bowl', 'Component (from Ingredient)': ['Veggies'], 'Ingredient ID': '11205 Cucumber', 'NDB': '11205', 'Ingredient Name': 'Cucumber', 'Grams': 100, 'Kcal': 15.9, 'Protein (g)': 0.625, 'Fat, Total (g)': 0.178, 'Dietary Fiber (g)': 0.0, 'Carbohydrate, total (g)': 2.95}, {'id': 589, 'Dish Name': 'Lemongrass Bowl', 'Component (from Ingredient)': ['Garnish'], 'Ingredient ID': '16396 Peanut', 'NDB': '16396', 'Ingredient Name': 'Peanut', 'Grams': 100, 'Kcal': 23.12, 'Protein (g)': 1.036, 'Fat, Total (g)': 1.9440000000000002, 'Dietary Fiber (g)': 0.35600000000000004, 'Carbohydrate, total (g)': 0.7959999999999999}, {'id': 589, 'Dish Name': 'Lemongrass Bowl', 'Component (from Ingredient)': ['Veggies'], 'Ingredient ID': '11124 Carrot', 'NDB': '11124', 'Ingredient Name': 'Carrot', 'Grams': 100, 'Kcal': 48.0, 'Protein (g)': 0.941, 'Fat, Total (g)': 0.351, 'Dietary Fiber (g)': 3.1, 'Carbohydrate, total (g)': 10.3}, {
                                                                     'id': 589, 'Dish Name': 'Lemongrass Bowl', 'Component (from Ingredient)': ['Meat'], 'Ingredient ID': 'SF Roasted Tofu', 'NDB': 'SF', 'Ingredient Name': 'Roasted Tofu', 'Grams': 155, 'Kcal': 172.0, 'Protein (g)': 16.0, 'Fat, Total (g)': 11.0, 'Dietary Fiber (g)': 1.0, 'Carbohydrate, total (g)': 1.0}, {'id': 589, 'Dish Name': 'Lemongrass Bowl', 'Component (from Ingredient)': ['Sauce'], 'Ingredient ID': 'SF Vietnamese Sauce', 'NDB': 'SF', 'Ingredient Name': 'Vietnamese Sauce', 'Grams': 20, 'Kcal': 54.0, 'Protein (g)': 1.0, 'Fat, Total (g)': 5.0, 'Dietary Fiber (g)': 0.0, 'Carbohydrate, total (g)': 0.0}, {'id': 589, 'Dish Name': 'Lemongrass Bowl', 'Component (from Ingredient)': ['Veggies'], 'Ingredient ID': '11109 Cabbage', 'NDB': '11109', 'Ingredient Name': 'Cabbage', 'Grams': 100, 'Kcal': 166.0, 'Protein (g)': 12.9, 'Fat, Total (g)': 1.18, 'Dietary Fiber (g)': 1.2, 'Carbohydrate, total (g)': 25.9}, {'id': 589, 'Dish Name': 'Lemongrass Bowl', 'Component (from Ingredient)': ['Starch'], 'Ingredient ID': '20137 Quinoa', 'NDB': '20137', 'Ingredient Name': 'Quinoa', 'Grams': 100, 'Kcal': 120.0, 'Protein (g)': 4.4, 'Fat, Total (g)': 1.92, 'Dietary Fiber (g)': 2.8, 'Carbohydrate, total (g)': 21.3}]
    '''

    def get_protein_group_mapping(self):
        # Define field IDs for variants table
        INGREDIENT_FIELD = 'fldSh0BApXpOMKYb5'  # Replace with actual field ID for Ingredient
        PROTEIN_TYPE_FIELD = 'fldnDXrKI4vXeSm7Q'  # Replace with actual field ID for Final Protein Type
        records = self.variants_rule_table.all()
        protein_type_map = {}
    
        for record in records:
            fields = record.get('fields', {})
            
            # Check if this record has both required fields
            if 'Ingredient' in fields and 'Final Protein Type (portioning)' in fields:
                ingredients = fields['Ingredient']
                protein_type = fields['Final Protein Type (portioning)'].lower()
                
                # Add each ingredient to the map
                if isinstance(ingredients, list):
                    for ingredient_id in ingredients:
                        protein_type_map[ingredient_id] = protein_type

        return protein_type_map


    def get_dish_calc_nutritions_by_dishId(self, dish_id):
        try:
            DISH_FIELDS = {
                'DISH_ID': 'fldQbBplmx4oOHhR4',
                'INGREDIENT': 'fldaaMEjUKH2bCtEj'
            }

            formula = f"{DISH_FIELDS['DISH_ID']}={dish_id}"

            dish_all_ingrdts = self.dishes_table.all(formula=formula)
            if not dish_all_ingrdts:
                raise AirtableDataError(f"No dish found with ID {dish_id}")

            dishes_information = []
            for dish_ingrdt in dish_all_ingrdts:
                crt_ingrdt = dish_ingrdt['fields']

                # Check for required fields
                if 'Ingredient' not in crt_ingrdt or not crt_ingrdt['Ingredient']:
                    raise AirtableDataError(f"Missing Ingredient field in dish {dish_id}")
                if 'Grams' not in crt_ingrdt:
                    raise AirtableDataError(f"Missing Grams field in dish {dish_id}")
                if 'Dish ID' not in crt_ingrdt:
                    raise AirtableDataError(f"Missing Dish ID field in dish {dish_id}")

                try:
                    crt_ingrdt_nutrition = self.get_ingredient_details_by_rcd_id(
                        crt_ingrdt['Ingredient'][0])
                    if not crt_ingrdt_nutrition:
                        raise AirtableDataError(f"Could not find nutrition details for ingredient {crt_ingrdt['Ingredient'][0]} in dish {dish_id}")
                except Exception as e:
                    raise AirtableDataError(f"Error fetching ingredient details: {str(e)}")

                # Check for required nutrition fields
                required_fields = ['Grams', 'Energy (kcal)', 'Protein (g)', 'Fat, Total (g)', 
                                 'Dietary Fiber (g)', 'Carbohydrate, total (g)']
                missing_fields = [field for field in required_fields if field not in crt_ingrdt_nutrition]
                if missing_fields:
                    raise AirtableDataError(f"Missing nutrition fields {missing_fields} for ingredient {crt_ingrdt_nutrition.get('Ingredient Name')} in dish {dish_id}")

                crt_ingrdt_nutrition["(g)"] = crt_ingrdt_nutrition.pop("Grams")
                merged_dish_ingrdt = {'id': crt_ingrdt['Dish ID'], **
                                    dish_ingrdt['fields'], **crt_ingrdt_nutrition}

                try:
                    calculate_rate = crt_ingrdt['Grams']/crt_ingrdt_nutrition['(g)']
                    if calculate_rate <= 0:
                        raise AirtableDataError(f"Invalid calculation rate ({calculate_rate}) for ingredient {crt_ingrdt['Ingredient'][0]} in dish {dish_id}")
                except ZeroDivisionError:
                    raise AirtableDataError(f"Zero division error when calculating rate for ingredient {crt_ingrdt['Ingredient'][0]} in dish {dish_id}")

                merged_dish_ingrdt['id'] = crt_ingrdt['Ingredient'][0]
                merged_dish_ingrdt['energy'] = calculate_rate * (
                    merged_dish_ingrdt['Energy (kcal)'] if merged_dish_ingrdt['Energy (kcal)'] > 0 else
                    merged_dish_ingrdt['Energy (Atwater General Factors) (kcal)'] if merged_dish_ingrdt['Energy (Atwater General Factors) (kcal)'] > 0 else
                    0
                )
                merged_dish_ingrdt['Kcal'] = merged_dish_ingrdt['energy']
                merged_dish_ingrdt['Protein (g)'] = calculate_rate * merged_dish_ingrdt['Protein (g)']
                merged_dish_ingrdt['Fat, Total (g)'] = calculate_rate * merged_dish_ingrdt['Fat, Total (g)']
                merged_dish_ingrdt['Dietary Fiber (g)'] = calculate_rate * merged_dish_ingrdt['Dietary Fiber (g)']
                merged_dish_ingrdt['Carbohydrate, total (g)'] = calculate_rate * merged_dish_ingrdt['Carbohydrate, total (g)']
                merged_dish_ingrdt['Sodium (mg)'] = calculate_rate * merged_dish_ingrdt.get('Sodium (mg)', 0)
                merged_dish_ingrdt['Calcium (mg)'] = calculate_rate * merged_dish_ingrdt.get('Calcium (mg)', 0)
                merged_dish_ingrdt['Phosphorus, P (mg)'] = calculate_rate * merged_dish_ingrdt.get('Phosphorus, P (mg)', 0)
                merged_dish_ingrdt['Fatty acids, total saturated (g)'] = calculate_rate * merged_dish_ingrdt.get('Fatty acids, total saturated (g)', 0)

                new_dish_ingrdt = {key: merged_dish_ingrdt[key] for key in [
                    'id', 'Airtable Dish Name', 'Component (from Ingredient)', 'Ingredient ID', 'NDB', 
                    'Ingredient Name', 'Grams', 'Kcal', 'Protein (g)', 'Fat, Total (g)', 
                    'Dietary Fiber (g)', 'Carbohydrate, total (g)', 'Sodium (mg)', 'Calcium (mg)', 'Phosphorus, P (mg)', 'Fatty acids, total saturated (g)']}
                dishes_information.append(new_dish_ingrdt)

            if not dishes_information:
                raise AirtableDataError(f"No valid ingredients found for dish {dish_id}")

            return dishes_information

        except AirtableDataError:
            raise
        except Exception as e:
            raise AirtableDataError(f"Unexpected error processing dish {dish_id}: {str(e)}")

    # changed
    def get_client_email(self, id):
        try:
            # Direct access assuming 'Email' is the key
            email = self.client_table.get(id)['fields']['TypeForm_Email']
            return email
        except KeyError:
            return None

    def get_client_details(self, recId):
        formula = dict()
        result = dict()
        formula['id'] = recId
        fields_to_return = ['identifier', 'First_Name', 'Last_Name', 'goal_calories', 'goal_carbs(g)',
                            'goal_fiber(g)', 'goal_fat(g)', 'goal_protein(g)', 'Portion Algo Constraints', 'Meal','# of snacks per day','Customization Tags']
        ingredients = self.client_table.get(recId)
        if ingredients:
            for field in fields_to_return:
                result[field] = ingredients['fields'].get(field, None)
            return result
        else:
            return None

    # method name changed
    # def get_sku(self,id):
    def get_shopify_id(self, id):
        try:
            # Direct access assuming 'Email' is the key
            shopify_id = self.shopify_product_table.get(
                id)['fields']['∞ Shopify Id']
            return shopify_id
        except KeyError:
            return None

    # changed, get the "name" column
    def get_identifier(self, id):
        try:
            # Direct access assuming 'Email' is the key
            identifier = self.client_table.get(id)['fields']['Name']
            return identifier
        except KeyError:
            return None

    # Return constraints information
    def get_constraints(self, name=None):
        fields_here = dict()
        fields_here['Name'] = name
        formula = match(fields_here)
        constraints_data = self.portion_algo_constraints_table.all(
            formula=formula)
        return constraints_data

    # Return constraints information
    def get_constraints_details_by_rcdId(self, id):
        fields_here = self.portion_algo_constraints_table.get(id)
        return fields_here

    def get_allergies_details_by_rcdId(self, id):
        allergies_diet_data = self.allergies_diet_table.get(
            id)['fields']['Ingredient to Avoid']
        return allergies_diet_data

    def delete_all_clientservings(self):
        all_records = self.clientserving_table.all(fields=['#'])
        id_list = [x["id"] for x in all_records]
        self.clientserving_table.batch_delete(id_list)

    def get_rcdid_by_shopify_orderlineitem(self, shopify_orderlineitem):
        SHOPIFY_INTERNAL_ID = 'flddofDLsRpVLe14s'
        fields_here = dict()
        fields_here[SHOPIFY_INTERNAL_ID] = shopify_orderlineitem
        formula = match(fields_here)
        open_orders = self.open_orders_table.all(formula=formula)
        return open_orders[0]['id']

    def output_clientservings(self, portion_recommendations):
        
        # Flatten and prepare the data for Airtable
        prepared_row = {}
        prepared_row['Linked OrderItem'] = [
            portion_recommendations['Linked OrderItem']]
        prepared_row['Meat'] = portion_recommendations['Meat']
        prepared_row['Sauce'] = portion_recommendations['Sauce']
        prepared_row['Starch'] = portion_recommendations['Starch']
        prepared_row['Veggies (g)'] = portion_recommendations['Veggies (g)']
        prepared_row['Garnish (g)'] = portion_recommendations['Garnish (g)']
        prepared_row['Meat (g)'] = portion_recommendations['Meat (g)']
        prepared_row['Sauce (g)'] = portion_recommendations['Sauce (g)']
        prepared_row['Starch (g)'] = portion_recommendations['Starch (g)']
        prepared_row['Veggies'] = portion_recommendations['Veggies']
        prepared_row['Garnish'] = portion_recommendations['Garnish']
        # prepared_row["Updated Nutrition Info"] = portion_recommendations["Updated Nutrition Info"]
        prepared_row['Portion Results Need Review'] = portion_recommendations['Review Needed']
        #prepared_row['Explanation'] = portion_recommendations['Explanation']
        prepared_row['Modified Recipe Details'] = portion_recommendations.get('Modified Recipe Details', "")

        # Get nutritional values and percentages directly from updated_nutrition_info
        nutrition_info = eval(portion_recommendations["Updated Nutrition Info"])
        prepared_row['Calories (kcal)'] = nutrition_info['Calories']
        prepared_row['Protein (g)'] = nutrition_info['Protein']
        prepared_row['Fat (g)'] = nutrition_info['Fat']
        prepared_row['Fiber (g)'] = nutrition_info['Fiber']
        prepared_row['Carbs (g)'] = nutrition_info['Carbohydrates']
        prepared_row['Sodium (mg)'] = nutrition_info.get('Sodium (mg)', 0)
        prepared_row['Calcium (mg)'] = nutrition_info.get('Calcium (mg)', 0)
        prepared_row['Phosphorus, P (mg)'] = nutrition_info.get('Phosphorus, P (mg)', 0)
        prepared_row['Fatty acids, total saturated (g)'] = nutrition_info.get('Fatty acids, total saturated (g)', 0)
        prepared_row['Calories %'] = nutrition_info['Calories %']/100
        prepared_row['Protein %'] = nutrition_info['Protein %']/100
        prepared_row['Fat %'] = nutrition_info['Fat %']/100
        prepared_row['Fiber %'] = nutrition_info['Fiber %']/100
        prepared_row['Carbs %'] = nutrition_info['Carbs %']/100

        # Create the record in Airtable
        self.clientserving_table.create(prepared_row)

        return

    def get_subscription_details_by_client_identifier(self, client_identifier):
        CLIENT_IDENTIFIER_FIELD = 'fld0igiHWQd8Sgh7S'
        fields_here = dict()
        fields_here[CLIENT_IDENTIFIER_FIELD] = client_identifier
        formula = match(fields_here)
        subscription_details = self.subscription_table.all(formula=formula)
        return subscription_details

    def get_clientservings_data(self,view):
        try:
            if view:
                records = self.clientserving_table.all(view=view)
            else:
                records = self.clientserving_table.all()
            return records
        except Exception as e:
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
                result = 1.0
                raise AirtableDataError(f"Ingredient '{ingredient_name}' not found in ingredients table")
            
            # Cache the result
            self.conversion_cache[ingredient_name] = result
            return result
                
        except Exception as e:
             # Cache the default value
            self.conversion_cache[ingredient_name] = 1.0
            raise AirTableError(f"Error getting conversion factor for '{ingredient_name}': {str(e)}")

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
        INGREDIENTS_TO_RECOMMEND = 'fldyxhwB2XWzXi43r'
        FINAL_INGREDIENTS_WITH_USER_EDITS = 'fldngsHO3R0KIVEXa'
        TAGS = 'fldGu2pk4yCqq45S9'
        
        fields_to_return = [CUSTOMER_NAME, MEAT, DELIVERY_DATE, DISH, QUANTITY, ALL_DELETIONS, SAUCE,
                           STARCH, VEGGIES_G, GARNISH_G, LINKED_ORDERITEM, MEAT_G, SAUCE_G, STARCH_G, 
                           VEGGIES, GARNISH, NUTRITION_NOTES_FROM_LINKED_ORDERITEM, 
                           MEAL_PORTION_FROM_LINKED_ORDERITEM, MEAL_STICKER, DISH_ID, POSITION_ID, NEW_INGREDIENTS, INGREDIENTS_TO_RECOMMEND, FINAL_INGREDIENTS_WITH_USER_EDITS,TAGS]
        
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

    def get_all_open_orders(self,view=None):
        # Initialize tables
        if view:
            return self.open_orders_table.all(view=view)
        else:
            return self.open_orders_table.all()


def new_database_access():
    return AirTable()


# AirTableAccessObject = default_store_access()
if __name__ == "__main__":
    ac = new_database_access()
    # ac.delete_all_recommendations()
    # result = ac.get_ingredient_details_by_recId('recBSki8u4LmiEVym')
    # result = ac.open_orders_table.all(fields='Final Ingredients')
    # result = ac.get_protein_group_mapping()
    result = ac.get_all_open_orders()
    print(result)
