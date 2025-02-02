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

class AirTable():
    def __init__(self, ex_api_key=None):
        # Set API key (use ex_api_key if provided, otherwise use hardcoded key)
        load_dotenv()
        # Get the API key from environment variables or the passed argument
        self.api_key = ex_api_key or st.secrets["AIRTABLE_API_KEY"]
        self.base_id = 'appEe646yuQexwHJo'  # Base ID for Airtable
        
        # Initialize tables
        self.clientserving_table = Table(self.api_key, self.base_id, 'tblVwpvUmsTS2Se51')
        self.dishes_table = Table(self.api_key, self.base_id, 'tblvTGgCq6k5iQBnL')
        self.ingredients_table = Table(self.api_key, self.base_id, 'tblPhcO06ce4VcAPD')

    def get_clientservings_one_dish(self, dish_name):
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
        
        field_to_return = [CUSTOMER_NAME, MEAT, DELIVERY_DATE, DISH, QUANTITY, ALL_DELETIONS, SAUCE,
                           STARCH, VEGGIES_G, GARNISH_G, LINKED_ORDERITEM, MEAT_G, SAUCE_G, STARCH_G, VEGGIES, GARNISH,NUTRITION_NOTES_FROM_LINKED_ORDERITEM, MEAL_PORTION_FROM_LINKED_ORDERITEM,MEAL_STICKER]
        fields_here = dict()
        fields_here[DISH] = dish_name
        formula = match(fields_here)

        return self.clientserving_table.all(formula=formula, fields=field_to_return)

    def get_dish_default_ingridients(self, dish_name):
        DISH_NAME = 'fld2FLuMLvlsY2dTJ'
        fields_here = dict()
        fields_here[DISH_NAME] = dish_name
        formula = match(fields_here)
        dish_all_ingrdts_rec = self.dishes_table.all(
            formula=formula)
        dish_all_ingrdts = []
        for dish_ingrdt in dish_all_ingrdts_rec:
            dish_all_ingrdts.append(dish_ingrdt['fields']['Ingredient'])
        return dish_all_ingrdts

    def format_output_order_ingrdts(self, deleted_ingrdts):
        components_output = {"Meat": [], "Sauce": [],
                             "Starch": [], "Veggies": [], "Garnish": []}

        if len(deleted_ingrdts) == 0:
            return components_output
        for one_ingrdt in deleted_ingrdts:
            one_ingrdt = self.get_ingredient_details_by_recId(one_ingrdt)
            ingrdt_name = one_ingrdt['Ingredient Name']
            ingrdt_component = one_ingrdt['Component']
            if ingrdt_component in components_output:
                components_output[ingrdt_component].append("NO " + ingrdt_name.upper())
        return components_output

    def format_output_defalut_ingrdts(self, default_ingrdts):
        components_output = {"Meat": [], "Sauce": [],
                             "Starch": [], "Veggies": [], "Garnish": []}

        for one_ingrdt in default_ingrdts:
            one_ingrdt = self.get_ingredient_details_by_recId(one_ingrdt[0])
            ingrdt_name = one_ingrdt['Ingredient Name']
            ingrdt_component = one_ingrdt['Component']
            if ingrdt_component in components_output:
                components_output[ingrdt_component].append(ingrdt_name)
        return components_output

    def get_ingredient_details_by_recId(self, recId):
        formula = dict()
        result = dict()
        formula['id'] = recId
        fields_to_return = ['Ingredient Name',
                            'Component']
        try:
            ingredients = self.ingredients_table.get(recId)
            for field in fields_to_return:
                result[field] = ingredients['fields'][field]
            return result
        except:
            print('Error: Ingredient not found '+recId)

    def one_dish_output(self, dish_name):
        outputs = []
        index = 1
        default_ingrdts = self.get_dish_default_ingridients(dish_name)
        ordered_clientservings = self.get_clientservings_one_dish(dish_name)
        default_ingrdts_in_name = self.format_output_defalut_ingrdts(
            default_ingrdts)
        default_ingrdts_in_name[' '] = dish_name
        outputs.append(default_ingrdts_in_name)
        
        for one_clientserving in ordered_clientservings:
            # print('one_clientserving',one_clientserving)
            def deleted(one_clientserving): return one_clientserving['fields'].get(
                'All Deletions', [])
            deleted_ingrdts = deleted(one_clientserving)
            output = dict()
            output['Position'] = index
            output['Delivery Date'] = one_clientserving['fields']['Delivery Date']
            output['Client'] = str(one_clientserving['fields']['Customer Name'][0])
            output['# Portions'] = one_clientserving['fields']['Quantity']
            output['Allergies'] = one_clientserving['fields'].get('Nutrition Notes (from Linked OrderItem)', [""])[0]
            output['Meal'] = one_clientserving['fields'].get('Meal Portion (from Linked OrderItem)',[''])[0].replace(' Subscriptions','')
            output['Name of Dish'] = one_clientserving['fields'].get('Meal Sticker (from Linked OrderItem)',[''])[0]
            #output['Dish Name'] = dish_name
            components_output = self.format_output_order_ingrdts(
                deleted_ingrdts)

            components_output['Meat'].append(
                '-' if one_clientserving['fields']['Meat (g)'] == 0 else round(one_clientserving['fields']['Meat (g)'], 1))
            components_output['Sauce'].append(
                '-' if one_clientserving['fields']['Sauce (g)'] == 0 else round(one_clientserving['fields']['Sauce (g)'], 1))
            components_output['Starch'].append(
                '-' if one_clientserving['fields']['Starch (g)'] == 0 else round(one_clientserving['fields']['Starch (g)'], 1))
            components_output['Veggies'].append(
                '-' if one_clientserving['fields']['Veggies (g)'] == 0 else round(one_clientserving['fields']['Veggies (g)'], 1))
            components_output['Garnish'].append(
                '-' if one_clientserving['fields']['Garnish (g)'] == 0 else round(one_clientserving['fields']['Garnish (g)'], 1))
            output.update(components_output)
            outputs.append(output)
            index += one_clientserving['fields']['Quantity'][0]
        return outputs

    def generate_formatted_clientservings_onedish(self, clientservings):
        def flatten(item):
            if isinstance(item, list):
                return ', '.join(map(str, item))
            return item

        flattened_data = [{k: flatten(v) for k, v in entry.items()}
                          for entry in clientservings]
        df = pd.DataFrame(flattened_data)
        return df

    def consolidated_all_dishes_output(self):
        all_clientservings = self.clientserving_table.all(fields=['Dish'])
        all_dishes = set()
        all_output = pd.DataFrame()
        for one_clientserving in all_clientservings:
            if one_clientserving['fields']['Dish']:
                all_dishes.add(one_clientserving['fields']['Dish'][0])
        for dish in all_dishes:
            result = self.one_dish_output(dish)
            result_pd = self.generate_formatted_clientservings_onedish(result)
            all_output = pd.concat(
                [all_output, result_pd], axis=0, ignore_index=True)
        return all_output

    def generate_clientservings_excel(self, formatted_clientservings):
        current_date = datetime.now().strftime("%Y%m%d")
        output = BytesIO()

        formatted_clientservings = formatted_clientservings[[
            ' ', 'Position', 'Client', 'Meal', 'Delivery Date', 'Name of Dish', 
            '# Portions', 'Sauce', 'Garnish', 'Starch', 'Veggies', 'Meat', 'Allergies'
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
                if pd.isna(formatted_clientservings.loc[i, 'Position']):
                    for cell in row:
                        cell.fill = blue_fill

        output.seek(0)  # Move to the beginning of the file
        return output

def new_database_access():
    return AirTable()


if __name__ == "__main__":
    ac = new_database_access()

    # result = ac.one_dish_output('Lemongrass Bowl')
    all_output = ac.consolidated_all_dishes_output()
    # print(result)
    ac.export_clientservings_to_excel(all_output)

