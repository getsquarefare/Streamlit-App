import streamlit as st
import pandas as pd
from pptx import Presentation
import copy
from datetime import datetime, timedelta
from pptx.util import Inches, Pt
from pyairtable.api.table import Table
from pyairtable.formulas import match
import os
from dotenv import load_dotenv
from functools import cache

class AirTable():
    def __init__(self, ex_api_key=None):
        # Load environment variables from the .env file
        load_dotenv()
        
        # Get the API key from environment variables or the passed argument
        self.api_key = ex_api_key or st.secrets["AIRTABLE_API_KEY"]
        self.base_id = "appEe646yuQexwHJo"
        
        # Initialize tables for dish stickers
        self.dish_orders_table = Table(self.api_key, self.base_id, 'tblxT3Pg9Qh0BVZhM')  # Replace with your actual table ID
        
        # Define field mappings for dish stickers
        self.fields = {
            'CLIENT': 'fldWYJStYSpX72pG3',
            'DISH_STICKER': 'fldYZYDRjScz6ig5a',
            'DELIVERY_DAY': 'flddRJziNBdEtrpmG',
            'MEAL_TYPE': 'fldCsBzoy9rxKlWmN',
            'PROTEIN': 'fldL1BpT5B6dPdh32',
            'QUANTITY': 'fldvkwFMlBOW5um2y',
        }
    
    def get_all_dish_orders(self): 
        data = self.dish_orders_table.all(fields=self.fields.values(), view='viwrZHgdsYWnAMhtX')
        df = pd.DataFrame([record['fields'] for record in data])
        
        # Rename columns to match the expected format
        column_mapping = {v.replace('fld', ''): k for k, v in self.fields.items()}
        df.rename(columns=column_mapping, inplace=True)
        
        return df

    def process_data(self, df):
        print(df)
        # Clean column names
        df.columns = [col.strip() for col in df.columns]
        
        # Create required fields for dish stickers
        df['line_1'] = df.apply(lambda row: f"Prepared for: {row['Customer Name']} - {row['Order Type']}", axis=1)
        
        # Add best before date (delivery date + 3 days)
        def add_days(date_str, days=3):
            date = datetime.strptime(date_str, "%Y-%m-%d")
            end_date = date + timedelta(days=days)
            return end_date.strftime('%m-%d-%Y')
        
        df['Best Before'] = df['Delivery Date'].apply(lambda x: add_days(x, 3))
        df['line_2'] = "Best before: " + df['Best Before']
        
        # Set dish name
        df['line_3'] = df['Meal Sticker']

        # Convert lines to strings
        df['line_1'] = df['line_1'].astype(str)
        df['line_2'] = df['line_2'].astype(str)
        df['line_3'] = df['line_3'].astype(str)
        
        # Duplicate rows based on portion count
        return df.loc[df.index.repeat(df['Quantity'])].reset_index(drop=True)

def insert_background(slide, img_path, prs):
    left = top = Inches(0)
    width = prs.slide_width
    height = prs.slide_height
    pic = slide.shapes.add_picture(img_path, left, top, width=width, height=height)

    # This moves it to the background
    slide.shapes._spTree.remove(pic._element)
    slide.shapes._spTree.insert(2, pic._element)
    return slide

def copy_slide(template, target_prs):
    new_slide = target_prs.slides.add_slide(template.slide_layout)
    for shape in new_slide.shapes:
        sp = shape._element
        sp.getparent().remove(sp)
    for shape in template.shapes:
        if not shape.has_text_frame:
            continue
        new_shape = copy.deepcopy(shape)
        new_slide.shapes._spTree.insert_element_before(new_shape._element, 'p:extLst')
    return new_slide

def generate_ppt(df, prs, background_path):
    # Get the first slide as a template
    template_slide = prs.slides[0]
    
    # iterate each row in df_dish
    for index, row in df.iterrows():
        # add one page copied from template
        new_slide = copy_slide(template_slide, prs)

        # change text for lines 1-3
        for i, shape in enumerate(new_slide.shapes):
            if shape.has_text_frame and i < 3:  # Only process the first 3 shapes with text
                # Get the text frame
                text_frame = shape.text_frame

                for paragraph in text_frame.paragraphs:
                    for run in paragraph.runs:
                        run.text = row[f'line_{i+1}']
                    break

        # add background
        if background_path:
            new_slide = insert_background(new_slide, background_path, prs)
    
    return prs

def new_database_access():
    return AirTable()

def generate_dish_stickers(prs, background_path=None):
    # Initialize Airtable connection
    ac = new_database_access()
    
    # Get data from Airtable
    data = ac.get_all_dish_orders()
    
    # Process data
    processed_data = ac.process_data(data)
    
    # Generate PowerPoint
    prs = generate_ppt(processed_data, prs, background_path)
    
    return prs

if __name__ == "__main__":
    # Paths to template and background
    template_path = 'template/Dish_Sticker_Template.pptx'
    # background_path = 'template/Dish_Sticker_Background.jpg'
    prs_file = Presentation(template_path)
    # Generate stickers
    prs = generate_dish_stickers(prs_file)
    
    # Save the result
    output_path = f'Dish_Stickers_{datetime.now().strftime("%Y%m%d_%H%M%S")}.pptx'
    prs.save(output_path)
    print(f"Dish stickers generated and saved as {output_path}")