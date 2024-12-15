import streamlit as st
import pandas as pd
from pptx import Presentation
import copy
from datetime import datetime
from math import ceil
from pptx.util import Pt
from pyairtable.api.table import Table  # Use direct table import
from pyairtable.formulas import match
import os
from dotenv import load_dotenv
from functools import cache
import streamlit as st

class AirTable():
    def __init__(self, ex_api_key=None):
        # Load environment variables from the .env file
        load_dotenv()
        
        # Get the API key from environment variables or the passed argument
        self.api_key = ex_api_key or st.secrets["AIRTABLE_API_KEY"]
        self.base_id = "appEe646yuQexwHJo"
        
        # Initialize tables
        self.open_orders_table = Table(self.api_key, self.base_id, 'tblxT3Pg9Qh0BVZhM')
        self.fields = {
            'SHIPPING_ADDRESS_1': 'flddnU6Y02iIHp16G',
            'SHIPPING_CITY': 'fldBIgt7ce5fEYLDG',
            'SHIPPING_PROVINCE': 'fldGeLuYoGKiw227Y',
            'SHIPPING_COUNTRY': 'fldnP3H74kAXKaIhN',
            'SHIPPING_POSTAL_CODE': 'fldwqwf7WSbiP0hJg',
            'SHIPPING_NAME': 'fldIEhlbz7JzbpTOK',
            'SHIPPING_ADDRESS_2': 'fldRUUiFRRYQ52k0W',
            'QUANTITY': 'fldvkwFMlBOW5um2y',
            'SHIPPING_PHONE': 'fldMuPbe4DX0rmq5z',
        }
        
    
    def get_all_open_orders(self): 
        data = self.open_orders_table.all(fields=self.fields.values())
        df = pd.DataFrame([record['fields'] for record in data])
        return df

    def process_data(self, df):
        portion_per_sticker=6
        df['Shipping Phone'] = (df['Shipping Phone'].astype(str).fillna('')
                .str.replace(r'\D', '', regex=True)  # Remove non-digits
                .str.replace(r'(\d{3})(\d{3})(\d{4})', r'\1-\2-\3', regex=True))
        df['Shipping Province'] = df['Shipping Province'].str.upper()
        df.fillna('', inplace=True)
        grouped_df = df.groupby([
        'Shipping Name',
        'Shipping Address 1',
        'Shipping Address 2',
        'Shipping City',
        'Shipping Province',
        'Shipping Country',
        'Shipping Postal Code',
        'Shipping Phone'
        ])['Quantity'].sum().reset_index()

        grouped_df['#_of_stickers'] = grouped_df['Quantity'].apply(lambda x: ceil(x / portion_per_sticker) * 2)
        
        return grouped_df



# Function to copy a slide from a template
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

def generate_ppt(df, prs):
    last_slide_index = 0
    template_slide = prs.slides[0]
    for index, row in df.iterrows():
        slide = copy_slide(template_slide, prs)
        last_slide_index += 1
        sticker_needed = row['#_of_stickers']
        for shape in slide.shapes:
            if shape.has_text_frame:
                text_frame = shape.text_frame
                for i, paragraph in enumerate(text_frame.paragraphs):
                    if 'Shipping Name' in paragraph.text:
                        paragraph.text = row['Shipping Name']
                        paragraph.font.size = Pt(28)
                        paragraph.font.name = "Calibri"
                    elif "Address" in paragraph.text:
                        if row['Shipping Address 2'] != '':
                            paragraph.text = f"{row['Shipping Address 1']}, {row['Shipping Address 2']}"
                            paragraph.font.size = Pt(24)
                            paragraph.font.name = "Calibri"
                        else:
                            paragraph.text = f"{row['Shipping Address 1']}"
                            paragraph.font.size = Pt(24)
                            paragraph.font.name = "Calibri"
                    elif "City" in paragraph.text:
                        paragraph.text = f"{row['Shipping City']}, {row['Shipping Province']} {row['Shipping Postal Code']}"
                        paragraph.font.size = Pt(24)
                        paragraph.font.name = "Calibri"
                    elif 'Shipping Phone' in paragraph.text:
                        paragraph.text = str(row['Shipping Phone'])
                        paragraph.font.size = Pt(24)
                        paragraph.font.name = "Calibri"
        if sticker_needed > 1:
            count = 1
            while count < sticker_needed:
                copy_slide(prs.slides[last_slide_index], prs)
                last_slide_index += 1
                count += 1
    return prs

def new_database_access():
    return AirTable()

def generate_shipping_stickers(template_ppt):
    ac = new_database_access()
    data = ac.get_all_open_orders()
    cleaned_data = ac.process_data(data)
    # print(cleaned_data[cleaned_data['Shipping Name'] == 'Christine Chung'])
    prs = generate_ppt(cleaned_data,template_ppt)
    return prs

if __name__ == "__main__":
    generate_shipping_stickers('template/Shipping_Sticker_Template.pptx')
    
