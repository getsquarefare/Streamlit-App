import streamlit as st
import pandas as pd
from pptx import Presentation
import copy
from datetime import datetime
from math import ceil
from pptx.util import Pt
from pyairtable.api.table import Table
from pyairtable.formulas import match
import os
from dotenv import load_dotenv
from functools import cache
import traceback
import logging

# Set up logging
logging.basicConfig(level=logging.INFO, 
                   format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class AirTableError(Exception):
    """Custom exception for AirTable related errors"""
    pass

class PPTGenerationError(Exception):
    """Custom exception for PowerPoint generation errors"""
    pass

class AirTable():
    def __init__(self, ex_api_key=None):
        try:
            # Load environment variables from the .env file
            load_dotenv()
            
            # Get the API key from environment variables or the passed argument
            self.api_key = ex_api_key or st.secrets.get("AIRTABLE_API_KEY")
            
            if not self.api_key:
                raise AirTableError("Airtable API key not found. Please check your environment variables or secrets.")
                
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
                'MEAL_TYPE': 'fldCsBzoy9rxKlWmN'
            }
        except Exception as e:
            logger.error(f"Error initializing AirTable: {str(e)}")
            raise AirTableError(f"Failed to initialize AirTable connection: {str(e)}")
    
    def get_all_open_orders(self): 
        try:
            logger.info("Fetching open orders from Airtable")
            data = self.open_orders_table.all(fields=self.fields.values(), view='viwDpTtU0qaT9NcvG')
            
            if not data:
                logger.warning("No open orders found in Airtable")
                return pd.DataFrame()
                
            df = pd.DataFrame([record['fields'] for record in data])
            
            # Rename columns to match expected format
            field_mapping = {v: k for k, v in self.fields.items()}
            df = df.rename(columns=field_mapping)
            
            # Standardize column names
            df.columns = [col.replace('_', ' ').title() for col in df.columns]
            
            # Make sure 'Order Type' column exists
            if 'Meal Type' in df.columns:
                df.rename(columns={'Meal Type': 'Order Type'}, inplace=True)
            
            if 'Order Type' not in df.columns:
                df['Order Type'] = 'Standard'  # Default value
                
            return df
            
        except Exception as e:
            logger.error(f"Error fetching data from Airtable: {str(e)}")
            raise AirTableError(f"Failed to fetch data from Airtable: {str(e)}")

    def process_data(self, df):
        try:
            if df.empty:
                logger.warning("No data to process")
                return pd.DataFrame()
            df.dropna(subset=['Shipping Name', 'Shipping Address 1','Quantity'], inplace=True)
            logger.info("Processing order data")
            portion_per_sticker = 6
            
            # Check if required columns exist
            required_cols = ['Quantity', 'Order Type', 'Shipping Phone', 'Shipping Province',
                            'Shipping Name', 'Shipping Address 1',
                            'Shipping City', 'Shipping Postal Code']
            
            missing_cols = [col for col in required_cols if col not in df.columns]
            if missing_cols:
                logger.error(f"Missing columns in data: {missing_cols}")
                raise ValueError(f"Missing required columns: {', '.join(missing_cols)}. Please check the Airtable structure.")
            
            # Process quantity based on order type
            df['Quantity'] = df.apply(
                lambda row: row['Quantity'] * 0.5 if row['Order Type'] == 'Breakfast' 
                else (row['Quantity'] * 0.25 if row['Order Type'] == 'Snack' else row['Quantity']), 
                axis=1
            )
            
            # Format phone numbers 
            df['Shipping Phone'] = (df['Shipping Phone'].astype(str).fillna('')
                    .str.replace(r'\D', '', regex=True)  # Remove non-digits
                    .str.replace(r'(\d{3})(\d{3})(\d{4})', r'\1-\2-\3', regex=True))
                    
            # Standardize province format
            df['Shipping Province'] = df['Shipping Province'].str.upper()
            
            # Fill any missing values
            df.fillna('', inplace=True)
            
            # Group by shipping information
            grouped_df = df.groupby([
                'Shipping Name',
                'Shipping Address 1',
                'Shipping Address 2',
                'Shipping City',
                'Shipping Province',
                'Shipping Postal Code',
                'Shipping Phone'
            ])['Quantity'].sum().reset_index()

            # Calculate number of stickers needed
            grouped_df['#_of_stickers'] = grouped_df['Quantity'].apply(
                lambda x: ceil(x / portion_per_sticker) * 2
            )
            
            return grouped_df
            
        except Exception as e:
            logger.error(f"Error processing data: {str(e)}")
            raise ValueError(f"Failed to process order data: {str(e)}")

# Function to copy a slide from a template
def copy_slide(template, target_prs):
    try:
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
    except Exception as e:
        logger.error(f"Error copying slide: {str(e)}")
        raise PPTGenerationError(f"Failed to copy slide template: {str(e)}")

def generate_ppt(df, prs):
    try:
        if df.empty:
            logger.warning("No data available to generate PowerPoint")
            raise ValueError("No shipping data available to generate stickers. Please check the Airtable data.")
        if len(prs.slides) == 0:
            logger.error("Template presentation has no slides")
            raise PPTGenerationError("The template presentation has no slides. Please check the template file.")
            
        logger.info(f"Generating PowerPoint with {len(df)} shipping entries")
        last_slide_index = 0
        template_slide = prs.slides[0]
        total_stickers = 0
        
        for index, row in df.iterrows():
            slide = copy_slide(template_slide, prs)
            last_slide_index += 1
            sticker_needed = row['#_of_stickers']
            total_stickers += sticker_needed
            
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
                    
        logger.info(f"Successfully generated {total_stickers} stickers")
        return prs
        
    except Exception as e:
        logger.error(f"Error generating PowerPoint: {str(e)}")
        raise PPTGenerationError(f"Failed to generate shipping stickers: {str(e)}")

def new_database_access():
    try:
        return AirTable()
    except AirTableError as e:
        raise e

def generate_shipping_stickers(template_ppt):
    try:
        logger.info("Starting shipping sticker generation process")
        ac = new_database_access()
        data = ac.get_all_open_orders()
        
        if data.empty:
            logger.warning("No open orders found")
            raise ValueError("No open orders found in Airtable. Please check the 'Open Orders > Running Portioning' view.")
            
        cleaned_data = ac.process_data(data)
        logger.info(f"Processing completed for {len(cleaned_data)} shipping entries")
        
        prs = generate_ppt(cleaned_data, template_ppt)
        return prs
        
    except AirTableError as e:
        logger.error(f"AirTable Error: {str(e)}")
        raise
    except PPTGenerationError as e:
        logger.error(f"PowerPoint Generation Error: {str(e)}")
        raise
    except Exception as e:
        logger.error(f"Unexpected error: {str(e)}")
        error_details = traceback.format_exc()
        logger.debug(f"Detailed error: {error_details}")
        raise ValueError(f"An unexpected error occurred: {str(e)}")


if __name__ == "__main__":
    try:
        prs_file = Presentation('template/Shipping_Sticker_Template.pptx')
        prs = generate_shipping_stickers(prs_file)
        output_path = f'shippingsticker_{datetime.now().strftime("%Y%m%d_%H%M%S")}.pptx'
        prs.save(output_path)
        print(f"Stickers successfully generated and saved to {output_path}")
    except Exception as e:
        print(f"Error: {str(e)}")