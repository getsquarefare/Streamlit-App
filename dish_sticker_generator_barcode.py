# Standard library imports
import os
import json
import glob
from datetime import datetime, timedelta

# Third-party imports
import barcode
from barcode.writer import ImageWriter
import brother_ql
from brother_ql.conversion import convert
from brother_ql.backends.helpers import send
from dotenv import load_dotenv
import pandas as pd
from PIL import Image
import pytz
from pptx import Presentation
from pptx.util import Inches
from pyairtable.api.table import Table
import streamlit as st

# Local imports
import copy

# BASE_ID = "appkvckTQ86vhjSC0"  # 041025 
BASE_ID = "appEe646yuQexwHJo" # production
TABLE_ID = "tblVwpvUmsTS2Se51"  # ClientServings
VIEW_ID = "viw5hROs9I9vV0YEq" # Sorted (For Dish Sticker)

# Set the time zone to Eastern Standard Time (EST)
est = pytz.timezone('US/Eastern')

class AirTableError(Exception):
    """Custom exception for AirTable related errors"""
    pass

class PPTGenerationError(Exception):
    """Custom exception for PowerPoint generation errors"""
    pass

def main():
    df = read_client_serving()
    # print(df.shape)
    # print(df.columns)
    df_dish = generate_sticker_df(df)
    # print(df_dish)
    # print(df_dish.head())

    create_presentation_stickers(df_dish)
    
    # Optionally print the generated stickers
    print_stickers()

def resize_image(image_path, target_width, target_height):
    """
    Resize the image to the specified width and height.

    Args:
        image_path (str): Path to the image file.
        target_width (int): Target width for the resized image.
        target_height (int): Target height for the resized image.
    """
    try:
        with Image.open(image_path) as img:
            img = img.resize((target_width, target_height), Image.Resampling.LANCZOS)
            img.save(image_path)
        print(f"Successfully resized {image_path} to {target_width}x{target_height}")
    except Exception as e:
        print(f"Error resizing {image_path}: {str(e)}")

def sort_key(filename):
    """
    Extract the numeric part from the filename for sorting.

    Args:
        filename (str): The filename to extract the number from.

    Returns:
        int: The extracted number for sorting.
    """
    last_part = filename.split('-')[-1]
    number = int(last_part.split('.')[0])
    return number

def print_stickers():
    """
    Print the generated dish stickers using Brother QL printer.
    """
    current_file_path = os.path.abspath(__file__)
    
    # Set up printer configuration
    model = 'QL-810W'
    
    # Configure source folder and printer settings
    source_folder = "/Users/clairegoldwitz/Desktop/image_test/test.jpg"
    backend = 'pyusb'
    identifier = 'usb://0x04f9:0x209c'
    
    # Alternative configuration for network printing
    # source_folder = os.path.join(os.path.dirname(current_file_path), 'assets', 'images')
    # backend = 'network'
    # identifier = 'tcp://192.168.4.30'
    
    # Load all JPG and PNG files in order
    try:
        label_files = glob.glob(f'{source_folder}/*.jpg') + glob.glob(f'{source_folder}/*.png')
        label_files.sort(key=sort_key, reverse=True)
        print(f"Found {len(label_files)} image files")
    except Exception as e:
        print(f"Error loading image files: {str(e)}")
        return
    
    if not label_files:
        print("No image files found in the directory.")
        return
    
    for i, label_file in enumerate(label_files):
        # print(f'LABEL FILE: {label_file} | {i+1} of {len(label_files)}')
        # Create the label maker object
        try:
            qlr = brother_ql.BrotherQLRaster(model)
            qlr.exception_on_warning = True
            print("Successfully created label maker object")
        except Exception as e:
            print(f"Error creating label maker object: {str(e)}")
            return
        
        img = Image.open(label_file)
        # img = img.resize((3800, 2512))
        
        try:
            instructions = convert(
                qlr=qlr,
                images=[img],
                label='62',
                cut=True,
                rotate='90'
            )
            print(f"Successfully converted label {i+1}")
        except Exception as e:
            print(f"Error converting label {i+1}: {str(e)}")
            return
        
        try:
            send(instructions=instructions, printer_identifier=identifier, backend_identifier=backend)
            print(f"Successfully sent label {i+1} to printer")
        except Exception as e:
            print(f"Error sending label {i+1} to printer: {e}")
            return
            
        # Delete previous file:
        try:
            if i > 0 and os.path.exists(label_files[i - 1]):
                print(f"Deleting previous file: {label_files[i - 1]}")
                os.remove(label_files[i - 1])
        except Exception as e:
            print(f"Error deleting previous file: {e}. Check permissions.")
            
    try:
        if os.path.exists(label_files[-1]):
            print(f"Deleting previous file: {label_files[-1]}")
            os.remove(label_files[-1])
    except Exception as e:
        print(f"Error deleting previous file: {e}. Check permissions.")
            
    print("\nPrinting process completed.")

# Process Data 
def generate_sticker_df(df):
    # Debug: Print available columns
    print(f"Available columns: {list(df.columns)}")
    
    # Define expected column names and their possible variations
    column_mapping = {
        "#": ["#", "ID", "id", "Record ID"],
        "Customer Name": ["Customer Name", "customer_name", "CustomerName"],
        "Meal Sticker (from Linked OrderItem)": ["Meal Sticker (from Linked OrderItem)", "meal_sticker", "MealSticker"],
        "Order Type (from Linked OrderItem)": ["Order Type (from Linked OrderItem)", "order_type", "OrderType"],
        "Delivery Date": ["Delivery Date", "delivery_date", "DeliveryDate"],
        "Position Id": ["Position Id", "position_id", "PositionId"],
        "Dish ID (from Linked OrderItem)": ["Dish ID (from Linked OrderItem)"]
    }
    
    # Try to find the correct column names
    found_columns = {}
    missing_columns = []
    
    for expected_col, possible_names in column_mapping.items():
        found = False
        for possible_name in possible_names:
            if possible_name in df.columns:
                found_columns[expected_col] = possible_name
                found = True
                break
        if not found:
            missing_columns.append(expected_col)
    
    if missing_columns:
        raise ValueError(f"Missing required columns: {missing_columns}. Available columns: {list(df.columns)}")
    
    # Select columns using the found column names
    selected_columns = [found_columns[col] for col in ["#", "Customer Name", "Meal Sticker (from Linked OrderItem)", "Order Type (from Linked OrderItem)", "Delivery Date", "Position Id", "Dish ID (from Linked OrderItem)"]]
    df_dish = df[selected_columns].dropna()
    
    # Rename columns to expected names for consistency
    df_dish = df_dish.rename(columns={found_columns[col]: col for col in found_columns})
    
    # Generate Requested Columns
    # barcode ID: got a float - change to string
    df_dish["#"] = df_dish["#"].apply(lambda x: str(int(x)))

    # all linked item needs to be extracted from a list
    df_dish['client_name'] = df_dish['Customer Name'].apply(lambda x: x[0] if isinstance(x, list) else x)

    # breakfast, lunch, etc.
    df_dish['meal'] = df_dish['Order Type (from Linked OrderItem)'].apply(lambda x: x[0] if isinstance(x, list) else x)

    def add_days(date_str, days = 3):  # by default: add 3 days based on EST
        date = datetime.now(est).strptime(date_str, "%Y-%m-%d")
        end_date = date + timedelta(days)
        return end_date.strftime('%m/%d/%Y')
    
    df_dish['Best Before'] = df_dish['Delivery Date'].apply(lambda x: add_days(x[0] if isinstance(x, list) else x, 3))
    df_dish['best_before'] = "Best before: " + df_dish['Best Before']

    meal_info = df_dish['Meal Sticker (from Linked OrderItem)'].apply(lambda x: x[0] if isinstance(x, list) else x)
    df_dish[['bowl_name', 'ingredients']] = meal_info.str.split(': ', n=1, expand=True)

    # duplicate according to '# portions'
    # df_dish = df_dish.loc[df_dish.index.repeat(df_dish['# portions'])]
    # df_dish = df_dish.reset_index(drop=True)
    
    return df_dish

"""### Draw Slides"""

def insert_background(slide, img_path, prs):
    left = top = Inches(0)
    pic = slide.shapes.add_picture(img_path, left, top, 
                                width=prs.slide_width, height=prs.slide_height)
    # This moves it to the background
    slide.shapes._spTree.remove(pic._element)
    slide.shapes._spTree.insert(2, pic._element)
    return slide

def copy_slide(template, prs):
    # Create a new slide with the same layout as the source slide
    new_slide = prs.slides.add_slide(template.slide_layout)

    # Copy the content from the source slide to the new slide
    for shape in template.shapes:
        new_shape = copy.deepcopy(shape)
        new_slide.shapes._spTree.insert_element_before(new_shape._element, 'p:extLst')

    return new_slide

def create_presentation_stickers(df):
    # Load the presentation
    prs = Presentation('template/Dish_Sticker_Template_Barcode.pptx')

    # Get the first slide
    slide = prs.slides[0]
    # for i in slide.shapes:
    #     print('%d %s' % (i.shape_id, i.name))
    #     print(i.text_frame.text)

    # https://python-barcode.readthedocs.io/en/stable/getting-started.html#usage
    ITF = barcode.get_barcode_class('ITF')
    BARCODE_HEIGHT = 0.75
    margin = Inches(0.05)

    
    if 'Dish ID (from Linked OrderItem)' in df.columns and 'Position Id' in df.columns:
        df['Dish ID (from Linked OrderItem)'] = df['Dish ID (from Linked OrderItem)'].apply(
            lambda x: int(x[0]) if isinstance(x, list) and len(x) > 0 and str(x[0]).isdigit() 
            else (int(x) if isinstance(x, (int, float)) else x)
        )
        df.sort_values(by=['Dish ID (from Linked OrderItem)', 'Position Id'], ascending=[True, True], inplace=True)

    # iterate each row in df_dish
    for _, row in df.iterrows():
        # add one page copied from template
        new_slide = copy_slide(slide, prs)

        # change the content of the text boxes
        for i, shape in enumerate(new_slide.shapes): 
            # Get the text frame
            if shape.has_text_frame:
                text_frame = shape.text_frame
                for paragraph in text_frame.paragraphs:
                    for run in paragraph.runs:
                        print(row[run.text])
                        run.text = row[run.text]
                    break

        # ITF barcode generator, with transparent background
        itf = ITF(row["#"], writer = ImageWriter(mode = "RGBA")) 
        itf.save("id_barcode", dict(quiet_zone=3, background = (255, 255, 255, 0), 
                                    font_size = 5, text_distance = 2.5,
                                    module_width = 0.3))
        
        id_barcode = Image.open("id_barcode.png")
        barcode_size = ((Inches(id_barcode.size[0] /id_barcode.size[1] * BARCODE_HEIGHT)), Inches(BARCODE_HEIGHT))

        new_slide.shapes.add_picture('id_barcode.png', prs.slide_width - margin - barcode_size[0],
                                        prs.slide_height - margin - barcode_size[1],
                                        width = barcode_size[0], height = barcode_size[1])

        # add background
        new_slide = insert_background(new_slide, 'template/Dish_Sticker_Template_Barcode.png', prs)
    current_date_time = datetime.now(est).strftime("%Y%m%d_%H%M")
    updated_ppt_name = f'{current_date_time}_dish_sticker_barcode.pptx'
    prs.save(updated_ppt_name)


def read_client_serving():
    load_dotenv()
    api_key = st.secrets.get("AIRTABLE_API_KEY")
    
    if not api_key:
        raise AirTableError("Airtable API key not found. Please check your environment variables or secrets.")
    
    try:
        # Initialize table
        table = Table(api_key, BASE_ID, TABLE_ID)
        
        # Get all records from the specified view
        records = table.all(view=VIEW_ID)
        
        if not records:
            raise AirTableError("No records found in the specified view.")
        
        # Convert records to DataFrame
        df = pd.DataFrame(records)
        
        # Extract fields from the records
        if 'fields' in df.columns:
            fields_df = pd.json_normalize(df['fields'])
            # Combine the records with their fields
            df_flat = pd.concat([df.drop('fields', axis=1), fields_df], axis=1)
        else:
            df_flat = df
            
        return df_flat
        
    except Exception as e:
        raise AirTableError(f"Error fetching data from Airtable: {str(e)}")

if __name__ == "__main__":
    main()