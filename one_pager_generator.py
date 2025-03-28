import streamlit as st
import pandas as pd
from pptx import Presentation
import copy
from datetime import datetime, timedelta
from pptx.util import Inches, Pt
from pyairtable.api.table import Table
from pyairtable.formulas import match
from dotenv import load_dotenv
from functools import cache
import math

class AirTable():
    def __init__(self, ex_api_key=None):
        # Load environment variables from the .env file
        load_dotenv()
        
        # Get the API key from environment variables or the passed argument
        self.api_key = ex_api_key or st.secrets["AIRTABLE_API_KEY"]
        self.base_id = "appEe646yuQexwHJo"  # Use the same base_id as in shipping stickers
        
        # Initialize tables for client servings and client information
        self.open_orders_table = Table(self.api_key, self.base_id, 'tblxT3Pg9Qh0BVZhM')  # Replace with your actual table ID
        self.clients_table = Table(self.api_key, self.base_id, 'tbl63hIXZYUYY774v')  # Replace with your actual table ID
        
        # Define field mappings for client servings
        self.open_orders_fields = {
            'CLIENT': 'fldjEgeRh2bGxajXT',            
            'DISH_STICKER': 'fldYZYDRjScz6ig5a',       
            'DELIVERY_DAY': 'flddRJziNBdEtrpmG', 
            'PORTIONS': 'fldCsBzoy9rxKlWmN'         
        }
        
        # Define field mappings for clients info
        self.clients_fields = {
            'IDENTIFIER': 'fldDVCUtcmEv5ZkEv',              
            'KCAL': 'fldgaURVgPSD6UFVU',
            'CARBS': 'fldRohbKpwyIyyuhe',
            'PROTEIN': 'fldLpIgHNZFxyZ1PI',
            'FAT': 'fldMSEudjCwovSD8A',
            'FIBER': 'fld1TirSqBsA1GqK5',
            'CLIENT_FNAME': 'fldGBn8BpGwFIHxgi',
            'CLIENT_LNAME': 'fldIq6giA1dDcut8T'
        }
    
    def get_open_orders(self):
        data = self.open_orders_table.all(fields=self.open_orders_fields.values(), view='viwrZHgdsYWnAMhtX')  # Replace with actual view ID
        
        # Create DataFrame and map column names
        df = pd.DataFrame([record['fields'] for record in data])
        column_mapping = {v.replace('fld', ''): k for k, v in self.open_orders_fields.items()}
        df.rename(columns=column_mapping, inplace=True)
        
        return df
    
    def get_clients_info(self):
        data = self.clients_table.all(fields=self.clients_fields.values())  # Replace with actual view ID
        
        # Create DataFrame and map column names
        df = pd.DataFrame([{**record['fields'], 'id': record['id']} for record in data])
        column_mapping = {v.replace('fld', ''): k for k, v in self.clients_fields.items()}
        df.rename(columns=column_mapping, inplace=True)
        
        return df
    
    def process_data(self):
        # Get data from both tables
        df_orders = self.get_open_orders()
        df_clients = self.get_clients_info()
        
        # Clean column names and values
        df_orders.columns = [col.strip() for col in df_orders.columns]
        df_clients.columns = [col.strip() for col in df_clients.columns]
        
        # Process client servings data
        df_orders = df_orders[['Delivery Date', 'Meal Sticker', 'Order Type', 'To_Match_Client_Nutrition']].dropna()
        
        # Convert any list values to strings in all relevant columns
        # For client identifier
        df_orders["To_Match_Client_Nutrition"] = df_orders["To_Match_Client_Nutrition"].apply(
            lambda x: x[0] if isinstance(x, list) and len(x) > 0 else 
                     (str(x) if not isinstance(x, list) else "Unknown")
        )
        
        # For meal sticker
        df_orders["Meal Sticker"] = df_orders.apply(
            lambda row: (row["Meal Sticker"][0] if isinstance(row["Meal Sticker"], list) and len(row["Meal Sticker"]) > 0 else 
                        (str(row["Meal Sticker"]) if not isinstance(row["Meal Sticker"], list) else "Unknown")) 
                        + ' - ' + row["Order Type"],
            axis=1
        )
        
        # For order type
        df_orders["Order Type"] = df_orders["Order Type"].apply(
            lambda x: x[0] if isinstance(x, list) and len(x) > 0 else 
                     (str(x) if not isinstance(x, list) else "Unknown")
        )
        
        # Group the DataFrame by client identifier column and create page indices
        page_size = 6
        grouped_df = df_orders.groupby("To_Match_Client_Nutrition")
        df_orders["Index"] = grouped_df.cumcount()
        df_orders['page_number'] = df_orders['Index'] // page_size
        
        # Create grouped dataframe with clients and pages
        df_orders_grouped = pd.DataFrame(df_orders.groupby(['To_Match_Client_Nutrition', 'page_number']).agg({
            'To_Match_Client_Nutrition': 'first', 
            'page_number': 'first', 
            'Delivery Date': 'first',
            'Order Type': 'first'
        })).reset_index(drop=True)
        
        # Calculate portion strings with appropriate line breaks
        standard_str = "Fish Taco Bowl with Grilled Chicken, Shredded Red and Green Cabbage, Roasted"
        unit_len = len(standard_str)
        
        # Determine portions based on order type
        df_orders['Portions'] = df_orders.apply(
            lambda row: 0.5 if row['Order Type'] == 'Breakfast' 
                       else (0.25 if row['Order Type'] == 'Snack' else 1),
            axis=1
        )
        
        df_orders['portion_str'] = df_orders.apply(
            lambda row: '[ ' + str(int(row['Portions'])) + ' ] ' + 
            '\n' * (math.ceil(len(row['Meal Sticker'])/unit_len) - 1), 
            axis=1
        )
        
        # Group portions and dish names by client and page
        df_orders_grouped['portion_list'] = df_orders.groupby(['To_Match_Client_Nutrition', 'page_number'])['portion_str'].apply(
            lambda x: '\n\n'.join([str(item) for item in x])
        ).reset_index(drop=True)
        
        df_orders_grouped['dish_list'] = df_orders.groupby(['To_Match_Client_Nutrition', 'page_number'])['Meal Sticker'].apply(
            lambda x: '\n\n'.join([str(item) for item in x])
        ).reset_index(drop=True)
        
        # Process client nutritional information
        df_clients_grouped = pd.DataFrame(
            df_clients.groupby('id').agg({
                'id': 'first',
                'First_Name': 'first',
                'Last_Name': 'first'
            })
        ).reset_index(drop=True)
        
        # Format nutrition information from individual components
        nutrition_lines = []
        for _, row in df_clients.iterrows():
            nutrition_line = ""
            if 'goal_calories' in row and not pd.isna(row['goal_calories']):
                nutrition_line += f"{row['goal_calories']} kcals | "
            if 'goal_carbs(g)' in row and not pd.isna(row['goal_carbs(g)']):
                nutrition_line += f"{row['goal_carbs(g)']}g carbs, "
            if 'goal_protein(g)' in row and not pd.isna(row['goal_protein(g)']):
                nutrition_line += f"{row['goal_protein(g)']}g protein, "
            if 'goal_fat(g)' in row and not pd.isna(row['goal_fat(g)']):
                nutrition_line += f"{row['goal_fat(g)']}g fat, "
            if 'goal_fiber(g)' in row and not pd.isna(row['goal_fiber(g)']):
                nutrition_line += f"{row['goal_fiber(g)']}g fiber"
            
            df_clients.loc[_, 'NUTRITION'] = nutrition_line
        
        # Now group the nutrition info
        df_clients_grouped['NUTRITION'] = df_clients.groupby('id')['NUTRITION'].apply(
            lambda x: '\n'.join(filter(None, x))
        ).reset_index(drop=True)
        
        # Rename client column in grouped dataframes for merging
        df_orders_grouped.rename(columns={'To_Match_Client_Nutrition': 'CLIENT'}, inplace=True)
        df_clients_grouped.rename(columns={'id': 'CLIENT'}, inplace=True)
        
        # Ensure client names are cleaned
        df_orders_grouped.CLIENT = df_orders_grouped.CLIENT.apply(lambda x: str(x).strip() if isinstance(x, str) else x)
        df_clients_grouped.CLIENT = df_clients_grouped.CLIENT.apply(lambda x: str(x).strip() if isinstance(x, str) else x)
        
        # Merge the two dataframes
        df_merge = df_orders_grouped.merge(df_clients_grouped, on='CLIENT', how='left')
        df_merge.NUTRITION = df_merge.NUTRITION.fillna('')
        
        # Generate final text fields
        df_merge['line_1'] = "Prepared for: " + df_merge['First_Name'] + " " + df_merge['Last_Name'] 
        
        # Add best before date (delivery date + 3 days)
        def add_days(date_str, days=3):
            # Handle different date formats
            try:
                # Try to parse as mm/dd/yyyy
                date = datetime.strptime(date_str, "%Y-%m-%d")
            except ValueError:
                try:
                    # Try ISO format
                    date = datetime.fromisoformat(date_str.replace('Z', '+00:00'))
                except ValueError:
                    # If all else fails, use today's date
                    date = datetime.now()
            
            end_date = date + timedelta(days=days)
            return end_date.strftime('%m/%d/%Y')
        
        df_merge['Best Before'] = df_merge['Delivery Date'].apply(lambda x: add_days(x, 3))
        df_merge['line_2'] = "Best before: " + df_merge['Best Before']
        df_merge['line_3'] = df_merge['Delivery Date'] + " Delivery"
        df_merge['line_4'] = 'Nutrition target per serving: ' + df_merge.NUTRITION
        df_merge['line_5'] = df_merge.portion_list
        df_merge['line_6'] = df_merge.dish_list
        
        return df_merge

def insert_background(slide, img_path):
    left = top = Inches(0)
    pic = slide.shapes.add_picture(img_path, left, top, width=slide.prs.slide_width, height=slide.prs.slide_height)

    # This moves it to the background
    slide.shapes._spTree.remove(pic._element)
    slide.shapes._spTree.insert(2, pic._element)
    return slide

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

def generate_ppt(df, prs, background_path):
    # Get the first slide as a template
    template_slide = prs.slides[0]
    text_type = template_slide.shapes[1].shape_type  # Identify the text shape type from template
    
    # Iterate each row in df_merge
    for index, row in df.iterrows():
        # Add one page copied from template
        new_slide = copy_slide(template_slide, prs)
        line_index = 0

        # Change text from line1 to line6
        for shape in new_slide.shapes:
            if shape.shape_type == text_type and shape.has_text_frame:
                # Get the text frame
                text_frame = shape.text_frame

                for paragraph in text_frame.paragraphs:
                    for run in paragraph.runs:
                        run.text = row['line_' + str(line_index+1)]
                        line_index += 1
                    break

        # Add background
        if background_path:
            new_slide = insert_background(new_slide, background_path)
    
    return prs

def new_database_access():
    return AirTable()

def generate_one_pagers(prs, background_path=None):
    # Initialize Airtable connection
    ac = new_database_access()
    
    # Process data from Airtable
    processed_data = ac.process_data()
    
    # Generate PowerPoint
    prs = generate_ppt(processed_data, prs, background_path)
    
    return prs

if __name__ == "__main__":
    # Paths to template and background
    template_path = 'template/one_pager_template.pptx'
    #background_path = 'templates/background.jpg'
    prs_file = Presentation(template_path)
    # Generate one pagers
    prs = generate_one_pagers(prs_file)
    
    # Save the result
    output_path = f'OnePager_{datetime.now().strftime("%Y%m%d_%H%M%S")}.pptx'
    prs.save(output_path)
    print(f"One pagers generated and saved as {output_path}")