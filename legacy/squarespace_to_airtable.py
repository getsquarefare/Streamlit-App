import requests
import pandas as pd
from io import StringIO
import math
# API configuration
AIRTABLE_API_URL = 'https://api.airtable.com/v0/'
AIRTABLE_API_KEY = 'patYVk8Cd5TMVRvQr.9a6a92d2ed5c56581edb161ab316d2d2770d858a151acac0a6b47cbe7a8f05fd'

# Airtable table IDs, which can be obtained from table url
AIRTABLE_PRODUCT_TABLE_ID = 'appfSNkCNSk5pjPzM/tblIuOHMzC5kxVZjm/sync/oD12EB1F'
AIRTABLE_PROFILE_TABLE_ID = 'appfSNkCNSk5pjPzM/tblYDlcDc1wqPKCZS/sync/CWjTIeGS'
AIRTABLE_ORDER_TABLE_ID = 'appfSNkCNSk5pjPzM/tblkoGhSMHKvxOZAv/sync/uyEruRkq'

post_headers = {
    'Authorization': 'Bearer ' + AIRTABLE_API_KEY,
    'Content-Type': 'text/csv'
}

def product_data_clean(product_file):
    """
    Cleans and processes product data from a CSV file.

    Args:
        product_file (str): The path to the CSV file containing the product data.

    Returns:
        pandas.DataFrame: The cleaned and processed product data.
    """
    df_product = pd.read_csv(product_file)
    selected_columns = ['Title', 'Product Page', 'Visible', 'Categories']
    df_product[selected_columns] = df_product[selected_columns].fillna(method='ffill')

    df_product['Subscription?'] = df_product['Categories'].str.contains('/subscriptions', case=False, na=False) | \
                                   df_product['Title'].str.contains('Subscription', case=False, na=False) | \
                                   df_product['Title'].str.contains('-Day', case=False, na=False) | \
                                   df_product['Title'].str.contains('Jenny', case=False, na=False) | \
                                   df_product['Title'].str.contains('Custom Plan', case=False, na=False)

    df_product['Details'] = df_product.apply(lambda row: ';'.join(
        '{}:{}'.format(row[f'Option Name {i}'], row[f'Option Value {i}'])
        for i in range(1, 7) if pd.notnull(row[f'Option Name {i}'])
    ), axis=1)

    df_product_final = df_product[['SKU', 'Title', 'Subscription?', 'Details', 'Visible']].copy()

    # Initialize counts dictionary of meal breakdowns 
    data = {
        'SKU': [],
        'singleMeal_type/size': [],
        'singleMeal_protein': [],
        'subscr_# of breakfast': [],
        'subscr_# of lunch': [],
        'subscr_# of dinner': [],
        'subscr_# of snacks': [],
        'subscr_# of people': [],
        'subscr_meal_details': []
    }

    # Iterate over each row
    for index, row in df_product.iterrows():
        # Initialize counts for each row
        breakfast_count = 0
        lunch_count = 0
        dinner_count = 0
        snacks_count = 0
        people_count = 1
        days_count = 1
        subscr_meal_details = {}
        meal_type = ''
        protein = ''

        if row['Subscription?']:
            # Special custom plan
            special_plans = {
            "SQ2017246": (0, 5, 7, 6, 1, 1),
            "SQ3883851": (4, 4, 5, 0, 1, 1),
            "SQ4186073": (7, 14, 7, 14, 1, 1),
            "SQ0114848": (1, 4, 0, 6, 1, 1),
            "SQ5837440": (0, 5, 5, 0, 1, 1),
            "SQ6823138":(4, 2, 2, 4, 1, 1)
            }
            
            if row["SKU"] in special_plans:
                breakfast_count, lunch_count, dinner_count, snacks_count, people_count, days_count = special_plans[row["SKU"]]
            # Normal subscription
            # Iterate over each option name and value
            else:
                for i in range(1, 7):
                    option_name = row[f'Option Name {i}']
                    option_value = row[f'Option Value {i}']

                    if option_name == "Meals":
                        if option_value == 'Breakfast, Lunch, Dinner & Snacks':
                            breakfast_count = lunch_count = dinner_count = snacks_count = 1
                        elif option_value in ['Lunch & Dinner only', 'Breakfast and Snacks Not Included']:
                            lunch_count = dinner_count = 1

                    if option_name == '# of days':
                        days_count = int(option_value)

                    if option_name == '# people':
                        people_count = int(option_value)

                    if option_name in ['# of Lunches / Week']:
                        lunch_count = 1 if option_value == 'Lunch Included' else 0 if option_value == 'Lunch Not Included' else int(option_value)

                    if option_name in ['# of Dinners / Week', '# of dishes/ week']:
                        dinner_count = 1 if option_value == 'Dinner Included' else 0 if option_value == 'Dinner Not Included' else int(option_value)

                    if option_name == '# Days / Week':
                        days_count = int(option_value)
                    
                    if option_name == 'Lunch':
                        try:
                            lunch_count = int(option_value)
                        except ValueError:
                            if option_value is None or (isinstance(option_value, str) and option_value.lower() == 'none'):
                                lunch_count = 0
                            elif isinstance(option_value, float) and math.isnan(option_value):
                                lunch_count = 0
                            else:
                                lunch_count = 1
                                subscr_meal_details['Lunch'] = option_value
                    
                    if option_name == 'Dinner':
                        try:
                            dinner_count = int(option_value)
                        except ValueError:
                            if option_value is None or (isinstance(option_value, str) and option_value.lower() == 'none'):
                                dinner_count = 0
                            elif isinstance(option_value, float) and math.isnan(option_value):
                                dinner_count = 0
                            else:
                                dinner_count = 1
                                subscr_meal_details['Dinner'] = option_value

                    if option_name == 'Breakfast and snacks':
                        if option_value == 'Breakfast and Snacks Included':
                            try:
                                breakfast_count = snacks_count = 1
                            except ValueError:
                                if option_value is None or (isinstance(option_value, str) and option_value.lower() == 'none'):
                                    breakfast_count=snacks_count = 0
                                elif isinstance(option_value, float) and math.isnan(option_value):
                                    breakfast_count=snacks_count = 0
                                else:
                                    breakfast_count=snacks_count = 1

                    if option_name == 'Breakfast':
                        try:
                            breakfast_count = int(option_value)
                        except ValueError:
                            if option_value is None or (isinstance(option_value, str) and option_value.lower() == 'none'):
                                breakfast_count = 0
                            elif isinstance(option_value, float) and math.isnan(option_value):
                                breakfast_count = 0
                            else:
                                breakfast_count = 1
                                subscr_meal_details['Breakfast'] = option_value

                    if option_name == 'Snacks':
                        try:
                            snacks_count = int(option_value)
                        except ValueError:
                            if option_value is None or (isinstance(option_value, str) and option_value.lower() == 'none'):
                                snacks_count = 0
                            elif isinstance(option_value, float) and math.isnan(option_value):
                                snacks_count = 0
                            else:
                                snacks_count = 1
                                subscr_meal_details['Snacks'] = option_value

            
        else:
            # Non-subscription
            for i in range(1, 7):
                option_name = row[f'Option Name {i}']
                option_value = row[f'Option Value {i}']

                if option_name in ['Meal', 'Type of Meal', 'Meals', 'Size']:
                    meal_type = option_value

                if option_name == 'Protein':
                    protein = option_value

        # Append counts to the dictionary
        data['SKU'].append(row['SKU'])
        data['subscr_# of breakfast'].append(breakfast_count * days_count * people_count)
        data['subscr_# of lunch'].append(lunch_count * days_count * people_count)
        data['subscr_# of dinner'].append(dinner_count * days_count * people_count)
        data['subscr_# of snacks'].append(snacks_count * days_count * people_count)
        data['subscr_# of people'].append(people_count)
        data['subscr_meal_details'].append(subscr_meal_details)
        data['singleMeal_type/size'].append(meal_type)
        data['singleMeal_protein'].append(protein)

    # Create a new DataFrame from counts
    new_df = pd.DataFrame(data)
    df_product_final = df_product_final.merge(right=new_df, on='SKU', how='left')

    #df_product_final.to_csv('detailed_cleaned_subscription_products_data.csv', index=False)
    return df_product_final

def sync_to_airtable(df, table_id):
    csv_buffer = StringIO()

    # Order table needs the index to be the primary key in airtable
    if table_id ==AIRTABLE_ORDER_TABLE_ID:
        df.to_csv(csv_buffer, index=True)
        csv_content = "Index" + csv_buffer.getvalue()
    else:
        df.to_csv(csv_buffer, index=False)
        csv_content = csv_buffer.getvalue()
    post_url = AIRTABLE_API_URL + table_id

    response = requests.post(post_url, headers=post_headers, data=csv_content.encode('utf-8'))
    if response.status_code in [200, 202]:
        return f"Request successful: {response.json()}"
    else:
        return f"Request failed with status code: {response.status_code}\nResponse body: {response.text}"
    

def product_sync(df_product):
    return sync_to_airtable(df_product, AIRTABLE_PRODUCT_TABLE_ID)

def profile_sync(profile_file):
    df_profile = pd.read_csv(profile_file)
    df_profile.fillna('', inplace=True)
    return sync_to_airtable(df_profile, AIRTABLE_PROFILE_TABLE_ID)

def orders_sync(orders_file):
    df_order = pd.read_csv(orders_file)
    # List of columns except the two columns to fill
    columns_to_fill = df_order.columns.difference(['Checkout Form: Note / Additional Info + Name of referring nutritionist / fitness trainer if applicable. ','Lineitem variant'])
    # Apply forward fill only to the selected columns
    df_order[columns_to_fill] = df_order[columns_to_fill].fillna(method='ffill')
    df_order['Checkout Form: Note / Additional Info + Name of referring nutritionist / fitness trainer if applicable. '] = df_order.groupby('Order ID')['Checkout Form: Note / Additional Info + Name of referring nutritionist / fitness trainer if applicable. '].fillna(method='ffill')
    df_order.dropna(subset=['Email'], inplace=True)
    return sync_to_airtable(df_order, AIRTABLE_ORDER_TABLE_ID)