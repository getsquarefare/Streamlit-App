import streamlit as st
import pandas as pd
from pptx import Presentation
from datetime import datetime
from portionController import MealRecommendation  # Specific import
from shipping_sticker_generator import *
from shipping_sticker_generator_v2 import *
from store_access import new_database_access  # Add this import
import time
from clientservings_excel_output import *

# Streamlit app
def main():
    st.title('Square Fare Toolkits 🧰')
    current_date = datetime.now().strftime('%m%d%Y')

# portion algo trigger
    st.divider()
    st.header(':orange[Portion] Algorithm 🍽️')
    st.markdown("⚠️ Please first confirm :red[Open Orders] are :red[reviewed] and :red[approved]")
    st.markdown("⚠️ Currently :red[only] orders included in :red[this view] will be processed: [Open Orders > Running Portioning](https://airtable.com/appEe646yuQexwHJo/tblxT3Pg9Qh0BVZhM/viwrZHgdsYWnAMhtX?blocks=hide)")
    portion_generate_button = st.button("Yeh! Run Portioning Now")
    if portion_generate_button:
        # Display spinner and timer
        with st.spinner('Running the portioning algorithm... 🕐'):
            start_time = time.time()  # Record start time
            meal_recommendation = MealRecommendation()
            finishedCount, failedCount, failedCases = meal_recommendation.generate_recommendations_with_thread()
            end_time = time.time()
            elapsed_time = end_time - start_time

        st.success(f"{finishedCount} orders completed in {elapsed_time:.2f} seconds! ✅")
        if failedCount > 0:
            st.error(f"{failedCount} orders failed to process. Please review the following cases:")
            st.write(failedCases)

    # clientservings csv download
    st.divider()
    st.header(':green[ClientServings Excel] Generator 📑')
    st.markdown("⚠️ Source Table: [ClientServings](https://airtable.com/appEe646yuQexwHJo/tblVwpvUmsTS2Se51/viwfdnCtkFK4EGFM4?blocks=hide)")
    st.markdown("⚠️ If noticed any issues or missing data, please first check data in source table and then re-run the generator")
    clientservings_generate_button = st.button("Get ClientServings")
    if clientservings_generate_button:
        with st.spinner('Generating clientservings... It may take a few minutes 🕐'):
            updated_xlsx_name = f'{current_date}_clientservings.xlsx'

            airTable = new_database_access() 
            all_output = airTable.consolidated_all_dishes_output()
            excel_data = airTable.generate_clientservings_excel(all_output)

        # Allow the user to download the file
        st.download_button(
            label="Download ClientServings Excel",
            data=excel_data,
            file_name=updated_xlsx_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # shipping_sticker_generator_v2
    st.divider()
    st.header(':blue[Shipping Sticker] Generator 🚚')
    st.markdown("⚠️ Source Table: [Open Orders > All](https://airtable.com/appEe646yuQexwHJo/tblxT3Pg9Qh0BVZhM/viwICiVvohvm7Zfuo?blocks=hide)")
    st.markdown("⚠️ If noticed any issues or missing data, please first check data in source table and then re-run the generator")
    new_template_file_v2 = st.file_uploader(":blue[Optionally Upload template.pptx]", type="pptx",key="new_template_file_v2")
    if new_template_file_v2 is not None:
        prs_file = Presentation(new_template_file_v2)
    else:
        prs_file = Presentation('template/Shipping_Sticker_Template.pptx')
    shipping_sticker_generate_button = st.button("Try the new Shipping Sticker Generator")
    if shipping_sticker_generate_button:
        with st.spinner('Generating shipping stickers... It may take a few minutes 🕐'):
            prs = generate_shipping_stickers(prs_file)
            updated_ppt_name = f'{current_date}_shipping_sticker.pptx'
            prs.save(updated_ppt_name)
        with open(updated_ppt_name, "rb") as file:
            st.download_button(
                label="Download Stickers",
                data=file,
                file_name=updated_ppt_name,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )

if __name__ == "__main__":
    main()
