import streamlit as st
import pandas as pd
from pptx import Presentation
from datetime import datetime,timedelta
from portionController import MealRecommendation  # Specific import
from shipping_sticker_generator import *
from shipping_sticker_generator_v2 import *
from store_access import new_database_access  # Add this import
import time
from clientservings_excel_output import *
import pytz

# Streamlit app
def main():
    st.title('Square Fare Toolkits 🧰')
    # Set the time zone to Eastern Standard Time (EST)
    est = pytz.timezone('US/Eastern')
    # Get the current date in EST
    current_date = datetime.now(est).strftime('%m%d%Y')

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
            elapsed_time_str = str(timedelta(seconds=elapsed_time))
            elapsed_time_str = elapsed_time_str.split(".")[0] + "." + elapsed_time_str.split(".")[1][:2]
        st.success(f"{finishedCount} orders completed in {elapsed_time_str} seconds! ✅")
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
    # shipping_sticker_generator_v1
    st.divider()
    st.header(':blue[Shipping Sticker] Generator - Google Sheets 🚚')
    st.markdown("⚠️ Source Table: Uploaded Sheet")
    # Display the shipping template file for double-checking
    st.subheader("Step1: Upload Shipping Sticker Template",divider="blue")
   
    # File uploader for the new template
    new_template_file = st.file_uploader(":blue[Upload template.ppt]", type="pptx")
    prs_file = Presentation(new_template_file)
    
    st.subheader("Step2: Upload Shipping Data",divider="blue")
    uploaded_shipping_file = st.file_uploader(":blue[Upload shippingdata.csv]", type="csv")

    if uploaded_shipping_file is not None:
        df_shipping = process_shipping_data(uploaded_shipping_file)

        sticker_generate_button = st.button("Generate Stickers")
        if sticker_generate_button:
            # Load client servings from Google Sheets
            sheet_id = "1rorOBlH_K9qH4L39KehvI_rYGHo7agNVdCDisWydEj8"
            LA_sheet_name = "ClientServings-LA"
            NY_sheet_name = "ClientServings"
            client_sheet_name = "Clients"

            LA_clientservings_df = fetch_client_servings(sheet_id, LA_sheet_name)
            NY_clientservings_df = fetch_client_servings(sheet_id, NY_sheet_name)

            # Combine LA and NY client servings
            all_client_servings = pd.concat([LA_clientservings_df, NY_clientservings_df], ignore_index=True)

            # Fetch package recipent from Client
            package_recipient_df = fetch_package_recipient(sheet_id, client_sheet_name)

            # Attach package recipent to client servings
            updated_all_client_servings = add_recipient_clientservings(all_client_servings,package_recipient_df)

            # Match orders with shipping info
            final_match_result_with_portion_df = match_orders_to_shipping_data(updated_all_client_servings, df_shipping)

            # Generate the PowerPoint file
            prs =generate_ppt_v1(final_match_result_with_portion_df, prs_file)

            # Optional: Display the final DataFrame for checking
            st.subheader("Shipping Stickers")
            st.markdown(":green[Stickers are ready! Click to download]")
            updated_ppt_name = f'{current_date}_shipping_sticker_updated.pptx'
            prs.save(updated_ppt_name)
            with open(updated_ppt_name, "rb") as file:
                st.download_button(
                    label="Download Stickers",
                    data=file,
                    file_name=updated_ppt_name,
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
            st.caption(":orange[This is the final table which is fed to sticker ppt, for review purpose]")
            st.dataframe(final_match_result_with_portion_df)


    # shipping_sticker_generator_v2
    st.divider()
    st.header(':blue[Shipping Sticker] Generator - Airtable 🚚')
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
