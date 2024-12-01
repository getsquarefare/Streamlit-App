import streamlit as st
import pandas as pd
from pptx import Presentation
from datetime import datetime
from portionController import MealRecommendation  # Specific import
from shipping_sticker_generator import *
from store_access import new_database_access  # Add this import
import time

# Streamlit app
def main():
    st.title(':orange[SqaureFare] Toolkits 🧰')
    current_date = datetime.now().strftime('%m%d%Y')

# portion algo trigger
    st.divider()
    st.header(':violet[Portion] Algorithm 🍽️')
    st.markdown("⚠️ Please first confirm Open Orders in Airtable are reviewed and approved")
    portion_generate_button = st.button("Yeh! Run Portioning Now")
    if portion_generate_button:
        # Display spinner and timer
        with st.spinner('Running the portioning algorithm... 🕐'):
            start_time = time.time()  # Record start time
            meal_recommendation = MealRecommendation()
            meal_recommendation.generate_recommendations()
            end_time = time.time()
            elapsed_time = end_time - start_time

        st.success(f"Portioning algorithm completed in {elapsed_time:.2f} seconds! ✅")

# clientservings csv download
    st.divider()
    st.header(':green[ClientServings Excel] Generator 📑')
    st.markdown("⚠️ Please first confirm the data in Airtable's clientservings table is correct")
    clientservings_generate_button = st.button("Get ClientServings")
    if clientservings_generate_button:
        ac = new_database_access()
        all_output = ac.consolidated_all_dishes_output()
        ac.export_clientservings_to_excel(all_output)

# shipping_sticker_generator
    st.divider()
    st.header(':blue[Shipping Sticker] Generator 🚚')
    

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
            prs = generate_ppt(final_match_result_with_portion_df, prs_file)

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


if __name__ == "__main__":
    main()
