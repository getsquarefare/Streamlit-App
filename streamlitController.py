import streamlit as st
import pandas as pd
from pptx import Presentation
from datetime import datetime
from shipping_sticker_generator import *
from squarespace_to_airtable import *

# Streamlit app
def main():
    st.title(':orange[SqaureFare] Toolkits 🧰')
    st.divider()
    st.header('Shipping Sticker Generator 🚚')
    current_date = datetime.now().strftime('%m%d%Y')

    # Display the shipping template file for double-checking
    st.subheader("Step1: Upload Shipping Sticker Template")
   
    # File uploader for the new template
    new_template_file = st.file_uploader(":green[Upload Template]", type="pptx")
    prs_file = Presentation(new_template_file)
    
    st.subheader("Step2: Upload Shipping Data")
    uploaded_shipping_file = st.file_uploader(":green[Upload shippingdata.csv]", type="csv")

    if uploaded_shipping_file is not None:
        df_shipping = process_shipping_data(uploaded_shipping_file)
        generate_button = st.button("Generate Stickers")

        if generate_button:
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

# squarespace_to_airtable
    st.divider()
    st.header('SqaureSpace Reports Sync to Airtable 📑')
    st.subheader("Step1: Products")
    product_file = st.file_uploader(":green[Upload products<XX-XX-XXXX>.csv]", type="csv")
    product_sync_button = st.button("Sync Products to Airtable ")
    if product_sync_button:
        product_final_df = product_data_clean(product_file)
        #product_final_df.to_csv('detailed_cleaned_subscription_products_data.csv', index=False)
        result = product_sync(product_final_df)
        st.markdown(result)

    
    st.subheader("Step2: Profile")
    profile_file = st.file_uploader(":green[Upload profile.csv]", type="csv")
    profile_sync_button = st.button("Sync Profile to Airtable ")
    if profile_sync_button:
        result = profile_sync(profile_file)
        st.markdown(result)
    
   
    st.subheader("Step3: Orders")
    order_file = st.file_uploader(":green[Upload orders.csv]", type="csv")
    order_sync_button = st.button("Sync Orders to Airtable ")
    if order_sync_button:
        result = orders_sync(order_file)
        st.markdown(result)

if __name__ == "__main__":
    main()
