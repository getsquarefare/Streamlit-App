import streamlit as st
import pandas as pd
from pptx import Presentation
from datetime import datetime
from shipping_sticker_generator import *

# Streamlit app
def main():
    st.title('Shipping Sticker Generator')
    current_date = datetime.now().strftime('%m%d%Y')

    # Display the shipping template file for double-checking
    st.subheader("Shipping Sticker Template")
   
    # File uploader for the new template
    new_template_file = st.file_uploader(":orange[Upload your shipping sticker template]", type="pptx")
    prs_file = Presentation(new_template_file)
    
    st.subheader("Shipping Data")
    uploaded_shipping_file = st.file_uploader(":orange[Upload your shippingdata.csv]", type="csv")

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

            # Fetch package recipent in Client
            package_recipient_df = fetch_package_recipient(sheet_id, client_sheet_name)

            # Attach package recipent to client servings
            updated_all_client_servings = add_recipient_clientservings(all_client_servings,package_recipient_df)

            # Match orders with shipping info
            final_match_result_with_portion_df = match_orders_to_shipping_data(updated_all_client_servings, df_shipping)

            # Generate the PowerPoint file
            prs = generate_ppt(final_match_result_with_portion_df, prs_file)

            # Save and download the PowerPoint file
            
            

            # Optional: Display the final DataFrame for checking
            st.subheader("Shipping Stickers")
            st.markdown(":orange[Stickers are ready! Click to download]")
            updated_ppt_name = f'{current_date}_shipping_sticker_updated.pptx'
            prs.save(updated_ppt_name)
            with open(updated_ppt_name, "rb") as file:
                st.download_button(
                    label="Download Stickers",
                    data=file,
                    file_name=updated_ppt_name,
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
            st.markdown(":orange[This is the final table which is fed to sticker ppt. For review purpose]")
            st.dataframe(final_match_result_with_portion_df)


if __name__ == "__main__":
    main()
