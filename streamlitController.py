import streamlit as st
import pandas as pd
from pptx import Presentation
from datetime import datetime,timedelta
from portionController import MealRecommendation  # Specific import
from shipping_sticker_generator import *
from shipping_sticker_generator_v2 import *
from dish_sticker_generator_airtable import *
from one_pager_generator import *  # Import for one pager generation
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
    st.markdown("⚠️ Source Table: [Clientservings > Sorted](https://airtable.com/appEe646yuQexwHJo/tblVwpvUmsTS2Se51/viw5hROs9I9vV0YEq?blocks=hide)")
    st.markdown("⚠️ If noticed any issues or missing data, please first check data in source table and then re-run the generator")
    clientservings_generate_button = st.button("Get ClientServings")

    if clientservings_generate_button:
        try:
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
            
            # Success message
            st.success("Excel file generated successfully!")
            
        except AirTableError as e:
            # Handle specific AirTable errors
            st.error(f"Error generating ClientServings Excel: {str(e)}")
            logging.error(f"AirTable error: {str(e)}")
            
        except Exception as e:
            # Handle unexpected errors
            st.error("An unexpected error occurred. Please check the data in source table and try again.")

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
    st.markdown("⚠️ Source Table: [Open Orders > For Shipping Stickers](https://airtable.com/appEe646yuQexwHJo/tblxT3Pg9Qh0BVZhM/viwDpTtU0qaT9NcvG?blocks=hide)")
    st.markdown("⚠️ If noticed any issues or missing data, please first check data in source table and then re-run the generator")

    # Template uploader with error handling
    new_template_file_v2 = st.file_uploader(":blue[Optionally Upload Shipping_Sticker_Template.pptx]", type="pptx", key="new_template_file_v2")

    # Template validation
    try:
        if new_template_file_v2 is not None:
            try:
                prs_file = Presentation(new_template_file_v2)
                if len(prs_file.slides) == 0:
                    st.error("⚠️ The uploaded template has no slides. Please check your template file.")
                    valid_template = False
                else:
                    valid_template = True
            except Exception as e:
                st.error(f"⚠️ Invalid template file: {str(e)}")
                valid_template = False
        else:
            try:
                # Check if default template exists
                if not os.path.exists('template/Shipping_Sticker_Template.pptx'):
                    st.error("⚠️ Default template file not found. Please upload a template file.")
                    valid_template = False
                else:
                    prs_file = Presentation('template/Shipping_Sticker_Template.pptx')
                    if len(prs_file.slides) == 0:
                        st.error("⚠️ The default template has no slides. Please check your template file.")
                        valid_template = False
                    else:
                        valid_template = True
            except Exception as e:
                st.error(f"⚠️ Error loading default template: {str(e)}")
                valid_template = False
    except Exception as e:
        st.error(f"⚠️ Error during template preparation: {str(e)}")
        valid_template = False

    # Generate button - only show if we have a valid template
    if valid_template:
        shipping_sticker_generate_button = st.button("Generate Shipping Stickers")
        
        if shipping_sticker_generate_button:
            try:
                with st.spinner('Generating shipping stickers... It may take a few minutes 🕐'):
                    prs = generate_shipping_stickers(prs_file)
                    if prs and len(prs.slides) > 0:
                        updated_ppt_name = f'{current_date}_shipping_sticker.pptx'
                        prs.save(updated_ppt_name)
                        st.success(f"✅ Successfully generated {len(prs.slides)} shipping stickers!")
                        
                        # Provide download button
                        with open(updated_ppt_name, "rb") as file:
                            st.download_button(
                                label="Download Stickers",
                                data=file,
                                file_name=updated_ppt_name,
                                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                            )
                    else:
                        st.warning("⚠️ No stickers were generated. Please check if there are any open orders in Airtable.")
            except AirTableError as e:
                st.error(f"⚠️ Airtable Connection Error: {str(e)}")
                st.info("Please check your Airtable API key and connection settings.")
            except PPTGenerationError as e:
                st.error(f"⚠️ PowerPoint Generation Error: {str(e)}")
                st.info("There was a problem creating the PowerPoint file. Please check your template.")
            except ValueError as e:
                st.warning(f"⚠️ {str(e)}")
            except Exception as e:
                st.error(f"⚠️ Unexpected Error: {str(e)}")
                st.info("Please contact support with the error details and time of occurrence.")
                # Optionally show technical details in an expander
                with st.expander("Technical Error Details (for debugging)"):
                    st.code(traceback.format_exc())

    # dish_sticker_generator
    st.divider()
    st.header(':orange[Dish Sticker] Generator - Airtable 🥡 🍱')
    st.markdown("⚠️ Source Table: [Open Orders > For Dish Stickers](https://airtable.com/appEe646yuQexwHJo/tblxT3Pg9Qh0BVZhM/viwxYRtkq19KcXcZX?blocks=hide)")
    st.markdown("⚠️ If noticed any issues or missing data, please first check data in source table and then re-run the generator")

    # Template file handling with error checking
    try:
        new_dish_sticker_template = st.file_uploader(":blue[Optionally Upload a new Dish_Sticker_Template.pptx]", type="pptx", key="new_dish_sticker_template")
        
        if new_dish_sticker_template is not None:
            try:
                prs_file = Presentation(new_dish_sticker_template)
                st.success("Custom template loaded successfully!")
            except Exception as e:
                st.error(f"Error loading uploaded template: {str(e)}")
                st.info("Please upload a valid PowerPoint template file")
                st.stop()
        else:
            template_path = 'template/Dish_Sticker_Template.pptx'
            try:
                if not os.path.exists(template_path):
                    st.error(f"Default template not found at {template_path}")
                    st.info("Please upload a custom template to continue")
                    st.stop()
                prs_file = Presentation(template_path)
            except Exception as e:
                st.error(f"Error loading default template: {str(e)}")
                st.info("Please check that the template file exists and is not corrupted")
                st.stop()
    except Exception as e:
        st.error(f"File upload error: {str(e)}")
        st.stop()

    # Current date for filename
    current_date = datetime.now().strftime("%Y%m%d")

    # Generate button
    st.markdown("⚠️ Make sure you grouped orders by Dish ID and generate position id in Airtable, before running this generator")
    dish_sticker_generate_button = st.button("I have applied position id, lets get Dish Stickers")
    
    if dish_sticker_generate_button:
        try:
            with st.spinner('Generating dish stickers... It may take a few minutes 🕐'):
                # Check if the presentation has slides
                if len(prs_file.slides) == 0:
                    st.error("The template contains no slides. Please use a valid template.")
                    st.stop()
                
                # Generate stickers with error handling
                prs = generate_dish_stickers(prs_file)
                
                # Check if generation was successful
                if prs is None:
                    st.error("Failed to generate stickers. Please check the error messages above.")
                    st.stop()
                
                # Save the file with error handling
                updated_ppt_name = f'{current_date}_dish_sticker.pptx'
                try:
                    prs.save(updated_ppt_name)
                    st.success(f"Successfully saved {updated_ppt_name}")
                except Exception as e:
                    st.error(f"Failed to save PowerPoint file: {str(e)}")
                    st.info("Check if you have write permissions to the current directory")
                    st.stop()
            
            # Provide download button with error handling
            try:
                with open(updated_ppt_name, "rb") as file:
                    st.download_button(
                        label="Download Stickers",
                        data=file,
                        file_name=updated_ppt_name,
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    )
            except Exception as e:
                st.error(f"Error preparing download: {str(e)}")
                st.info(f"The file was saved as {updated_ppt_name} but cannot be downloaded directly. Please access it from your server.")
        except Exception as e:
            st.error(f"Unexpected error during sticker generation: {str(e)}")
            st.exception(e)  

    # one_pager_generator
    st.divider()
    st.header(':green[One-Sheeter] Generator - Airtable 📑')
    st.markdown("⚠️ Source Table: [Open Orders > For One-Sheeter](https://airtable.com/appEe646yuQexwHJo/tblxT3Pg9Qh0BVZhM/viwuVy9aN2LLZrcPF?blocks=hide)")
    st.markdown("⚠️ If noticed any issues or missing data, please first check data in source table and then re-run the generator")

    # File upload for template
    new_one_pager_template = st.file_uploader(":blue[Optionally Upload new One_Pager_Template.pptx]", 
                                            type="pptx", 
                                            key="new_one_pager_template",
                                            help="Include instruction slide as second slide")

    st.markdown("⚠️ Please make sure all meal stickers are generated before running this one-sheeter generator")
    # Generate button
    one_sheeter_generate_button = st.button("Yes I ran the meal sticker already, now lets generate One-Sheeter")

    if one_sheeter_generate_button:
        with st.spinner('Generating One-Sheeter... It may take a few minutes 🕐'):
            # Save uploaded files temporarily if needed
            template_path = 'template/One_Pager_Template_v2.pptx'  # Default path
            
            if new_one_pager_template is not None:
                template_path = f"temp_template_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pptx"
                with open(template_path, "wb") as f:
                    f.write(new_one_pager_template.getbuffer())
            
            # Load template
            prs_file = Presentation(template_path)
            
            # Check if template has at least 2 slides (display warning if not)
            if len(prs_file.slides) < 2:
                st.warning("Your template should include an instruction slide as the second slide. Proceeding without instructions.")
            
            # Generate one pagers (no background parameter)
            prs = generate_one_pagers(prs_file)
            
            # Save the result
            updated_ppt_name = f'{current_date}__one_sheeter.pptx'
            prs.save(updated_ppt_name)
            
            # Clean up temporary files
            if new_one_pager_template is not None and os.path.exists(template_path):
                os.remove(template_path)
            
        # Provide download link
        with open(updated_ppt_name, "rb") as file:
            st.download_button(
                label="Download One-Sheeter",
                data=file,
                file_name=updated_ppt_name,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
        
        st.success("One-Sheeter generated successfully!")
if __name__ == "__main__":
    main()
