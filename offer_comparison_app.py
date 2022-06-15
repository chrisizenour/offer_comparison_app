import streamlit as st
import pandas as pd
from datetime import date, datetime
from PIL import Image
from comparison_to_excel import comparison_inputs_to_excel

def main():
    logo = Image.open('freedom_logo.png')
    st.set_page_config(
        page_title='Freedom PM & Sales Offer Comparison App',
        page_icon=logo,
        layout='wide'
    )

    logo_container = st.container()
    disclaimer_container = st.container()
    password_container = st.container()
    description_container = st.container()
    instruction_container = st.container()
    intro_info_container = st.container()
    property_container = st.container()
    common_container = st.container()
    offer_1_container = st.container()
    offer_2_container = st.container()
    offer_3_container = st.container()
    offer_4_container = st.container()
    offer_5_container = st.container()
    offer_6_container = st.container()
    offer_7_container = st.container()
    offer_8_container = st.container()
    offer_9_container = st.container()
    offer_10_container = st.container()
    offer_11_container = st.container()
    offer_12_container = st.container()

    with logo_container:
        col1, col2, col3 = st.columns([1, 2, 1])
        with col1:
            st.write('')
        with col2:
            st.image(logo)
        with col3:
            st.write('')

    with disclaimer_container:
        with st.expander('DISCLOSURES'):
            st.markdown(
                '''
                *These estimates are not guaranteed and may not include escrows. Escrow balances are reimbursed by the existing lender. Taxes, rents & association dues are pro-rated at settlement. Under Virginia Law, the seller's proceeds may not be available for up to 2 business days following recording of the deed. Seller acknowledges receipt of this statement.*
                ''')

    with password_container:
        password_guess = st.text_input('Enter a password to gain access to this app', key='password_guess')
        if password_guess != st.secrets['password']:
            st.stop()

    with description_container:
        with st.expander('App Description'):
            st.markdown(
                '''
                ##### Offer Comparison Tool
                - This application is used to compare different offers for a property
                - The output of this application will be an Excel workbook which shows the different offers side-by-side
                '''
            )

    with instruction_container:
        with st.expander('App Instructions'):
            st.markdown(
                '''
                This application is built into multiple, separated data input forms\n
                To ensure the form is populating correctly, open the Common Data Form to check to see that if the pre-set values are loaded\n
                If the pre-set values are not initialized, refresh the webpage, re-enter the password, and again check the Common Data Form to see if the pre-set values loaded\n
                Perform this process until the pre-set values load\n
                When entering values into percentage fields, enter the value as the percentage you want
                - For example, if the Listing Company Compensation percentage needs to be 2.25%, enter 2.25 into the number input field
                - Form Introduction Data
                    - Enter the name of the agent preparing the offer comparison, the date the comparison is being created, and the number of offers being compared
                    - Press the 'Submit Information' button
                - Property Data Form
                    - Enter data related to the property being offered for sale
                    - Press the 'Submit Property Information' button
                - Common Data Form
                    - Enter data that is common to all offers for the property
                    - Press the 'Submit Common Information' button
                - Offers 1 thru n Form
                    - For each offer being compared, enter data related to that particular Offer in that Offer's Form
                    - Press the 'Submit Offer (n)'s Information' button
                - After all data has been updated/entered and then submitted in their respective forms, press the 'Download Offer Comparison Form'
                - The app will generate an MS Excel wworkbook for the Offer Comparison Form and will be located in your downloads folder
                '''
            )

    if 'preparer' not in st.session_state:
        st.session_state['update_preparer'] = ''
        st.session_state['preparer'] = ''
        st.session_state['update_prep_date'] = date.today()
        st.session_state['prep_date'] = ''
        st.session_state['update_offer_qty'] = 1
        st.session_state['offer_qty'] = 0

        st.session_state['update_seller_name'] = ''
        st.session_state['seller_name'] = ''
        st.session_state['update_address'] = ''
        st.session_state['address'] = ''
        st.session_state['update_list_price'] = 0
        st.session_state['list_price'] = 0
        st.session_state['update_payoff_amt_first_trust'] = 0
        st.session_state['payoff_amt_first_trust'] = 0
        st.session_state['update_payoff_amt_second_trust'] = 0
        st.session_state['payoff_amt_second_trust'] = 0
        st.session_state['update_annual_tax_amt'] = 0
        st.session_state['annual_tax_amt'] = 0
        st.session_state['prorated_annual_tax_amt'] = 0.0
        st.session_state['annual_hoa_condo_fee_amt'] = 0
        st.session_state['update_annual_hoa_condo_fee_amt'] = 0
        st.session_state['prorated_annual_hoa_condo_fee_amt'] = 0.0

        st.session_state['update_listing_company_pct'] = 2.5
        st.session_state['listing_company_pct'] = 0.0
        st.session_state['update_selling_company_pct'] = 2.5
        st.session_state['selling_company_pct'] = 0.0
        st.session_state['update_processing_fee'] = 0
        st.session_state['processing_fee'] = 0
        st.session_state['update_settlement_fee'] = 450
        st.session_state['settlement_fee'] = 0
        st.session_state['update_deed_preparation_fee'] = 150
        st.session_state['deed_preparation_fee'] = 0
        st.session_state['update_lien_trust_release_fee'] = 100
        st.session_state['lien_trust_release_fee'] = 0
        st.session_state['update_lien_trust_release_qty'] = 1
        st.session_state['lien_trust_release_qty'] = 0
        st.session_state['update_recording_release_fee'] = 38
        st.session_state['recording_release_fee'] = 0
        st.session_state['update_recording_release_qty'] = 1
        st.session_state['recording_release_qty'] = 0
        st.session_state['update_grantors_tax_pct'] = 0.1
        st.session_state['grantors_tax_pct'] = 0.0
        st.session_state['update_congestion_tax_pct'] = 0.2
        st.session_state['congestion_tax_pct'] = 0.0
        st.session_state['update_pest_inspection_fee'] = 50
        st.session_state['pest_inspection_fee'] = 0
        st.session_state['update_poa_condo_disclosure_fee'] = 350
        st.session_state['poa_condo_disclosure_fee'] = 0

        st.session_state['update_offer_1_name'] = 'Offer 1'
        st.session_state['offer_1_name'] = ''
        st.session_state['update_offer_1_settlement_date'] = date.today()
        st.session_state['offer_1_settlement_date'] = ''
        st.session_state['update_offer_1_settlement_company'] = ''
        st.session_state['offer_1_settlement_company'] = ''
        st.session_state['update_offer_1_amt'] = 0
        st.session_state['offer_1_amt'] = 0
        st.session_state['update_offer_1_emd_amt'] = 0
        st.session_state['offer_1_emd_amt'] = 0
        st.session_state['offer_1_finance_type'] = 'Select Financing Type'
        st.session_state['update_offer_1_down_pmt_pct'] = 0.0
        st.session_state['offer_1_down_pmt_pct'] = 0.0
        st.session_state['offer_1_closing_subsidy_radio'] = 'Percent of Offer Amt (%)'
        st.session_state['offer_1_update_closing_subsidy_pct'] = 0.0
        st.session_state['offer_1_closing_subsidy_flat_amt'] = 0
        st.session_state['offer_1_closing_subsidy_amt'] = 0.0
        st.session_state['offer_1_home_inspection_check'] = False
        st.session_state['offer_1_home_inspection_value'] = ''
        st.session_state['offer_1_home_inspection_days'] = 0
        st.session_state['offer_1_home_inspection_days_string'] = ''
        st.session_state['offer_1_radon_inspection_check'] = False
        st.session_state['offer_1_radon_inspection_value'] = ''
        st.session_state['offer_1_radon_inspection_days'] = 0
        st.session_state['offer_1_radon_inspection_days_string'] = ''
        st.session_state['offer_1_septic_inspection_check'] = False
        st.session_state['offer_1_septic_inspection_value'] = ''
        st.session_state['offer_1_septic_inspection_days'] = 0
        st.session_state['offer_1_septic_inspection_days_string'] = ''
        st.session_state['offer_1_well_inspection_check'] = False
        st.session_state['offer_1_well_inspection_value'] = ''
        st.session_state['offer_1_well_inspection_days'] = 0
        st.session_state['offer_1_well_inspection_days_string'] = ''
        st.session_state['offer_1_financing_contingency_check'] = False
        st.session_state['offer_1_financing_contingency_value'] = ''
        st.session_state['offer_1_financing_contingency_days'] = 0
        st.session_state['offer_1_financing_contingency_days_string'] = ''
        st.session_state['offer_1_appraisal_contingency_check'] = False
        st.session_state['offer_1_appraisal_contingency_value'] = ''
        st.session_state['offer_1_appraisal_contingency_days'] = 0
        st.session_state['offer_1_appraisal_contingency_days_string'] = ''
        st.session_state['offer_1_home_sale_contingency_check'] = False
        st.session_state['offer_1_home_sale_contingency_value'] = ''
        st.session_state['offer_1_home_sale_contingency_days'] = 0
        st.session_state['offer_1_home_sale_contingency_days_string'] = ''
        st.session_state['offer_1_pre_occupancy_request'] = False
        st.session_state['offer_1_pre_occupancy_credit_to_seller_amt'] = 0
        st.session_state['offer_1_pre_occupancy_date'] = date.today()
        st.session_state['offer_1_post_occupancy_request'] = False
        st.session_state['offer_1_post_occupancy_cost_to_seller_amt'] = 0
        st.session_state['offer_1_post_occupancy_date'] = date.today()
        # st.session_state['offer_1_update_post_occupancy_date'] = date.today()

        st.session_state['update_offer_2_name'] = 'Offer 2'
        st.session_state['offer_2_name'] = ''
        st.session_state['update_offer_2_settlement_date'] = date.today()
        st.session_state['offer_2_settlement_date'] = ''
        st.session_state['update_offer_2_settlement_company'] = ''
        st.session_state['offer_2_settlement_company'] = ''
        st.session_state['update_offer_2_amt'] = 0
        st.session_state['offer_2_amt'] = 0
        st.session_state['update_offer_2_emd_amt'] = 0
        st.session_state['offer_2_emd_amt'] = 0
        st.session_state['offer_2_finance_type'] = 'Select Financing Type'
        st.session_state['update_offer_2_down_pmt_pct'] = 0.0
        st.session_state['offer_2_down_pmt_pct'] = 0.0
        st.session_state['offer_2_closing_subsidy_radio'] = 'Percent of Offer Amt (%)'
        st.session_state['offer_2_update_closing_subsidy_pct'] = 0.0
        st.session_state['offer_2_closing_subsidy_flat_amt'] = 0
        st.session_state['offer_2_closing_subsidy_amt'] = 0.0
        st.session_state['offer_2_home_inspection_check'] = False
        st.session_state['offer_2_home_inspection_value'] = ''
        st.session_state['offer_2_home_inspection_days'] = 0
        st.session_state['offer_2_home_inspection_days_string'] = ''
        st.session_state['offer_2_radon_inspection_check'] = False
        st.session_state['offer_2_radon_inspection_value'] = ''
        st.session_state['offer_2_radon_inspection_days'] = 0
        st.session_state['offer_2_radon_inspection_days_string'] = ''
        st.session_state['offer_2_septic_inspection_check'] = False
        st.session_state['offer_2_septic_inspection_value'] = ''
        st.session_state['offer_2_septic_inspection_days'] = 0
        st.session_state['offer_2_septic_inspection_days_string'] = ''
        st.session_state['offer_2_well_inspection_check'] = False
        st.session_state['offer_2_well_inspection_value'] = ''
        st.session_state['offer_2_well_inspection_days'] = 0
        st.session_state['offer_2_well_inspection_days_string'] = ''
        st.session_state['offer_2_financing_contingency_check'] = False
        st.session_state['offer_2_financing_contingency_value'] = ''
        st.session_state['offer_2_financing_contingency_days'] = 0
        st.session_state['offer_2_financing_contingency_days_string'] = ''
        st.session_state['offer_2_appraisal_contingency_check'] = False
        st.session_state['offer_2_appraisal_contingency_value'] = ''
        st.session_state['offer_2_appraisal_contingency_days'] = 0
        st.session_state['offer_2_appraisal_contingency_days_string'] = ''
        st.session_state['offer_2_home_sale_contingency_check'] = False
        st.session_state['offer_2_home_sale_contingency_value'] = ''
        st.session_state['offer_2_home_sale_contingency_days'] = 0
        st.session_state['offer_2_home_sale_contingency_days_string'] = ''
        st.session_state['offer_2_pre_occupancy_request'] = False
        st.session_state['offer_2_pre_occupancy_credit_to_seller_amt'] = 0
        st.session_state['offer_2_pre_occupancy_date'] = date.today()
        st.session_state['offer_2_post_occupancy_request'] = False
        st.session_state['offer_2_post_occupancy_cost_to_seller_amt'] = 0
        st.session_state['offer_2_post_occupancy_date'] = date.today()
        # st.session_state['offer_2_update_post_occupancy_date'] = date.today()

        st.session_state['update_offer_3_name'] = 'Offer 3'
        st.session_state['offer_3_name'] = ''
        st.session_state['update_offer_3_settlement_date'] = date.today()
        st.session_state['offer_3_settlement_date'] = ''
        st.session_state['update_offer_3_settlement_company'] = ''
        st.session_state['offer_3_settlement_company'] = ''
        st.session_state['update_offer_3_amt'] = 0
        st.session_state['offer_3_amt'] = 0
        st.session_state['update_offer_3_emd_amt'] = 0
        st.session_state['offer_3_emd_amt'] = 0
        st.session_state['offer_3_finance_type'] = 'Select Financing Type'
        st.session_state['update_offer_3_down_pmt_pct'] = 0.0
        st.session_state['offer_3_down_pmt_pct'] = 0.0
        st.session_state['offer_3_closing_subsidy_radio'] = 'Percent of Offer Amt (%)'
        st.session_state['offer_3_update_closing_subsidy_pct'] = 0.0
        st.session_state['offer_3_closing_subsidy_flat_amt'] = 0
        st.session_state['offer_3_closing_subsidy_amt'] = 0.0
        st.session_state['offer_3_home_inspection_check'] = False
        st.session_state['offer_3_home_inspection_value'] = ''
        st.session_state['offer_3_home_inspection_days'] = 0
        st.session_state['offer_3_home_inspection_days_string'] = ''
        st.session_state['offer_3_radon_inspection_check'] = False
        st.session_state['offer_3_radon_inspection_value'] = ''
        st.session_state['offer_3_radon_inspection_days'] = 0
        st.session_state['offer_3_radon_inspection_days_string'] = ''
        st.session_state['offer_3_septic_inspection_check'] = False
        st.session_state['offer_3_septic_inspection_value'] = ''
        st.session_state['offer_3_septic_inspection_days'] = 0
        st.session_state['offer_3_septic_inspection_days_string'] = ''
        st.session_state['offer_3_well_inspection_check'] = False
        st.session_state['offer_3_well_inspection_value'] = ''
        st.session_state['offer_3_well_inspection_days'] = 0
        st.session_state['offer_3_well_inspection_days_string'] = ''
        st.session_state['offer_3_financing_contingency_check'] = False
        st.session_state['offer_3_financing_contingency_value'] = ''
        st.session_state['offer_3_financing_contingency_days'] = 0
        st.session_state['offer_3_financing_contingency_days_string'] = ''
        st.session_state['offer_3_appraisal_contingency_check'] = False
        st.session_state['offer_3_appraisal_contingency_value'] = ''
        st.session_state['offer_3_appraisal_contingency_days'] = 0
        st.session_state['offer_3_appraisal_contingency_days_string'] = ''
        st.session_state['offer_3_home_sale_contingency_check'] = False
        st.session_state['offer_3_home_sale_contingency_value'] = ''
        st.session_state['offer_3_home_sale_contingency_days'] = 0
        st.session_state['offer_3_home_sale_contingency_days_string'] = ''
        st.session_state['offer_3_pre_occupancy_request'] = False
        st.session_state['offer_3_pre_occupancy_credit_to_seller_amt'] = 0
        st.session_state['offer_3_pre_occupancy_date'] = date.today()
        st.session_state['offer_3_post_occupancy_request'] = False
        st.session_state['offer_3_post_occupancy_cost_to_seller_amt'] = 0
        st.session_state['offer_3_post_occupancy_date'] = date.today()
        # st.session_state['offer_3_update_post_occupancy_date'] = date.today()

        st.session_state['update_offer_4_name'] = 'Offer 4'
        st.session_state['offer_4_name'] = ''
        st.session_state['update_offer_4_settlement_date'] = date.today()
        st.session_state['offer_4_settlement_date'] = ''
        st.session_state['update_offer_4_settlement_company'] = ''
        st.session_state['offer_4_settlement_company'] = ''
        st.session_state['update_offer_4_amt'] = 0
        st.session_state['offer_4_amt'] = 0
        st.session_state['update_offer_4_emd_amt'] = 0
        st.session_state['offer_4_emd_amt'] = 0
        st.session_state['offer_4_finance_type'] = 'Select Financing Type'
        st.session_state['update_offer_4_down_pmt_pct'] = 0.0
        st.session_state['offer_4_down_pmt_pct'] = 0.0
        st.session_state['offer_4_closing_subsidy_radio'] = 'Percent of Offer Amt (%)'
        st.session_state['offer_4_update_closing_subsidy_pct'] = 0.0
        st.session_state['offer_4_closing_subsidy_flat_amt'] = 0
        st.session_state['offer_4_closing_subsidy_amt'] = 0.0
        st.session_state['offer_4_home_inspection_check'] = False
        st.session_state['offer_4_home_inspection_value'] = ''
        st.session_state['offer_4_home_inspection_days'] = 0
        st.session_state['offer_4_home_inspection_days_string'] = ''
        st.session_state['offer_4_radon_inspection_check'] = False
        st.session_state['offer_4_radon_inspection_value'] = ''
        st.session_state['offer_4_radon_inspection_days'] = 0
        st.session_state['offer_4_radon_inspection_days_string'] = ''
        st.session_state['offer_4_septic_inspection_check'] = False
        st.session_state['offer_4_septic_inspection_value'] = ''
        st.session_state['offer_4_septic_inspection_days'] = 0
        st.session_state['offer_4_septic_inspection_days_string'] = ''
        st.session_state['offer_4_well_inspection_check'] = False
        st.session_state['offer_4_well_inspection_value'] = ''
        st.session_state['offer_4_well_inspection_days'] = 0
        st.session_state['offer_4_well_inspection_days_string'] = ''
        st.session_state['offer_4_financing_contingency_check'] = False
        st.session_state['offer_4_financing_contingency_value'] = ''
        st.session_state['offer_4_financing_contingency_days'] = 0
        st.session_state['offer_4_financing_contingency_days_string'] = ''
        st.session_state['offer_4_appraisal_contingency_check'] = False
        st.session_state['offer_4_appraisal_contingency_value'] = ''
        st.session_state['offer_4_appraisal_contingency_days'] = 0
        st.session_state['offer_4_appraisal_contingency_days_string'] = ''
        st.session_state['offer_4_home_sale_contingency_check'] = False
        st.session_state['offer_4_home_sale_contingency_value'] = ''
        st.session_state['offer_4_home_sale_contingency_days'] = 0
        st.session_state['offer_4_home_sale_contingency_days_string'] = ''
        st.session_state['offer_4_pre_occupancy_request'] = False
        st.session_state['offer_4_pre_occupancy_credit_to_seller_amt'] = 0
        st.session_state['offer_4_pre_occupancy_date'] = date.today()
        st.session_state['offer_4_post_occupancy_request'] = False
        st.session_state['offer_4_post_occupancy_cost_to_seller_amt'] = 0
        st.session_state['offer_4_post_occupancy_date'] = date.today()
        # st.session_state['offer_4_update_post_occupancy_date'] = date.today()

        st.session_state['update_offer_5_name'] = 'Offer 5'
        st.session_state['offer_5_name'] = ''
        st.session_state['update_offer_5_settlement_date'] = date.today()
        st.session_state['offer_5_settlement_date'] = ''
        st.session_state['update_offer_5_settlement_company'] = ''
        st.session_state['offer_5_settlement_company'] = ''
        st.session_state['update_offer_5_amt'] = 0
        st.session_state['offer_5_amt'] = 0
        st.session_state['update_offer_5_emd_amt'] = 0
        st.session_state['offer_5_emd_amt'] = 0
        st.session_state['offer_5_finance_type'] = 'Select Financing Type'
        st.session_state['update_offer_5_down_pmt_pct'] = 0.0
        st.session_state['offer_5_down_pmt_pct'] = 0.0
        st.session_state['offer_5_closing_subsidy_radio'] = 'Percent of Offer Amt (%)'
        st.session_state['offer_5_update_closing_subsidy_pct'] = 0.0
        st.session_state['offer_5_closing_subsidy_flat_amt'] = 0
        st.session_state['offer_5_closing_subsidy_amt'] = 0.0
        st.session_state['offer_5_home_inspection_check'] = False
        st.session_state['offer_5_home_inspection_value'] = ''
        st.session_state['offer_5_home_inspection_days'] = 0
        st.session_state['offer_5_home_inspection_days_string'] = ''
        st.session_state['offer_5_radon_inspection_check'] = False
        st.session_state['offer_5_radon_inspection_value'] = ''
        st.session_state['offer_5_radon_inspection_days'] = 0
        st.session_state['offer_5_radon_inspection_days_string'] = ''
        st.session_state['offer_5_septic_inspection_check'] = False
        st.session_state['offer_5_septic_inspection_value'] = ''
        st.session_state['offer_5_septic_inspection_days'] = 0
        st.session_state['offer_5_septic_inspection_days_string'] = ''
        st.session_state['offer_5_well_inspection_check'] = False
        st.session_state['offer_5_well_inspection_value'] = ''
        st.session_state['offer_5_well_inspection_days'] = 0
        st.session_state['offer_5_well_inspection_days_string'] = ''
        st.session_state['offer_5_financing_contingency_check'] = False
        st.session_state['offer_5_financing_contingency_value'] = ''
        st.session_state['offer_5_financing_contingency_days'] = 0
        st.session_state['offer_5_financing_contingency_days_string'] = ''
        st.session_state['offer_5_appraisal_contingency_check'] = False
        st.session_state['offer_5_appraisal_contingency_value'] = ''
        st.session_state['offer_5_appraisal_contingency_days'] = 0
        st.session_state['offer_5_appraisal_contingency_days_string'] = ''
        st.session_state['offer_5_home_sale_contingency_check'] = False
        st.session_state['offer_5_home_sale_contingency_value'] = ''
        st.session_state['offer_5_home_sale_contingency_days'] = 0
        st.session_state['offer_5_home_sale_contingency_days_string'] = ''
        st.session_state['offer_5_pre_occupancy_request'] = False
        st.session_state['offer_5_pre_occupancy_credit_to_seller_amt'] = 0
        st.session_state['offer_5_pre_occupancy_date'] = date.today()
        st.session_state['offer_5_post_occupancy_request'] = False
        st.session_state['offer_5_post_occupancy_cost_to_seller_amt'] = 0
        st.session_state['offer_5_post_occupancy_date'] = date.today()
        # st.session_state['offer_5_update_post_occupancy_date'] = date.today()

        st.session_state['update_offer_6_name'] = 'Offer 6'
        st.session_state['offer_6_name'] = ''
        st.session_state['update_offer_6_settlement_date'] = date.today()
        st.session_state['offer_6_settlement_date'] = ''
        st.session_state['update_offer_6_settlement_company'] = ''
        st.session_state['offer_6_settlement_company'] = ''
        st.session_state['update_offer_6_amt'] = 0
        st.session_state['offer_6_amt'] = 0
        st.session_state['update_offer_6_emd_amt'] = 0
        st.session_state['offer_6_emd_amt'] = 0
        st.session_state['offer_6_finance_type'] = 'Select Financing Type'
        st.session_state['update_offer_6_down_pmt_pct'] = 0.0
        st.session_state['offer_6_down_pmt_pct'] = 0.0
        st.session_state['offer_6_closing_subsidy_radio'] = 'Percent of Offer Amt (%)'
        st.session_state['offer_6_update_closing_subsidy_pct'] = 0.0
        st.session_state['offer_6_closing_subsidy_flat_amt'] = 0
        st.session_state['offer_6_closing_subsidy_amt'] = 0.0
        st.session_state['offer_6_home_inspection_check'] = False
        st.session_state['offer_6_home_inspection_value'] = ''
        st.session_state['offer_6_home_inspection_days'] = 0
        st.session_state['offer_6_home_inspection_days_string'] = ''
        st.session_state['offer_6_radon_inspection_check'] = False
        st.session_state['offer_6_radon_inspection_value'] = ''
        st.session_state['offer_6_radon_inspection_days'] = 0
        st.session_state['offer_6_radon_inspection_days_string'] = ''
        st.session_state['offer_6_septic_inspection_check'] = False
        st.session_state['offer_6_septic_inspection_value'] = ''
        st.session_state['offer_6_septic_inspection_days'] = 0
        st.session_state['offer_6_septic_inspection_days_string'] = ''
        st.session_state['offer_6_well_inspection_check'] = False
        st.session_state['offer_6_well_inspection_value'] = ''
        st.session_state['offer_6_well_inspection_days'] = 0
        st.session_state['offer_6_well_inspection_days_string'] = ''
        st.session_state['offer_6_financing_contingency_check'] = False
        st.session_state['offer_6_financing_contingency_value'] = ''
        st.session_state['offer_6_financing_contingency_days'] = 0
        st.session_state['offer_6_financing_contingency_days_string'] = ''
        st.session_state['offer_6_appraisal_contingency_check'] = False
        st.session_state['offer_6_appraisal_contingency_value'] = ''
        st.session_state['offer_6_appraisal_contingency_days'] = 0
        st.session_state['offer_6_appraisal_contingency_days_string'] = ''
        st.session_state['offer_6_home_sale_contingency_check'] = False
        st.session_state['offer_6_home_sale_contingency_value'] = ''
        st.session_state['offer_6_home_sale_contingency_days'] = 0
        st.session_state['offer_6_home_sale_contingency_days_string'] = ''
        st.session_state['offer_6_pre_occupancy_request'] = False
        st.session_state['offer_6_pre_occupancy_credit_to_seller_amt'] = 0
        st.session_state['offer_6_pre_occupancy_date'] = date.today()
        st.session_state['offer_6_post_occupancy_request'] = False
        st.session_state['offer_6_post_occupancy_cost_to_seller_amt'] = 0
        st.session_state['offer_6_post_occupancy_date'] = date.today()
        # st.session_state['offer_6_update_post_occupancy_date'] = date.today()

        st.session_state['update_offer_7_name'] = 'Offer 7'
        st.session_state['offer_7_name'] = ''
        st.session_state['update_offer_7_settlement_date'] = date.today()
        st.session_state['offer_7_settlement_date'] = ''
        st.session_state['update_offer_7_settlement_company'] = ''
        st.session_state['offer_7_settlement_company'] = ''
        st.session_state['update_offer_7_amt'] = 0
        st.session_state['offer_7_amt'] = 0
        st.session_state['update_offer_7_emd_amt'] = 0
        st.session_state['offer_7_emd_amt'] = 0
        st.session_state['offer_7_finance_type'] = 'Select Financing Type'
        st.session_state['update_offer_7_down_pmt_pct'] = 0.0
        st.session_state['offer_7_down_pmt_pct'] = 0.0
        st.session_state['offer_7_closing_subsidy_radio'] = 'Percent of Offer Amt (%)'
        st.session_state['offer_7_update_closing_subsidy_pct'] = 0.0
        st.session_state['offer_7_closing_subsidy_flat_amt'] = 0
        st.session_state['offer_7_closing_subsidy_amt'] = 0.0
        st.session_state['offer_7_home_inspection_check'] = False
        st.session_state['offer_7_home_inspection_value'] = ''
        st.session_state['offer_7_home_inspection_days'] = 0
        st.session_state['offer_7_home_inspection_days_string'] = ''
        st.session_state['offer_7_radon_inspection_check'] = False
        st.session_state['offer_7_radon_inspection_value'] = ''
        st.session_state['offer_7_radon_inspection_days'] = 0
        st.session_state['offer_7_radon_inspection_days_string'] = ''
        st.session_state['offer_7_septic_inspection_check'] = False
        st.session_state['offer_7_septic_inspection_value'] = ''
        st.session_state['offer_7_septic_inspection_days'] = 0
        st.session_state['offer_7_septic_inspection_days_string'] = ''
        st.session_state['offer_7_well_inspection_check'] = False
        st.session_state['offer_7_well_inspection_value'] = ''
        st.session_state['offer_7_well_inspection_days'] = 0
        st.session_state['offer_7_well_inspection_days_string'] = ''
        st.session_state['offer_7_financing_contingency_check'] = False
        st.session_state['offer_7_financing_contingency_value'] = ''
        st.session_state['offer_7_financing_contingency_days'] = 0
        st.session_state['offer_7_financing_contingency_days_string'] = ''
        st.session_state['offer_7_appraisal_contingency_check'] = False
        st.session_state['offer_7_appraisal_contingency_value'] = ''
        st.session_state['offer_7_appraisal_contingency_days'] = 0
        st.session_state['offer_7_appraisal_contingency_days_string'] = ''
        st.session_state['offer_7_home_sale_contingency_check'] = False
        st.session_state['offer_7_home_sale_contingency_value'] = ''
        st.session_state['offer_7_home_sale_contingency_days'] = 0
        st.session_state['offer_7_home_sale_contingency_days_string'] = ''
        st.session_state['offer_7_pre_occupancy_request'] = False
        st.session_state['offer_7_pre_occupancy_credit_to_seller_amt'] = 0
        st.session_state['offer_7_pre_occupancy_date'] = date.today()
        st.session_state['offer_7_post_occupancy_request'] = False
        st.session_state['offer_7_post_occupancy_cost_to_seller_amt'] = 0
        st.session_state['offer_7_post_occupancy_date'] = date.today()
        # st.session_state['offer_7_update_post_occupancy_date'] = date.today()

        st.session_state['update_offer_8_name'] = 'Offer 8'
        st.session_state['offer_8_name'] = ''
        st.session_state['update_offer_8_settlement_date'] = date.today()
        st.session_state['offer_8_settlement_date'] = ''
        st.session_state['update_offer_8_settlement_company'] = ''
        st.session_state['offer_8_settlement_company'] = ''
        st.session_state['update_offer_8_amt'] = 0
        st.session_state['offer_8_amt'] = 0
        st.session_state['update_offer_8_emd_amt'] = 0
        st.session_state['offer_8_emd_amt'] = 0
        st.session_state['offer_8_finance_type'] = 'Select Financing Type'
        st.session_state['update_offer_8_down_pmt_pct'] = 0.0
        st.session_state['offer_8_down_pmt_pct'] = 0.0
        st.session_state['offer_8_closing_subsidy_radio'] = 'Percent of Offer Amt (%)'
        st.session_state['offer_8_update_closing_subsidy_pct'] = 0.0
        st.session_state['offer_8_closing_subsidy_flat_amt'] = 0
        st.session_state['offer_8_closing_subsidy_amt'] = 0.0
        st.session_state['offer_8_home_inspection_check'] = False
        st.session_state['offer_8_home_inspection_value'] = ''
        st.session_state['offer_8_home_inspection_days'] = 0
        st.session_state['offer_8_home_inspection_days_string'] = ''
        st.session_state['offer_8_radon_inspection_check'] = False
        st.session_state['offer_8_radon_inspection_value'] = ''
        st.session_state['offer_8_radon_inspection_days'] = 0
        st.session_state['offer_8_radon_inspection_days_string'] = ''
        st.session_state['offer_8_septic_inspection_check'] = False
        st.session_state['offer_8_septic_inspection_value'] = ''
        st.session_state['offer_8_septic_inspection_days'] = 0
        st.session_state['offer_8_septic_inspection_days_string'] = ''
        st.session_state['offer_8_well_inspection_check'] = False
        st.session_state['offer_8_well_inspection_value'] = ''
        st.session_state['offer_8_well_inspection_days'] = 0
        st.session_state['offer_8_well_inspection_days_string'] = ''
        st.session_state['offer_8_financing_contingency_check'] = False
        st.session_state['offer_8_financing_contingency_value'] = ''
        st.session_state['offer_8_financing_contingency_days'] = 0
        st.session_state['offer_8_financing_contingency_days_string'] = ''
        st.session_state['offer_8_appraisal_contingency_check'] = False
        st.session_state['offer_8_appraisal_contingency_value'] = ''
        st.session_state['offer_8_appraisal_contingency_days'] = 0
        st.session_state['offer_8_appraisal_contingency_days_string'] = ''
        st.session_state['offer_8_home_sale_contingency_check'] = False
        st.session_state['offer_8_home_sale_contingency_value'] = ''
        st.session_state['offer_8_home_sale_contingency_days'] = 0
        st.session_state['offer_8_home_sale_contingency_days_string'] = ''
        st.session_state['offer_8_pre_occupancy_request'] = False
        st.session_state['offer_8_pre_occupancy_credit_to_seller_amt'] = 0
        st.session_state['offer_8_pre_occupancy_date'] = date.today()
        st.session_state['offer_8_post_occupancy_request'] = False
        st.session_state['offer_8_post_occupancy_cost_to_seller_amt'] = 0
        st.session_state['offer_8_post_occupancy_date'] = date.today()
        # st.session_state['offer_8_update_post_occupancy_date'] = date.today()

        st.session_state['update_offer_9_name'] = 'Offer 9'
        st.session_state['offer_9_name'] = ''
        st.session_state['update_offer_9_settlement_date'] = date.today()
        st.session_state['offer_9_settlement_date'] = ''
        st.session_state['update_offer_9_settlement_company'] = ''
        st.session_state['offer_9_settlement_company'] = ''
        st.session_state['update_offer_9_amt'] = 0
        st.session_state['offer_9_amt'] = 0
        st.session_state['update_offer_9_emd_amt'] = 0
        st.session_state['offer_9_emd_amt'] = 0
        st.session_state['offer_9_finance_type'] = 'Select Financing Type'
        st.session_state['update_offer_9_down_pmt_pct'] = 0.0
        st.session_state['offer_9_down_pmt_pct'] = 0.0
        st.session_state['offer_9_closing_subsidy_radio'] = 'Percent of Offer Amt (%)'
        st.session_state['offer_9_update_closing_subsidy_pct'] = 0.0
        st.session_state['offer_9_closing_subsidy_flat_amt'] = 0
        st.session_state['offer_9_closing_subsidy_amt'] = 0.0
        st.session_state['offer_9_home_inspection_check'] = False
        st.session_state['offer_9_home_inspection_value'] = ''
        st.session_state['offer_9_home_inspection_days'] = 0
        st.session_state['offer_9_home_inspection_days_string'] = ''
        st.session_state['offer_9_radon_inspection_check'] = False
        st.session_state['offer_9_radon_inspection_value'] = ''
        st.session_state['offer_9_radon_inspection_days'] = 0
        st.session_state['offer_9_radon_inspection_days_string'] = ''
        st.session_state['offer_9_septic_inspection_check'] = False
        st.session_state['offer_9_septic_inspection_value'] = ''
        st.session_state['offer_9_septic_inspection_days'] = 0
        st.session_state['offer_9_septic_inspection_days_string'] = ''
        st.session_state['offer_9_well_inspection_check'] = False
        st.session_state['offer_9_well_inspection_value'] = ''
        st.session_state['offer_9_well_inspection_days'] = 0
        st.session_state['offer_9_well_inspection_days_string'] = ''
        st.session_state['offer_9_financing_contingency_check'] = False
        st.session_state['offer_9_financing_contingency_value'] = ''
        st.session_state['offer_9_financing_contingency_days'] = 0
        st.session_state['offer_9_financing_contingency_days_string'] = ''
        st.session_state['offer_9_appraisal_contingency_check'] = False
        st.session_state['offer_9_appraisal_contingency_value'] = ''
        st.session_state['offer_9_appraisal_contingency_days'] = 0
        st.session_state['offer_9_appraisal_contingency_days_string'] = ''
        st.session_state['offer_9_home_sale_contingency_check'] = False
        st.session_state['offer_9_home_sale_contingency_value'] = ''
        st.session_state['offer_9_home_sale_contingency_days'] = 0
        st.session_state['offer_9_home_sale_contingency_days_string'] = ''
        st.session_state['offer_9_pre_occupancy_request'] = False
        st.session_state['offer_9_pre_occupancy_credit_to_seller_amt'] = 0
        st.session_state['offer_9_pre_occupancy_date'] = date.today()
        st.session_state['offer_9_post_occupancy_request'] = False
        st.session_state['offer_9_post_occupancy_cost_to_seller_amt'] = 0
        st.session_state['offer_9_post_occupancy_date'] = date.today()
        # st.session_state['offer_9_update_post_occupancy_date'] = date.today()

        st.session_state['update_offer_10_name'] = 'Offer 10'
        st.session_state['offer_10_name'] = ''
        st.session_state['update_offer_10_settlement_date'] = date.today()
        st.session_state['offer_10_settlement_date'] = ''
        st.session_state['update_offer_10_settlement_company'] = ''
        st.session_state['offer_10_settlement_company'] = ''
        st.session_state['update_offer_10_amt'] = 0
        st.session_state['offer_10_amt'] = 0
        st.session_state['update_offer_10_emd_amt'] = 0
        st.session_state['offer_10_emd_amt'] = 0
        st.session_state['offer_10_finance_type'] = 'Select Financing Type'
        st.session_state['update_offer_10_down_pmt_pct'] = 0.0
        st.session_state['offer_10_down_pmt_pct'] = 0.0
        st.session_state['offer_10_closing_subsidy_radio'] = 'Percent of Offer Amt (%)'
        st.session_state['offer_10_update_closing_subsidy_pct'] = 0.0
        st.session_state['offer_10_closing_subsidy_flat_amt'] = 0
        st.session_state['offer_10_closing_subsidy_amt'] = 0.0
        st.session_state['offer_10_home_inspection_check'] = False
        st.session_state['offer_10_home_inspection_value'] = ''
        st.session_state['offer_10_home_inspection_days'] = 0
        st.session_state['offer_10_home_inspection_days_string'] = ''
        st.session_state['offer_10_radon_inspection_check'] = False
        st.session_state['offer_10_radon_inspection_value'] = ''
        st.session_state['offer_10_radon_inspection_days'] = 0
        st.session_state['offer_10_radon_inspection_days_string'] = ''
        st.session_state['offer_10_septic_inspection_check'] = False
        st.session_state['offer_10_septic_inspection_value'] = ''
        st.session_state['offer_10_septic_inspection_days'] = 0
        st.session_state['offer_10_septic_inspection_days_string'] = ''
        st.session_state['offer_10_well_inspection_check'] = False
        st.session_state['offer_10_well_inspection_value'] = ''
        st.session_state['offer_10_well_inspection_days'] = 0
        st.session_state['offer_10_well_inspection_days_string'] = ''
        st.session_state['offer_10_financing_contingency_check'] = False
        st.session_state['offer_10_financing_contingency_value'] = ''
        st.session_state['offer_10_financing_contingency_days'] = 0
        st.session_state['offer_10_financing_contingency_days_string'] = ''
        st.session_state['offer_10_appraisal_contingency_check'] = False
        st.session_state['offer_10_appraisal_contingency_value'] = ''
        st.session_state['offer_10_appraisal_contingency_days'] = 0
        st.session_state['offer_10_appraisal_contingency_days_string'] = ''
        st.session_state['offer_10_home_sale_contingency_check'] = False
        st.session_state['offer_10_home_sale_contingency_value'] = ''
        st.session_state['offer_10_home_sale_contingency_days'] = 0
        st.session_state['offer_10_home_sale_contingency_days_string'] = ''
        st.session_state['offer_10_pre_occupancy_request'] = False
        st.session_state['offer_10_pre_occupancy_credit_to_seller_amt'] = 0
        st.session_state['offer_10_pre_occupancy_date'] = date.today()
        st.session_state['offer_10_post_occupancy_request'] = False
        st.session_state['offer_10_post_occupancy_cost_to_seller_amt'] = 0
        st.session_state['offer_10_post_occupancy_date'] = date.today()
        # st.session_state['offer_10_update_post_occupancy_date'] = date.today()

        st.session_state['update_offer_11_name'] = 'Offer 11'
        st.session_state['offer_11_name'] = ''
        st.session_state['update_offer_11_settlement_date'] = date.today()
        st.session_state['offer_11_settlement_date'] = ''
        st.session_state['update_offer_11_settlement_company'] = ''
        st.session_state['offer_11_settlement_company'] = ''
        st.session_state['update_offer_11_amt'] = 0
        st.session_state['offer_11_amt'] = 0
        st.session_state['update_offer_11_emd_amt'] = 0
        st.session_state['offer_11_emd_amt'] = 0
        st.session_state['offer_11_finance_type'] = 'Select Financing Type'
        st.session_state['update_offer_11_down_pmt_pct'] = 0.0
        st.session_state['offer_11_down_pmt_pct'] = 0.0
        st.session_state['offer_11_closing_subsidy_radio'] = 'Percent of Offer Amt (%)'
        st.session_state['offer_11_update_closing_subsidy_pct'] = 0.0
        st.session_state['offer_11_closing_subsidy_flat_amt'] = 0
        st.session_state['offer_11_closing_subsidy_amt'] = 0.0
        st.session_state['offer_11_home_inspection_check'] = False
        st.session_state['offer_11_home_inspection_value'] = ''
        st.session_state['offer_11_home_inspection_days'] = 0
        st.session_state['offer_11_home_inspection_days_string'] = ''
        st.session_state['offer_11_radon_inspection_check'] = False
        st.session_state['offer_11_radon_inspection_value'] = ''
        st.session_state['offer_11_radon_inspection_days'] = 0
        st.session_state['offer_11_radon_inspection_days_string'] = ''
        st.session_state['offer_11_septic_inspection_check'] = False
        st.session_state['offer_11_septic_inspection_value'] = ''
        st.session_state['offer_11_septic_inspection_days'] = 0
        st.session_state['offer_11_septic_inspection_days_string'] = ''
        st.session_state['offer_11_well_inspection_check'] = False
        st.session_state['offer_11_well_inspection_value'] = ''
        st.session_state['offer_11_well_inspection_days'] = 0
        st.session_state['offer_11_well_inspection_days_string'] = ''
        st.session_state['offer_11_financing_contingency_check'] = False
        st.session_state['offer_11_financing_contingency_value'] = ''
        st.session_state['offer_11_financing_contingency_days'] = 0
        st.session_state['offer_11_financing_contingency_days_string'] = ''
        st.session_state['offer_11_appraisal_contingency_check'] = False
        st.session_state['offer_11_appraisal_contingency_value'] = ''
        st.session_state['offer_11_appraisal_contingency_days'] = 0
        st.session_state['offer_11_appraisal_contingency_days_string'] = ''
        st.session_state['offer_11_home_sale_contingency_check'] = False
        st.session_state['offer_11_home_sale_contingency_value'] = ''
        st.session_state['offer_11_home_sale_contingency_days'] = 0
        st.session_state['offer_11_home_sale_contingency_days_string'] = ''
        st.session_state['offer_11_pre_occupancy_request'] = False
        st.session_state['offer_11_pre_occupancy_credit_to_seller_amt'] = 0
        st.session_state['offer_11_pre_occupancy_date'] = date.today()
        st.session_state['offer_11_post_occupancy_request'] = False
        st.session_state['offer_11_post_occupancy_cost_to_seller_amt'] = 0
        st.session_state['offer_11_post_occupancy_date'] = date.today()
        # st.session_state['offer_11_update_post_occupancy_date'] = date.today()

        st.session_state['update_offer_12_name'] = 'Offer 12'
        st.session_state['offer_12_name'] = ''
        st.session_state['update_offer_12_settlement_date'] = date.today()
        st.session_state['offer_12_settlement_date'] = ''
        st.session_state['update_offer_12_settlement_company'] = ''
        st.session_state['offer_12_settlement_company'] = ''
        st.session_state['update_offer_12_amt'] = 0
        st.session_state['offer_12_amt'] = 0
        st.session_state['update_offer_12_emd_amt'] = 0
        st.session_state['offer_12_emd_amt'] = 0
        st.session_state['offer_12_finance_type'] = 'Select Financing Type'
        st.session_state['update_offer_12_down_pmt_pct'] = 0.0
        st.session_state['offer_12_down_pmt_pct'] = 0.0
        st.session_state['offer_12_closing_subsidy_radio'] = 'Percent of Offer Amt (%)'
        st.session_state['offer_12_update_closing_subsidy_pct'] = 0.0
        st.session_state['offer_12_closing_subsidy_flat_amt'] = 0
        st.session_state['offer_12_closing_subsidy_amt'] = 0.0
        st.session_state['offer_12_home_inspection_check'] = False
        st.session_state['offer_12_home_inspection_value'] = ''
        st.session_state['offer_12_home_inspection_days'] = 0
        st.session_state['offer_12_home_inspection_days_string'] = ''
        st.session_state['offer_12_radon_inspection_check'] = False
        st.session_state['offer_12_radon_inspection_value'] = ''
        st.session_state['offer_12_radon_inspection_days'] = 0
        st.session_state['offer_12_radon_inspection_days_string'] = ''
        st.session_state['offer_12_septic_inspection_check'] = False
        st.session_state['offer_12_septic_inspection_value'] = ''
        st.session_state['offer_12_septic_inspection_days'] = 0
        st.session_state['offer_12_septic_inspection_days_string'] = ''
        st.session_state['offer_12_well_inspection_check'] = False
        st.session_state['offer_12_well_inspection_value'] = ''
        st.session_state['offer_12_well_inspection_days'] = 0
        st.session_state['offer_12_well_inspection_days_string'] = ''
        st.session_state['offer_12_financing_contingency_check'] = False
        st.session_state['offer_12_financing_contingency_value'] = ''
        st.session_state['offer_12_financing_contingency_days'] = 0
        st.session_state['offer_12_financing_contingency_days_string'] = ''
        st.session_state['offer_12_appraisal_contingency_check'] = False
        st.session_state['offer_12_appraisal_contingency_value'] = ''
        st.session_state['offer_12_appraisal_contingency_days'] = 0
        st.session_state['offer_12_appraisal_contingency_days_string'] = ''
        st.session_state['offer_12_home_sale_contingency_check'] = False
        st.session_state['offer_12_home_sale_contingency_value'] = ''
        st.session_state['offer_12_home_sale_contingency_days'] = 0
        st.session_state['offer_12_home_sale_contingency_days_string'] = ''
        st.session_state['offer_12_pre_occupancy_request'] = False
        st.session_state['offer_12_pre_occupancy_credit_to_seller_amt'] = 0
        st.session_state['offer_12_pre_occupancy_date'] = date.today()
        st.session_state['offer_12_post_occupancy_request'] = False
        st.session_state['offer_12_post_occupancy_cost_to_seller_amt'] = 0
        st.session_state['offer_12_post_occupancy_date'] = date.today()
        # st.session_state['offer_12_update_post_occupancy_date'] = date.today()


    contingencies = ['Home Inspection', 'Financing', 'Appraisal', 'Pest Inspection']
    financing_types = ['Select Financing Type', 'Cash', 'Conventional', 'FHA', 'VA', 'USDA', 'Other']


    def update_intro_info_form():
        st.session_state.preparer = st.session_state.update_preparer
        st.session_state.prep_date = st.session_state.update_prep_date
        st.session_state.offer_qty = st.session_state.update_offer_qty


    def update_property_info_form():
        st.session_state.seller_name = st.session_state.update_seller_name
        st.session_state.address = st.session_state.update_address
        st.session_state.list_price = st.session_state.update_list_price
        st.session_state.payoff_amt_first_trust = st.session_state.update_payoff_amt_first_trust
        st.session_state.payoff_amt_second_trust = st.session_state.update_payoff_amt_second_trust
        st.session_state.annual_tax_amt = st.session_state.update_annual_tax_amt
        st.session_state.prorated_annual_tax_amt = st.session_state.annual_tax_amt / 12 * 3
        st.session_state.annual_hoa_condo_fee_amt = st.session_state.update_annual_hoa_condo_fee_amt
        st.session_state.prorated_annual_hoa_condo_fee_amt = st.session_state.annual_hoa_condo_fee_amt / 12 * 3


    def update_common_info_form():
        st.session_state.listing_company_pct = st.session_state.update_listing_company_pct / 100
        st.session_state.selling_company_pct = st.session_state.update_selling_company_pct / 100
        st.session_state.processing_fee = st.session_state.update_processing_fee
        st.session_state.settlement_fee = st.session_state.update_settlement_fee
        st.session_state.deed_preparation_fee = st.session_state.update_deed_preparation_fee
        st.session_state.lien_trust_release_fee = st.session_state.update_lien_trust_release_fee
        st.session_state.lien_trust_release_qty = st.session_state.update_lien_trust_release_qty
        st.session_state.recording_release_fee = st.session_state.update_recording_release_fee
        st.session_state.recording_release_qty = st.session_state.update_recording_release_qty
        st.session_state.grantors_tax_pct = st.session_state.update_grantors_tax_pct / 100
        st.session_state.congestion_tax_pct = st.session_state.update_congestion_tax_pct / 100
        st.session_state.pest_inspection_fee = st.session_state.update_pest_inspection_fee
        st.session_state.poa_condo_disclosure_fee = st.session_state.update_poa_condo_disclosure_fee


    def days_int_to_string(x):
        if x == 0:
            string_value = ''
        return string_value


    def update_offer_1_info_form():
        st.session_state.offer_1_name = st.session_state.update_offer_1_name
        st.session_state.offer_1_settlement_date = st.session_state.update_offer_1_settlement_date
        st.session_state.offer_1_settlement_company = st.session_state.update_offer_1_settlement_company
        st.session_state.offer_1_amt = st.session_state.update_offer_1_amt
        st.session_state.offer_1_emd_amt = st.session_state.update_offer_1_emd_amt
        st.session_state.offer_1_down_pmt_pct = st.session_state.update_offer_1_down_pmt_pct
        st.session_state.offer_1_down_pmt_pct = st.session_state.offer_1_down_pmt_pct / 100
        st.session_state.offer_1_closing_subsidy_pct = st.session_state.offer_1_update_closing_subsidy_pct / 100
        if st.session_state.offer_1_closing_subsidy_radio == 'Percent of Offer Amt (%)':
            st.session_state.offer_1_closing_subsidy_amt = st.session_state.offer_1_closing_subsidy_pct * st.session_state.offer_1_amt
        else:
            st.session_state.offer_1_closing_subsidy_amt = st.session_state.offer_1_closing_subsidy_flat_amt

        if st.session_state.offer_1_home_inspection_check:
            st.session_state.offer_1_home_inspection_value = 'Y'
            st.session_state.offer_1_home_inspection_days = st.session_state.offer_1_home_inspection_days
            st.session_state.offer_1_home_inspection_days_string = st.session_state.offer_1_home_inspection_days
        else:
            st.session_state.offer_1_home_inspection_value = ''
            st.session_state.offer_1_home_inspection_days = 0
            st.session_state.offer_1_home_inspection_days_string = days_int_to_string(st.session_state.offer_1_home_inspection_days)

        if st.session_state.offer_1_radon_inspection_check:
            st.session_state.offer_1_radon_inspection_value = 'Y'
            st.session_state.offer_1_radon_inspection_days = st.session_state.offer_1_radon_inspection_days
            st.session_state.offer_1_radon_inspection_days_string = st.session_state.offer_1_radon_inspection_days
        else:
            st.session_state.offer_1_radon_inspection_value = ''
            st.session_state.offer_1_radon_inspection_days = 0
            st.session_state.offer_1_radon_inspection_days_string = days_int_to_string(st.session_state.offer_1_radon_inspection_days)

        if st.session_state.offer_1_septic_inspection_check:
            st.session_state.offer_1_septic_inspection_value = 'Y'
            st.session_state.offer_1_septic_inspection_days = st.session_state.offer_1_septic_inspection_days
            st.session_state.offer_1_septic_inspection_days_string = st.session_state.offer_1_septic_inspection_days
        else:
            st.session_state.offer_1_septic_inspection_value = ''
            st.session_state.offer_1_septic_inspection_days = 0
            st.session_state.offer_1_septic_inspection_days_string = days_int_to_string(st.session_state.offer_1_septic_inspection_days)

        if st.session_state.offer_1_well_inspection_check:
            st.session_state.offer_1_well_inspection_value = 'Y'
            st.session_state.offer_1_well_inspection_days = st.session_state.offer_1_well_inspection_days
            st.session_state.offer_1_well_inspection_days_string = st.session_state.offer_1_well_inspection_days
        else:
            st.session_state.offer_1_well_inspection_value = ''
            st.session_state.offer_1_well_inspection_days = 0
            st.session_state.offer_1_well_inspection_days_string = days_int_to_string(st.session_state.offer_1_well_inspection_days)

        if st.session_state.offer_1_financing_contingency_check:
            st.session_state.offer_1_financing_contingency_value = 'Y'
            st.session_state.offer_1_financing_contingency_days = st.session_state.offer_1_financing_contingency_days
            st.session_state.offer_1_financing_contingency_days_string = st.session_state.offer_1_financing_contingency_days
        else:
            st.session_state.offer_1_financing_contingency_value = ''
            st.session_state.offer_1_financing_contingency_days = 0
            st.session_state.offer_1_financing_contingency_days_string = days_int_to_string(st.session_state.offer_1_financing_contingency_days)

        if st.session_state.offer_1_appraisal_contingency_check:
            st.session_state.offer_1_appraisal_contingency_value = 'Y'
            st.session_state.offer_1_appraisal_contingency_days = st.session_state.offer_1_appraisal_contingency_days
            st.session_state.offer_1_appraisal_contingency_days_string = st.session_state.offer_1_appraisal_contingency_days
        else:
            st.session_state.offer_1_appraisal_contingency_value = ''
            st.session_state.offer_1_appraisal_contingency_days = 0
            st.session_state.offer_1_appraisal_contingency_days_string = days_int_to_string(st.session_state.offer_1_appraisal_contingency_days)

        if st.session_state.offer_1_home_sale_contingency_check:
            st.session_state.offer_1_home_sale_contingency_value = 'Y'
            st.session_state.offer_1_home_sale_contingency_days = st.session_state.offer_1_home_inspection_days
            st.session_state.offer_1_home_sale_contingency_days_string = st.session_state.offer_1_home_sale_contingency_days
        else:
            st.session_state.offer_1_home_sale_contingency_value = ''
            st.session_state.offer_1_home_sale_contingency_days = 0
            st.session_state.offer_1_home_sale_contingency_days_string = days_int_to_string(
                st.session_state.offer_1_home_sale_contingency_days)

        if st.session_state.offer_1_pre_occupancy_request:
            st.session_state.offer_1_pre_occupancy_date = st.session_state.offer_1_update_pre_occupancy_date
        else:
            st.session_state.offer_1_pre_occupancy_date = ''

        if st.session_state.offer_1_post_occupancy_request:
            st.session_state.offer_1_post_occupancy_date = st.session_state.offer_1_update_post_occupancy_date
        else:
            st.session_state.offer_1_post_occupancy_date = ''

    def update_offer_2_info_form():
        st.session_state.offer_2_name = st.session_state.update_offer_2_name
        st.session_state.offer_2_settlement_date = st.session_state.update_offer_2_settlement_date
        st.session_state.offer_2_settlement_company = st.session_state.update_offer_2_settlement_company
        st.session_state.offer_2_amt = st.session_state.update_offer_2_amt
        st.session_state.offer_2_emd_amt = st.session_state.update_offer_2_emd_amt
        st.session_state.offer_2_down_pmt_pct = st.session_state.update_offer_2_down_pmt_pct
        st.session_state.offer_2_down_pmt_pct = st.session_state.offer_2_down_pmt_pct / 100
        st.session_state.offer_2_closing_subsidy_pct = st.session_state.offer_2_update_closing_subsidy_pct / 100
        if st.session_state.offer_2_closing_subsidy_radio == 'Percent of Offer Amt (%)':
            st.session_state.offer_2_closing_subsidy_amt = st.session_state.offer_2_closing_subsidy_pct * st.session_state.offer_2_amt
        else:
            st.session_state.offer_2_closing_subsidy_amt = st.session_state.offer_2_closing_subsidy_flat_amt
        if st.session_state.offer_2_home_inspection_check:
            st.session_state.offer_2_home_inspection_value = 'Y'
            st.session_state.offer_2_home_inspection_days = st.session_state.offer_2_home_inspection_days
            st.session_state.offer_2_home_inspection_days_string = st.session_state.offer_2_home_inspection_days
        else:
            st.session_state.offer_2_home_inspection_value = ''
            st.session_state.offer_2_home_inspection_days = 0
            st.session_state.offer_2_home_inspection_days_string = days_int_to_string(st.session_state.offer_2_home_inspection_days)
        if st.session_state.offer_2_radon_inspection_check:
            st.session_state.offer_2_radon_inspection_value = 'Y'
            st.session_state.offer_2_radon_inspection_days = st.session_state.offer_2_radon_inspection_days
            st.session_state.offer_2_radon_inspection_days_string = st.session_state.offer_2_radon_inspection_days
        else:
            st.session_state.offer_2_radon_inspection_value = ''
            st.session_state.offer_2_radon_inspection_days = 0
            st.session_state.offer_2_radon_inspection_days_string = days_int_to_string(st.session_state.offer_2_radon_inspection_days)
        if st.session_state.offer_2_septic_inspection_check:
            st.session_state.offer_2_septic_inspection_value = 'Y'
            st.session_state.offer_2_septic_inspection_days = st.session_state.offer_2_septic_inspection_days
            st.session_state.offer_2_septic_inspection_days_string = st.session_state.offer_2_septic_inspection_days
        else:
            st.session_state.offer_2_septic_inspection_value = ''
            st.session_state.offer_2_septic_inspection_days = 0
            st.session_state.offer_2_septic_inspection_days_string = days_int_to_string(st.session_state.offer_2_septic_inspection_days)
        if st.session_state.offer_2_well_inspection_check:
            st.session_state.offer_2_well_inspection_value = 'Y'
            st.session_state.offer_2_well_inspection_days = st.session_state.offer_2_well_inspection_days
            st.session_state.offer_2_well_inspection_days_string = st.session_state.offer_2_well_inspection_days
        else:
            st.session_state.offer_2_well_inspection_value = ''
            st.session_state.offer_2_well_inspection_days = 0
            st.session_state.offer_2_well_inspection_days_string = days_int_to_string(st.session_state.offer_2_well_inspection_days)
        if st.session_state.offer_2_financing_contingency_check:
            st.session_state.offer_2_financing_contingency_value = 'Y'
            st.session_state.offer_2_financing_contingency_days = st.session_state.offer_2_financing_contingency_days
            st.session_state.offer_2_financing_contingency_days_string = st.session_state.offer_2_financing_contingency_days
        else:
            st.session_state.offer_2_financing_contingency_value = ''
            st.session_state.offer_2_financing_contingency_days = 0
            st.session_state.offer_2_financing_contingency_days_string = days_int_to_string(st.session_state.offer_2_financing_contingency_days)
        if st.session_state.offer_2_appraisal_contingency_check:
            st.session_state.offer_2_appraisal_contingency_value = 'Y'
            st.session_state.offer_2_appraisal_contingency_days = st.session_state.offer_2_appraisal_contingency_days
            st.session_state.offer_2_appraisal_contingency_days_string = st.session_state.offer_2_appraisal_contingency_days
        else:
            st.session_state.offer_2_appraisal_contingency_value = ''
            st.session_state.offer_2_appraisal_contingency_days = 0
            st.session_state.offer_2_appraisal_contingency_days_string = days_int_to_string(st.session_state.offer_2_appraisal_contingency_days)
        if st.session_state.offer_2_home_sale_contingency_check:
            st.session_state.offer_2_home_sale_contingency_value = 'Y'
            st.session_state.offer_2_home_sale_contingency_days = st.session_state.offer_2_home_inspection_days
            st.session_state.offer_2_home_sale_contingency_days_string = st.session_state.offer_2_home_sale_contingency_days
        else:
            st.session_state.offer_2_home_sale_contingency_value = ''
            st.session_state.offer_2_home_sale_contingency_days = 0
            st.session_state.offer_2_home_sale_contingency_days_string = days_int_to_string(st.session_state.offer_2_home_sale_contingency_days)
        if st.session_state.offer_2_pre_occupancy_request:
            st.session_state.offer_2_pre_occupancy_date = st.session_state.offer_2_update_pre_occupancy_date
        else:
            st.session_state.offer_2_pre_occupancy_date = ''
        if st.session_state.offer_2_post_occupancy_request:
            st.session_state.offer_2_post_occupancy_date = st.session_state.offer_2_update_post_occupancy_date
        else:
            st.session_state.offer_2_post_occupancy_date = ''

    def update_offer_3_info_form():
        st.session_state.offer_3_name = st.session_state.update_offer_3_name
        st.session_state.offer_3_settlement_date = st.session_state.update_offer_3_settlement_date
        st.session_state.offer_3_settlement_company = st.session_state.update_offer_3_settlement_company
        st.session_state.offer_3_amt = st.session_state.update_offer_3_amt
        st.session_state.offer_3_emd_amt = st.session_state.update_offer_3_emd_amt
        st.session_state.offer_3_down_pmt_pct = st.session_state.update_offer_3_down_pmt_pct
        st.session_state.offer_3_down_pmt_pct = st.session_state.offer_3_down_pmt_pct / 100
        st.session_state.offer_3_closing_subsidy_pct = st.session_state.offer_3_update_closing_subsidy_pct / 100
        if st.session_state.offer_3_closing_subsidy_radio == 'Percent of Offer Amt (%)':
            st.session_state.offer_3_closing_subsidy_amt = st.session_state.offer_3_closing_subsidy_pct * st.session_state.offer_3_amt
        else:
            st.session_state.offer_3_closing_subsidy_amt = st.session_state.offer_3_closing_subsidy_flat_amt

        if st.session_state.offer_3_home_inspection_check:
            st.session_state.offer_3_home_inspection_value = 'Y'
            st.session_state.offer_3_home_inspection_days = st.session_state.offer_3_home_inspection_days
            st.session_state.offer_3_home_inspection_days_string = st.session_state.offer_3_home_inspection_days
        else:
            st.session_state.offer_3_home_inspection_value = ''
            st.session_state.offer_3_home_inspection_days = 0
            st.session_state.offer_3_home_insptection_days_string = days_int_to_string(st.session_state.offer_3_home_inspection_days)

        if st.session_state.offer_3_radon_inspection_check:
            st.session_state.offer_3_radon_inspection_value = 'Y'
            st.session_state.offer_3_radon_inspection_days = st.session_state.offer_3_radon_inspection_days
            st.session_state.offer_3_radon_inspection_days_string =st.session_state.offer_3_radon_inspection_days
        else:
            st.session_state.offer_3_radon_inspection_value = ''
            st.session_state.offer_3_radon_inspection_days = 0
            st.session_state.offer_3_radon_inspection_days_string = days_int_to_string(st.session_state.offer_3_radon_inspection_days)

        if st.session_state.offer_3_septic_inspection_check:
            st.session_state.offer_3_septic_inspection_value = 'Y'
            st.session_state.offer_3_septic_inspection_days = st.session_state.offer_3_septic_inspection_days
            st.session_state.offer_3_septic_inspection_days_string = st.session_state.offer_3_septic_inspection_days
        else:
            st.session_state.offer_3_septic_inspection_value = ''
            st.session_state.offer_3_septic_inspection_days = 0
            st.session_state.offer_3_septic_inspection_days_string = days_int_to_string(st.session_state.offer_3_septic_inspection_days)

        if st.session_state.offer_3_well_inspection_check:
            st.session_state.offer_3_well_inspection_value = 'Y'
            st.session_state.offer_3_well_inspection_days = st.session_state.offer_3_well_inspection_days
            st.session_state.offer_3_well_inspection_days_string = st.session_state.offer_3_well_inspection_days
        else:
            st.session_state.offer_3_well_inspection_value = ''
            st.session_state.offer_3_well_inspection_days = 0
            st.session_state.offer_3_well_inspection_days_string = days_int_to_string(st.session_state.offer_3_well_inspection_days)

        if st.session_state.offer_3_financing_contingency_check:
            st.session_state.offer_3_financing_contingency_value = 'Y'
            st.session_state.offer_3_financing_contingency_days = st.session_state.offer_3_financing_contingency_days
            st.session_state.offer_3_financing_contingency_days_string = st.session_state.offer_3_financing_contingency_days
        else:
            st.session_state.offer_3_financing_contingency_value = ''
            st.session_state.offer_3_financing_contingency_days = 0
            st.session_state.offer_3_financing_contingency_days_string = days_int_to_string(st.session_state.offer_3_financing_contingency_days)

        if st.session_state.offer_3_appraisal_contingency_check:
            st.session_state.offer_3_appraisal_contingency_value = 'Y'
            st.session_state.offer_3_appraisal_contingency_days = st.session_state.offer_3_appraisal_contingency_days
            st.session_state.offer_3_appraisal_contingency_days_string = st.session_state.offer_3_appraisal_contingency_days
        else:
            st.session_state.offer_3_appraisal_contingency_value = ''
            st.session_state.offer_3_appraisal_contingency_days = 0
            st.session_state.offer_3_appraisal_contingency_days_string = days_int_to_string(st.session_state.offer_3_appraisal_contingency_days)

        if st.session_state.offer_3_home_sale_contingency_check:
            st.session_state.offer_3_home_sale_contingency_value = "Y"
            st.session_state.offer_3_home_sale_contingency_days = st.session_state.offer_3_home_inspection_days
            st.session_state.offer_3_home_sale_contingency_days_string = st.session_state.offer_3_home_sale_contingency_days
        else:
            st.session_state.offer_3_home_sale_contingency_value = ''
            st.session_state.offer_3_home_sale_contingency_days = 0
            st.session_state.offer_3_home_sale_contingency_days_string = days_int_to_string(st.session_state.offer_3_home_sale_contingency_days)

        if st.session_state.offer_3_pre_occupancy_request:
            st.session_state.offer_3_pre_occupancy_date = st.session_state.offer_3_update_pre_occupancy_date
        else:
            st.session_state.offer_3_pre_occupancy_date = ''

        if st.session_state.offer_3_post_occupancy_request:
            st.session_state.offer_3_post_occupancy_date = st.session_state.offer_3_update_post_occupancy_date
        else:
            st.session_state.offer_3_post_occupancy_date = ''

    def update_offer_4_info_form():
        st.session_state.offer_4_name = st.session_state.update_offer_4_name
        st.session_state.offer_4_settlement_date = st.session_state.update_offer_4_settlement_date
        st.session_state.offer_4_settlement_company = st.session_state.update_offer_4_settlement_company
        st.session_state.offer_4_amt = st.session_state.update_offer_4_amt
        st.session_state.offer_4_emd_amt = st.session_state.update_offer_4_emd_amt
        st.session_state.offer_4_down_pmt_pct = st.session_state.update_offer_4_down_pmt_pct
        st.session_state.offer_4_down_pmt_pct = st.session_state.offer_4_down_pmt_pct / 100
        st.session_state.offer_4_closing_subsidy_pct = st.session_state.offer_4_update_closing_subsidy_pct / 100
        if st.session_state.offer_4_closing_subsidy_radio == 'Percent of Offer Amt (%)':
            st.session_state.offer_4_closing_subsidy_amt = st.session_state.offer_4_closing_subsidy_pct * st.session_state.offer_4_amt
        else:
            st.session_state.offer_4_closing_subsidy_amt = st.session_state.offer_4_closing_subsidy_flat_amt

        if st.session_state.offer_4_home_inspection_check:
            st.session_state.offer_4_home_inspection_value = 'Y'
            st.session_state.offer_4_home_inspection_days = st.session_state.offer_4_home_inspection_days
            st.session_state.offer_4_home_inspection_days_string = st.session_state.offer_4_home_inspection_days
        else:
            st.session_state.offer_4_home_inspection_value = ''
            st.session_state.offer_4_home_inspection_days = 0
            st.session_state.offer_4_home_inspection_days_string = days_int_to_string(st.session_state.offer_4_home_inspection_days)

        if st.session_state.offer_4_radon_inspection_check:
            st.session_state.offer_4_radon_inspection_value = 'Y'
            st.session_state.offer_4_radon_inspection_days = st.session_state.offer_4_radon_inspection_days
            st.session_state.offer_4_radon_inspection_days_string = st.session_state.offer_4_radon_inspection_days
        else:
            st.session_state.offer_4_radon_inspection_value = ''
            st.session_state.offer_4_radon_inspection_days = 0
            st.session_state.offer_4_radon_inspection_days_string = days_int_to_string(st.session_state.offer_4_radon_inspection_days)

        if st.session_state.offer_4_septic_inspection_check:
            st.session_state.offer_4_septic_inspection_value = 'Y'
            st.session_state.offer_4_septic_inspection_days = st.session_state.offer_4_septic_inspection_days
            st.session_state.offer_4_septic_inspection_days_string = st.session_state.offer_4_septic_inspection_days
        else:
            st.session_state.offer_4_septic_inspection_value = ''
            st.session_state.offer_4_septic_inspection_days = 0
            st.session_state.offer_4_septic_inspection_days_string = days_int_to_string(st.session_state.offer_4_septic_inspection_days)

        if st.session_state.offer_4_well_inspection_check:
            st.session_state.offer_4_well_inspection_value = 'Y'
            st.session_state.offer_4_well_inspection_days = st.session_state.offer_4_well_inspection_days
            st.session_state.offer_4_well_inspection_days_string = st.session_state.offer_4_well_inspection_days
        else:
            st.session_state.offer_4_well_inspection_value = ''
            st.session_state.offer_4_well_inspection_days = 0
            st.session_state.offer_4_well_inspection_days_string = days_int_to_string(st.session_state.offer_4_well_inspection_days)

        if st.session_state.offer_4_financing_contingency_check:
            st.session_state.offer_4_financing_contingency_value = 'Y'
            st.session_state.offer_4_financing_contingency_days = st.session_state.offer_4_financing_contingency_days
            st.session_state.offer_4_financing_contingency_days_string = st.session_state.offer_4_financing_contingency_days
        else:
            st.session_state.offer_4_financing_contingency_value = ''
            st.session_state.offer_4_financing_contingency_days = 0
            st.session_state.offer_4_financing_contingency_days_string = days_int_to_string(st.session_state.offer_4_financing_contingency_days)

        if st.session_state.offer_4_appraisal_contingency_check:
            st.session_state.offer_4_appraisal_contingency_value = 'Y'
            st.session_state.offer_4_appraisal_contingency_days = st.session_state.offer_4_appraisal_contingency_days
            st.session_state.offer_4_appraisal_contingency_days_string = st.session_state.offer_4_appraisal_contingency_days
        else:
            st.session_state.offer_4_appraisal_contingency_value = ''
            st.session_state.offer_4_appraisal_contingency_days = 0
            st.session_state.offer_4_appraisal_contingency_days_string = days_int_to_string(st.session_state.offer_4_appraisal_contingency_days)

        if st.session_state.offer_4_home_sale_contingency_check:
            st.session_state.offer_4_home_sale_contingency_value = 'Y'
            st.session_state.offer_4_home_sale_contingency_days = st.session_state.offer_4_home_inspection_days
            st.session_state.offer_4_home_sale_contingency_days_string = st.session_state.offer_4_home_sale_contingency_days
        else:
            st.session_state.offer_4_home_sale_contingency_value = ''
            st.session_state.offer_4_home_sale_contingency_days = 0
            st.session_state.offer_4_home_sale_contingency_days_string = days_int_to_string(
                st.session_state.offer_4_home_sale_contingency_days)

        if st.session_state.offer_4_pre_occupancy_request:
            st.session_state.offer_4_pre_occupancy_date = st.session_state.offer_4_update_pre_occupancy_date
        else:
            st.session_state.offer_4_pre_occupancy_date = ''

        if st.session_state.offer_4_post_occupancy_request:
            st.session_state.offer_4_post_occupancy_date = st.session_state.offer_4_update_post_occupancy_date
        else:
            st.session_state.offer_4_post_occupancy_date = ''

    def update_offer_5_info_form():
        st.session_state.offer_5_name = st.session_state.update_offer_5_name
        st.session_state.offer_5_settlement_date = st.session_state.update_offer_5_settlement_date
        st.session_state.offer_5_settlement_company = st.session_state.update_offer_5_settlement_company
        st.session_state.offer_5_amt = st.session_state.update_offer_5_amt
        st.session_state.offer_5_emd_amt = st.session_state.update_offer_5_emd_amt
        st.session_state.offer_5_down_pmt_pct = st.session_state.update_offer_5_down_pmt_pct
        st.session_state.offer_5_down_pmt_pct = st.session_state.offer_5_down_pmt_pct / 100
        st.session_state.offer_5_closing_subsidy_pct = st.session_state.offer_5_update_closing_subsidy_pct / 100
        if st.session_state.offer_5_closing_subsidy_radio == 'Percent of Offer Amt (%)':
            st.session_state.offer_5_closing_subsidy_amt = st.session_state.offer_5_closing_subsidy_pct * st.session_state.offer_5_amt
        else:
            st.session_state.offer_5_closing_subsidy_amt = st.session_state.offer_5_closing_subsidy_flat_amt

        if st.session_state.offer_5_home_inspection_check:
            st.session_state.offer_5_home_inspection_value = 'Y'
            st.session_state.offer_5_home_inspection_days = st.session_state.offer_5_home_inspection_days
            st.session_state.offer_5_home_inspection_days_string = st.session_state.offer_5_home_inspection_days
        else:
            st.session_state.offer_5_home_inspection_value = ''
            st.session_state.offer_5_home_inspection_days = 0
            st.session_state.offer_5_home_inspection_days_string = days_int_to_string(st.session_state.offer_5_home_inspection_days)

        if st.session_state.offer_5_radon_inspection_check:
            st.session_state.offer_5_radon_inspection_value = 'Y'
            st.session_state.offer_5_radon_inspection_days = st.session_state.offer_5_radon_inspection_days
            st.session_state.offer_5_radon_inspection_days_string = st.session_state.offer_5_radon_inspection_days
        else:
            st.session_state.offer_5_radon_inspection_value = ''
            st.session_state.offer_5_radon_inspection_days = 0
            st.session_state.offer_5_radon_inspection_days_string = days_int_to_string(st.session_state.offer_5_radon_inspection_days)

        if st.session_state.offer_5_septic_inspection_check:
            st.session_state.offer_5_septic_inspection_value = 'Y'
            st.session_state.offer_5_septic_inspection_days = st.session_state.offer_5_septic_inspection_days
            st.session_state.offer_5_septic_inspection_days_string = st.session_state.offer_5_septic_inspection_days
        else:
            st.session_state.offer_5_septic_inspection_value = ''
            st.session_state.offer_5_septic_inspection_days = 0
            st.session_state.offer_5_septic_inspection_days_string = days_int_to_string(st.session_state.offer_5_septic_inspection_days)

        if st.session_state.offer_5_well_inspection_check:
            st.session_state.offer_5_well_inspection_value = 'Y'
            st.session_state.offer_5_well_inspection_days = st.session_state.offer_5_well_inspection_days
            st.session_state.offer_5_well_inspection_days_string = st.session_state.offer_5_well_inspection_days
        else:
            st.session_state.offer_5_well_inspection_value = ''
            st.session_state.offer_5_well_inspection_days = 0
            st.session_state.offer_5_well_inspection_days_string = days_int_to_string(st.session_state.offer_5_well_inspection_days)

        if st.session_state.offer_5_financing_contingency_check:
            st.session_state.offer_5_financing_contingency_value = 'Y'
            st.session_state.offer_5_financing_contingency_days = st.session_state.offer_5_financing_contingency_days
            st.session_state.offer_5_financing_contingency_days_string = st.session_state.offer_5_financing_contingency_days
        else:
            st.session_state.offer_5_financing_contingency_value = ''
            st.session_state.offer_5_financing_contingency_days = 0
            st.session_state.offer_5_financing_contingency_days_string = days_int_to_string(st.session_state.offer_5_financing_contingency_days)

        if st.session_state.offer_5_appraisal_contingency_check:
            st.session_state.offer_5_appraisal_contingency_value = 'Y'
            st.session_state.offer_5_appraisal_contingency_days = st.session_state.offer_5_appraisal_contingency_days
            st.session_state.offer_5_appraisal_contingency_days_string = st.session_state.offer_5_appraisal_contingency_days
        else:
            st.session_state.offer_5_appraisal_contingency_value = ''
            st.session_state.offer_5_appraisal_contingency_days = 0
            st.session_state.offer_5_appraisal_contingency_days_string = days_int_to_string(st.session_state.offer_5_appraisal_contingency_days)

        if st.session_state.offer_5_home_sale_contingency_check:
            st.session_state.offer_5_home_sale_contingency_value = 'Y'
            st.session_state.offer_5_home_sale_contingency_days = st.session_state.offer_5_home_inspection_days
            st.session_state.offer_5_home_sale_contingency_days_string = st.session_state.offer_5_home_sale_contingency_days
        else:
            st.session_state.offer_5_home_sale_contingency_value = ''
            st.session_state.offer_5_home_sale_contingency_days = 0
            st.session_state.offer_5_home_sale_contingency_days_string = days_int_to_string(
                st.session_state.offer_5_home_sale_contingency_days)

        if st.session_state.offer_5_pre_occupancy_request:
            st.session_state.offer_5_pre_occupancy_date = st.session_state.offer_5_update_pre_occupancy_date
        else:
            st.session_state.offer_5_pre_occupancy_date = ''

        if st.session_state.offer_5_post_occupancy_request:
            st.session_state.offer_5_post_occupancy_date = st.session_state.offer_5_update_post_occupancy_date
        else:
            st.session_state.offer_5_post_occupancy_date = ''
            
    def update_offer_6_info_form():
        st.session_state.offer_6_name = st.session_state.update_offer_6_name
        st.session_state.offer_6_settlement_date = st.session_state.update_offer_6_settlement_date
        st.session_state.offer_6_settlement_company = st.session_state.update_offer_6_settlement_company
        st.session_state.offer_6_amt = st.session_state.update_offer_6_amt
        st.session_state.offer_6_emd_amt = st.session_state.update_offer_6_emd_amt
        st.session_state.offer_6_down_pmt_pct = st.session_state.update_offer_6_down_pmt_pct
        st.session_state.offer_6_down_pmt_pct = st.session_state.offer_6_down_pmt_pct / 100
        st.session_state.offer_6_closing_subsidy_pct = st.session_state.offer_6_update_closing_subsidy_pct / 100
        if st.session_state.offer_6_closing_subsidy_radio == 'Percent of Offer Amt (%)':
            st.session_state.offer_6_closing_subsidy_amt = st.session_state.offer_6_closing_subsidy_pct * st.session_state.offer_6_amt
        else:
            st.session_state.offer_6_closing_subsidy_amt = st.session_state.offer_6_closing_subsidy_flat_amt

        if st.session_state.offer_6_home_inspection_check:
            st.session_state.offer_6_home_inspection_value = 'Y'
            st.session_state.offer_6_home_inspection_days = st.session_state.offer_6_home_inspection_days
            st.session_state.offer_6_home_inspection_days_string = st.session_state.offer_6_home_inspection_days
        else:
            st.session_state.offer_6_home_inspection_value = ''
            st.session_state.offer_6_home_inspection_days = 0
            st.session_state.offer_6_home_inspection_days_string = days_int_to_string(st.session_state.offer_6_home_inspection_days)

        if st.session_state.offer_6_radon_inspection_check:
            st.session_state.offer_6_radon_inspection_value = 'Y'
            st.session_state.offer_6_radon_inspection_days = st.session_state.offer_6_radon_inspection_days
            st.session_state.offer_6_radon_inspection_days_string = st.session_state.offer_6_radon_inspection_days
        else:
            st.session_state.offer_6_radon_inspection_value = ''
            st.session_state.offer_6_radon_inspection_days = 0
            st.session_state.offer_6_radon_inspection_days_string = days_int_to_string(st.session_state.offer_6_radon_inspection_days)

        if st.session_state.offer_6_septic_inspection_check:
            st.session_state.offer_6_septic_inspection_value = 'Y'
            st.session_state.offer_6_septic_inspection_days = st.session_state.offer_6_septic_inspection_days
            st.session_state.offer_6_septic_inspection_days_string = st.session_state.offer_6_septic_inspection_days
        else:
            st.session_state.offer_6_septic_inspection_value = ''
            st.session_state.offer_6_septic_inspection_days = 0
            st.session_state.offer_6_septic_inspection_days_string = days_int_to_string(st.session_state.offer_6_septic_inspection_days)

        if st.session_state.offer_6_well_inspection_check:
            st.session_state.offer_6_well_inspection_value = 'Y'
            st.session_state.offer_6_well_inspection_days = st.session_state.offer_6_well_inspection_days
            st.session_state.offer_6_well_inspection_days_string = st.session_state.offer_6_well_inspection_days
        else:
            st.session_state.offer_6_well_inspection_value = ''
            st.session_state.offer_6_well_inspection_days = 0
            st.session_state.offer_6_well_inspection_days_string = days_int_to_string(st.session_state.offer_6_well_inspection_days)

        if st.session_state.offer_6_financing_contingency_check:
            st.session_state.offer_6_financing_contingency_value = 'Y'
            st.session_state.offer_6_financing_contingency_days = st.session_state.offer_6_financing_contingency_days
            st.session_state.offer_6_financing_contingency_days_string = st.session_state.offer_6_financing_contingency_days
        else:
            st.session_state.offer_6_financing_contingency_value = ''
            st.session_state.offer_6_financing_contingency_days = 0
            st.session_state.offer_6_financing_contingency_days_string = days_int_to_string(st.session_state.offer_6_financing_contingency_days)

        if st.session_state.offer_6_appraisal_contingency_check:
            st.session_state.offer_6_appraisal_contingency_value = 'Y'
            st.session_state.offer_6_appraisal_contingency_days = st.session_state.offer_6_appraisal_contingency_days
            st.session_state.offer_6_appraisal_contingency_days_string = st.session_state.offer_6_appraisal_contingency_days
        else:
            st.session_state.offer_6_appraisal_contingency_value = ''
            st.session_state.offer_6_appraisal_contingency_days = 0
            st.session_state.offer_6_appraisal_contingency_days_string = days_int_to_string(st.session_state.offer_6_appraisal_contingency_days)

        if st.session_state.offer_6_home_sale_contingency_check:
            st.session_state.offer_6_home_sale_contingency_value = 'Y'
            st.session_state.offer_6_home_sale_contingency_days = st.session_state.offer_6_home_inspection_days
            st.session_state.offer_6_home_sale_contingency_days_string = st.session_state.offer_6_home_sale_contingency_days
        else:
            st.session_state.offer_6_home_sale_contingency_value = ''
            st.session_state.offer_6_home_sale_contingency_days = 0
            st.session_state.offer_6_home_sale_contingency_days_string = days_int_to_string(
                st.session_state.offer_6_home_sale_contingency_days)

        if st.session_state.offer_6_pre_occupancy_request:
            st.session_state.offer_6_pre_occupancy_date = st.session_state.offer_6_update_pre_occupancy_date
        else:
            st.session_state.offer_6_pre_occupancy_date = ''

        if st.session_state.offer_6_post_occupancy_request:
            st.session_state.offer_6_post_occupancy_date = st.session_state.offer_6_update_post_occupancy_date
        else:
            st.session_state.offer_6_post_occupancy_date = ''
            
    def update_offer_7_info_form():
        st.session_state.offer_7_name = st.session_state.update_offer_7_name
        st.session_state.offer_7_settlement_date = st.session_state.update_offer_7_settlement_date
        st.session_state.offer_7_settlement_company = st.session_state.update_offer_7_settlement_company
        st.session_state.offer_7_amt = st.session_state.update_offer_7_amt
        st.session_state.offer_7_emd_amt = st.session_state.update_offer_7_emd_amt
        st.session_state.offer_7_down_pmt_pct = st.session_state.update_offer_7_down_pmt_pct
        st.session_state.offer_7_down_pmt_pct = st.session_state.offer_7_down_pmt_pct / 100
        st.session_state.offer_7_closing_subsidy_pct = st.session_state.offer_7_update_closing_subsidy_pct / 100
        if st.session_state.offer_7_closing_subsidy_radio == 'Percent of Offer Amt (%)':
            st.session_state.offer_7_closing_subsidy_amt = st.session_state.offer_7_closing_subsidy_pct * st.session_state.offer_7_amt
        else:
            st.session_state.offer_7_closing_subsidy_amt = st.session_state.offer_7_closing_subsidy_flat_amt

        if st.session_state.offer_7_home_inspection_check:
            st.session_state.offer_7_home_inspection_value = 'Y'
            st.session_state.offer_7_home_inspection_days = st.session_state.offer_7_home_inspection_days
            st.session_state.offer_7_home_inspection_days_string = st.session_state.offer_7_home_inspection_days
        else:
            st.session_state.offer_7_home_inspection_value = ''
            st.session_state.offer_7_home_inspection_days = 0
            st.session_state.offer_7_home_inspection_days_string = days_int_to_string(st.session_state.offer_7_home_inspection_days)

        if st.session_state.offer_7_radon_inspection_check:
            st.session_state.offer_7_radon_inspection_value = 'Y'
            st.session_state.offer_7_radon_inspection_days = st.session_state.offer_7_radon_inspection_days
            st.session_state.offer_7_radon_inspection_days_string = st.session_state.offer_7_radon_inspection_days
        else:
            st.session_state.offer_7_radon_inspection_value = ''
            st.session_state.offer_7_radon_inspection_days = 0
            st.session_state.offer_7_radon_inspection_days_string = days_int_to_string(st.session_state.offer_7_radon_inspection_days)

        if st.session_state.offer_7_septic_inspection_check:
            st.session_state.offer_7_septic_inspection_value = 'Y'
            st.session_state.offer_7_septic_inspection_days = st.session_state.offer_7_septic_inspection_days
            st.session_state.offer_7_septic_inspection_days_string = st.session_state.offer_7_septic_inspection_days
        else:
            st.session_state.offer_7_septic_inspection_value = ''
            st.session_state.offer_7_septic_inspection_days = 0
            st.session_state.offer_7_septic_inspection_days_string = days_int_to_string(st.session_state.offer_7_septic_inspection_days)

        if st.session_state.offer_7_well_inspection_check:
            st.session_state.offer_7_well_inspection_value = 'Y'
            st.session_state.offer_7_well_inspection_days = st.session_state.offer_7_well_inspection_days
            st.session_state.offer_7_well_inspection_days_string = st.session_state.offer_7_well_inspection_days
        else:
            st.session_state.offer_7_well_inspection_value = ''
            st.session_state.offer_7_well_inspection_days = 0
            st.session_state.offer_7_well_inspection_days_string = days_int_to_string(st.session_state.offer_7_well_inspection_days)

        if st.session_state.offer_7_financing_contingency_check:
            st.session_state.offer_7_financing_contingency_value = 'Y'
            st.session_state.offer_7_financing_contingency_days = st.session_state.offer_7_financing_contingency_days
            st.session_state.offer_7_financing_contingency_days_string = st.session_state.offer_7_financing_contingency_days
        else:
            st.session_state.offer_7_financing_contingency_value = ''
            st.session_state.offer_7_financing_contingency_days = 0
            st.session_state.offer_7_financing_contingency_days_string = days_int_to_string(st.session_state.offer_7_financing_contingency_days)

        if st.session_state.offer_7_appraisal_contingency_check:
            st.session_state.offer_7_appraisal_contingency_value = 'Y'
            st.session_state.offer_7_appraisal_contingency_days = st.session_state.offer_7_appraisal_contingency_days
            st.session_state.offer_7_appraisal_contingency_days_string = st.session_state.offer_7_appraisal_contingency_days
        else:
            st.session_state.offer_7_appraisal_contingency_value = ''
            st.session_state.offer_7_appraisal_contingency_days = 0
            st.session_state.offer_7_appraisal_contingency_days_string = days_int_to_string(st.session_state.offer_7_appraisal_contingency_days)

        if st.session_state.offer_7_home_sale_contingency_check:
            st.session_state.offer_7_home_sale_contingency_value = 'Y'
            st.session_state.offer_7_home_sale_contingency_days = st.session_state.offer_7_home_inspection_days
            st.session_state.offer_7_home_sale_contingency_days_string = st.session_state.offer_7_home_sale_contingency_days
        else:
            st.session_state.offer_7_home_sale_contingency_value = ''
            st.session_state.offer_7_home_sale_contingency_days = 0
            st.session_state.offer_7_home_sale_contingency_days_string = days_int_to_string(
                st.session_state.offer_7_home_sale_contingency_days)

        if st.session_state.offer_7_pre_occupancy_request:
            st.session_state.offer_7_pre_occupancy_date = st.session_state.offer_7_update_pre_occupancy_date
        else:
            st.session_state.offer_7_pre_occupancy_date = ''

        if st.session_state.offer_7_post_occupancy_request:
            st.session_state.offer_7_post_occupancy_date = st.session_state.offer_7_update_post_occupancy_date
        else:
            st.session_state.offer_7_post_occupancy_date = ''
            
    def update_offer_8_info_form():
        st.session_state.offer_8_name = st.session_state.update_offer_8_name
        st.session_state.offer_8_settlement_date = st.session_state.update_offer_8_settlement_date
        st.session_state.offer_8_settlement_company = st.session_state.update_offer_8_settlement_company
        st.session_state.offer_8_amt = st.session_state.update_offer_8_amt
        st.session_state.offer_8_emd_amt = st.session_state.update_offer_8_emd_amt
        st.session_state.offer_8_down_pmt_pct = st.session_state.update_offer_8_down_pmt_pct
        st.session_state.offer_8_down_pmt_pct = st.session_state.offer_8_down_pmt_pct / 100
        st.session_state.offer_8_closing_subsidy_pct = st.session_state.offer_8_update_closing_subsidy_pct / 100
        if st.session_state.offer_8_closing_subsidy_radio == 'Percent of Offer Amt (%)':
            st.session_state.offer_8_closing_subsidy_amt = st.session_state.offer_8_closing_subsidy_pct * st.session_state.offer_8_amt
        else:
            st.session_state.offer_8_closing_subsidy_amt = st.session_state.offer_8_closing_subsidy_flat_amt

        if st.session_state.offer_8_home_inspection_check:
            st.session_state.offer_8_home_inspection_value = 'Y'
            st.session_state.offer_8_home_inspection_days = st.session_state.offer_8_home_inspection_days
            st.session_state.offer_8_home_inspection_days_string = st.session_state.offer_8_home_inspection_days
        else:
            st.session_state.offer_8_home_inspection_value = ''
            st.session_state.offer_8_home_inspection_days = 0
            st.session_state.offer_8_home_inspection_days_string = days_int_to_string(st.session_state.offer_8_home_inspection_days)

        if st.session_state.offer_8_radon_inspection_check:
            st.session_state.offer_8_radon_inspection_value = 'Y'
            st.session_state.offer_8_radon_inspection_days = st.session_state.offer_8_radon_inspection_days
            st.session_state.offer_8_radon_inspection_days_string = st.session_state.offer_8_radon_inspection_days
        else:
            st.session_state.offer_8_radon_inspection_value = ''
            st.session_state.offer_8_radon_inspection_days = 0
            st.session_state.offer_8_radon_inspection_days_string = days_int_to_string(st.session_state.offer_8_radon_inspection_days)

        if st.session_state.offer_8_septic_inspection_check:
            st.session_state.offer_8_septic_inspection_value = 'Y'
            st.session_state.offer_8_septic_inspection_days = st.session_state.offer_8_septic_inspection_days
            st.session_state.offer_8_septic_inspection_days_string = st.session_state.offer_8_septic_inspection_days
        else:
            st.session_state.offer_8_septic_inspection_value = ''
            st.session_state.offer_8_septic_inspection_days = 0
            st.session_state.offer_8_septic_inspection_days_string = days_int_to_string(st.session_state.offer_8_septic_inspection_days)

        if st.session_state.offer_8_well_inspection_check:
            st.session_state.offer_8_well_inspection_value = 'Y'
            st.session_state.offer_8_well_inspection_days = st.session_state.offer_8_well_inspection_days
            st.session_state.offer_8_well_inspection_days_string = st.session_state.offer_8_well_inspection_days
        else:
            st.session_state.offer_8_well_inspection_value = ''
            st.session_state.offer_8_well_inspection_days = 0
            st.session_state.offer_8_well_inspection_days_string = days_int_to_string(st.session_state.offer_8_well_inspection_days)

        if st.session_state.offer_8_financing_contingency_check:
            st.session_state.offer_8_financing_contingency_value = 'Y'
            st.session_state.offer_8_financing_contingency_days = st.session_state.offer_8_financing_contingency_days
            st.session_state.offer_8_financing_contingency_days_string = st.session_state.offer_8_financing_contingency_days
        else:
            st.session_state.offer_8_financing_contingency_value = ''
            st.session_state.offer_8_financing_contingency_days = 0
            st.session_state.offer_8_financing_contingency_days_string = days_int_to_string(st.session_state.offer_8_financing_contingency_days)

        if st.session_state.offer_8_appraisal_contingency_check:
            st.session_state.offer_8_appraisal_contingency_value = 'Y'
            st.session_state.offer_8_appraisal_contingency_days = st.session_state.offer_8_appraisal_contingency_days
            st.session_state.offer_8_appraisal_contingency_days_string = st.session_state.offer_8_appraisal_contingency_days
        else:
            st.session_state.offer_8_appraisal_contingency_value = ''
            st.session_state.offer_8_appraisal_contingency_days = 0
            st.session_state.offer_8_appraisal_contingency_days_string = days_int_to_string(st.session_state.offer_8_appraisal_contingency_days)

        if st.session_state.offer_8_home_sale_contingency_check:
            st.session_state.offer_8_home_sale_contingency_value = 'Y'
            st.session_state.offer_8_home_sale_contingency_days = st.session_state.offer_8_home_inspection_days
            st.session_state.offer_8_home_sale_contingency_days_string = st.session_state.offer_8_home_sale_contingency_days
        else:
            st.session_state.offer_8_home_sale_contingency_value = ''
            st.session_state.offer_8_home_sale_contingency_days = 0
            st.session_state.offer_8_home_sale_contingency_days_string = days_int_to_string(
                st.session_state.offer_8_home_sale_contingency_days)

        if st.session_state.offer_8_pre_occupancy_request:
            st.session_state.offer_8_pre_occupancy_date = st.session_state.offer_8_update_pre_occupancy_date
        else:
            st.session_state.offer_8_pre_occupancy_date = ''

        if st.session_state.offer_8_post_occupancy_request:
            st.session_state.offer_8_post_occupancy_date = st.session_state.offer_8_update_post_occupancy_date
        else:
            st.session_state.offer_8_post_occupancy_date = ''
            
    def update_offer_9_info_form():
        st.session_state.offer_9_name = st.session_state.update_offer_9_name
        st.session_state.offer_9_settlement_date = st.session_state.update_offer_9_settlement_date
        st.session_state.offer_9_settlement_company = st.session_state.update_offer_9_settlement_company
        st.session_state.offer_9_amt = st.session_state.update_offer_9_amt
        st.session_state.offer_9_emd_amt = st.session_state.update_offer_9_emd_amt
        st.session_state.offer_9_down_pmt_pct = st.session_state.update_offer_9_down_pmt_pct
        st.session_state.offer_9_down_pmt_pct = st.session_state.offer_9_down_pmt_pct / 100
        st.session_state.offer_9_closing_subsidy_pct = st.session_state.offer_9_update_closing_subsidy_pct / 100
        if st.session_state.offer_9_closing_subsidy_radio == 'Percent of Offer Amt (%)':
            st.session_state.offer_9_closing_subsidy_amt = st.session_state.offer_9_closing_subsidy_pct * st.session_state.offer_9_amt
        else:
            st.session_state.offer_9_closing_subsidy_amt = st.session_state.offer_9_closing_subsidy_flat_amt

        if st.session_state.offer_9_home_inspection_check:
            st.session_state.offer_9_home_inspection_value = 'Y'
            st.session_state.offer_9_home_inspection_days = st.session_state.offer_9_home_inspection_days
            st.session_state.offer_9_home_inspection_days_string = st.session_state.offer_9_home_inspection_days
        else:
            st.session_state.offer_9_home_inspection_value = ''
            st.session_state.offer_9_home_inspection_days = 0
            st.session_state.offer_9_home_inspection_days_string = days_int_to_string(st.session_state.offer_9_home_inspection_days)

        if st.session_state.offer_9_radon_inspection_check:
            st.session_state.offer_9_radon_inspection_value = 'Y'
            st.session_state.offer_9_radon_inspection_days = st.session_state.offer_9_radon_inspection_days
            st.session_state.offer_9_radon_inspection_days_string = st.session_state.offer_9_radon_inspection_days
        else:
            st.session_state.offer_9_radon_inspection_value = ''
            st.session_state.offer_9_radon_inspection_days = 0
            st.session_state.offer_9_radon_inspection_days_string = days_int_to_string(st.session_state.offer_9_radon_inspection_days)

        if st.session_state.offer_9_septic_inspection_check:
            st.session_state.offer_9_septic_inspection_value = 'Y'
            st.session_state.offer_9_septic_inspection_days = st.session_state.offer_9_septic_inspection_days
            st.session_state.offer_9_septic_inspection_days_string = st.session_state.offer_9_septic_inspection_days
        else:
            st.session_state.offer_9_septic_inspection_value = ''
            st.session_state.offer_9_septic_inspection_days = 0
            st.session_state.offer_9_septic_inspection_days_string = days_int_to_string(st.session_state.offer_9_septic_inspection_days)

        if st.session_state.offer_9_well_inspection_check:
            st.session_state.offer_9_well_inspection_value = 'Y'
            st.session_state.offer_9_well_inspection_days = st.session_state.offer_9_well_inspection_days
            st.session_state.offer_9_well_inspection_days_string = st.session_state.offer_9_well_inspection_days
        else:
            st.session_state.offer_9_well_inspection_value = ''
            st.session_state.offer_9_well_inspection_days = 0
            st.session_state.offer_9_well_inspection_days_string = days_int_to_string(st.session_state.offer_9_well_inspection_days)

        if st.session_state.offer_9_financing_contingency_check:
            st.session_state.offer_9_financing_contingency_value = 'Y'
            st.session_state.offer_9_financing_contingency_days = st.session_state.offer_9_financing_contingency_days
            st.session_state.offer_9_financing_contingency_days_string = st.session_state.offer_9_financing_contingency_days
        else:
            st.session_state.offer_9_financing_contingency_value = ''
            st.session_state.offer_9_financing_contingency_days = 0
            st.session_state.offer_9_financing_contingency_days_string = days_int_to_string(st.session_state.offer_9_financing_contingency_days)

        if st.session_state.offer_9_appraisal_contingency_check:
            st.session_state.offer_9_appraisal_contingency_value = 'Y'
            st.session_state.offer_9_appraisal_contingency_days = st.session_state.offer_9_appraisal_contingency_days
            st.session_state.offer_9_appraisal_contingency_days_string = st.session_state.offer_9_appraisal_contingency_days
        else:
            st.session_state.offer_9_appraisal_contingency_value = ''
            st.session_state.offer_9_appraisal_contingency_days = 0
            st.session_state.offer_9_appraisal_contingency_days_string = days_int_to_string(st.session_state.offer_9_appraisal_contingency_days)

        if st.session_state.offer_9_home_sale_contingency_check:
            st.session_state.offer_9_home_sale_contingency_value = 'Y'
            st.session_state.offer_9_home_sale_contingency_days = st.session_state.offer_9_home_inspection_days
            st.session_state.offer_9_home_sale_contingency_days_string = st.session_state.offer_9_home_sale_contingency_days
        else:
            st.session_state.offer_9_home_sale_contingency_value = ''
            st.session_state.offer_9_home_sale_contingency_days = 0
            st.session_state.offer_9_home_sale_contingency_days_string = days_int_to_string(
                st.session_state.offer_9_home_sale_contingency_days)

        if st.session_state.offer_9_pre_occupancy_request:
            st.session_state.offer_9_pre_occupancy_date = st.session_state.offer_9_update_pre_occupancy_date
        else:
            st.session_state.offer_9_pre_occupancy_date = ''

        if st.session_state.offer_9_post_occupancy_request:
            st.session_state.offer_9_post_occupancy_date = st.session_state.offer_9_update_post_occupancy_date
        else:
            st.session_state.offer_9_post_occupancy_date = ''
            
    def update_offer_10_info_form():
        st.session_state.offer_10_name = st.session_state.update_offer_10_name
        st.session_state.offer_10_settlement_date = st.session_state.update_offer_10_settlement_date
        st.session_state.offer_10_settlement_company = st.session_state.update_offer_10_settlement_company
        st.session_state.offer_10_amt = st.session_state.update_offer_10_amt
        st.session_state.offer_10_emd_amt = st.session_state.update_offer_10_emd_amt
        st.session_state.offer_10_down_pmt_pct = st.session_state.update_offer_10_down_pmt_pct
        st.session_state.offer_10_down_pmt_pct = st.session_state.offer_10_down_pmt_pct / 100
        st.session_state.offer_10_closing_subsidy_pct = st.session_state.offer_10_update_closing_subsidy_pct / 100
        if st.session_state.offer_10_closing_subsidy_radio == 'Percent of Offer Amt (%)':
            st.session_state.offer_10_closing_subsidy_amt = st.session_state.offer_10_closing_subsidy_pct * st.session_state.offer_10_amt
        else:
            st.session_state.offer_10_closing_subsidy_amt = st.session_state.offer_10_closing_subsidy_flat_amt

        if st.session_state.offer_10_home_inspection_check:
            st.session_state.offer_10_home_inspection_value = 'Y'
            st.session_state.offer_10_home_inspection_days = st.session_state.offer_10_home_inspection_days
            st.session_state.offer_10_home_inspection_days_string = st.session_state.offer_10_home_inspection_days
        else:
            st.session_state.offer_10_home_inspection_value = ''
            st.session_state.offer_10_home_inspection_days = 0
            st.session_state.offer_10_home_inspection_days_string = days_int_to_string(st.session_state.offer_10_home_inspection_days)

        if st.session_state.offer_10_radon_inspection_check:
            st.session_state.offer_10_radon_inspection_value = 'Y'
            st.session_state.offer_10_radon_inspection_days = st.session_state.offer_10_radon_inspection_days
            st.session_state.offer_10_radon_inspection_days_string = st.session_state.offer_10_radon_inspection_days
        else:
            st.session_state.offer_10_radon_inspection_value = ''
            st.session_state.offer_10_radon_inspection_days = 0
            st.session_state.offer_10_radon_inspection_days_string = days_int_to_string(st.session_state.offer_10_radon_inspection_days)

        if st.session_state.offer_10_septic_inspection_check:
            st.session_state.offer_10_septic_inspection_value = 'Y'
            st.session_state.offer_10_septic_inspection_days = st.session_state.offer_10_septic_inspection_days
            st.session_state.offer_10_septic_inspection_days_string = st.session_state.offer_10_septic_inspection_days
        else:
            st.session_state.offer_10_septic_inspection_value = ''
            st.session_state.offer_10_septic_inspection_days = 0
            st.session_state.offer_10_septic_inspection_days_string = days_int_to_string(st.session_state.offer_10_septic_inspection_days)

        if st.session_state.offer_10_well_inspection_check:
            st.session_state.offer_10_well_inspection_value = 'Y'
            st.session_state.offer_10_well_inspection_days = st.session_state.offer_10_well_inspection_days
            st.session_state.offer_10_well_inspection_days_string = st.session_state.offer_10_well_inspection_days
        else:
            st.session_state.offer_10_well_inspection_value = ''
            st.session_state.offer_10_well_inspection_days = 0
            st.session_state.offer_10_well_inspection_days_string = days_int_to_string(st.session_state.offer_10_well_inspection_days)

        if st.session_state.offer_10_financing_contingency_check:
            st.session_state.offer_10_financing_contingency_value = 'Y'
            st.session_state.offer_10_financing_contingency_days = st.session_state.offer_10_financing_contingency_days
            st.session_state.offer_10_financing_contingency_days_string = st.session_state.offer_10_financing_contingency_days
        else:
            st.session_state.offer_10_financing_contingency_value = ''
            st.session_state.offer_10_financing_contingency_days = 0
            st.session_state.offer_10_financing_contingency_days_string = days_int_to_string(st.session_state.offer_10_financing_contingency_days)

        if st.session_state.offer_10_appraisal_contingency_check:
            st.session_state.offer_10_appraisal_contingency_value = 'Y'
            st.session_state.offer_10_appraisal_contingency_days = st.session_state.offer_10_appraisal_contingency_days
            st.session_state.offer_10_appraisal_contingency_days_string = st.session_state.offer_10_appraisal_contingency_days
        else:
            st.session_state.offer_10_appraisal_contingency_value = ''
            st.session_state.offer_10_appraisal_contingency_days = 0
            st.session_state.offer_10_appraisal_contingency_days_string = days_int_to_string(st.session_state.offer_10_appraisal_contingency_days)

        if st.session_state.offer_10_home_sale_contingency_check:
            st.session_state.offer_10_home_sale_contingency_value = 'Y'
            st.session_state.offer_10_home_sale_contingency_days = st.session_state.offer_10_home_inspection_days
            st.session_state.offer_10_home_sale_contingency_days_string = st.session_state.offer_10_home_sale_contingency_days
        else:
            st.session_state.offer_10_home_sale_contingency_value = ''
            st.session_state.offer_10_home_sale_contingency_days = 0
            st.session_state.offer_10_home_sale_contingency_days_string = days_int_to_string(
                st.session_state.offer_10_home_sale_contingency_days)

        if st.session_state.offer_10_pre_occupancy_request:
            st.session_state.offer_10_pre_occupancy_date = st.session_state.offer_10_update_pre_occupancy_date
        else:
            st.session_state.offer_10_pre_occupancy_date = ''

        if st.session_state.offer_10_post_occupancy_request:
            st.session_state.offer_10_post_occupancy_date = st.session_state.offer_10_update_post_occupancy_date
        else:
            st.session_state.offer_10_post_occupancy_date = ''
            
    def update_offer_11_info_form():
        st.session_state.offer_11_name = st.session_state.update_offer_11_name
        st.session_state.offer_11_settlement_date = st.session_state.update_offer_11_settlement_date
        st.session_state.offer_11_settlement_company = st.session_state.update_offer_11_settlement_company
        st.session_state.offer_11_amt = st.session_state.update_offer_11_amt
        st.session_state.offer_11_emd_amt = st.session_state.update_offer_11_emd_amt
        st.session_state.offer_11_down_pmt_pct = st.session_state.update_offer_11_down_pmt_pct
        st.session_state.offer_11_down_pmt_pct = st.session_state.offer_11_down_pmt_pct / 100
        st.session_state.offer_11_closing_subsidy_pct = st.session_state.offer_11_update_closing_subsidy_pct / 100
        if st.session_state.offer_11_closing_subsidy_radio == 'Percent of Offer Amt (%)':
            st.session_state.offer_11_closing_subsidy_amt = st.session_state.offer_11_closing_subsidy_pct * st.session_state.offer_11_amt
        else:
            st.session_state.offer_11_closing_subsidy_amt = st.session_state.offer_11_closing_subsidy_flat_amt

        if st.session_state.offer_11_home_inspection_check:
            st.session_state.offer_11_home_inspection_value = 'Y'
            st.session_state.offer_11_home_inspection_days = st.session_state.offer_11_home_inspection_days
            st.session_state.offer_11_home_inspection_days_string = st.session_state.offer_11_home_inspection_days
        else:
            st.session_state.offer_11_home_inspection_value = ''
            st.session_state.offer_11_home_inspection_days = 0
            st.session_state.offer_11_home_inspection_days_string = days_int_to_string(st.session_state.offer_11_home_inspection_days)

        if st.session_state.offer_11_radon_inspection_check:
            st.session_state.offer_11_radon_inspection_value = 'Y'
            st.session_state.offer_11_radon_inspection_days = st.session_state.offer_11_radon_inspection_days
            st.session_state.offer_11_radon_inspection_days_string = st.session_state.offer_11_radon_inspection_days
        else:
            st.session_state.offer_11_radon_inspection_value = ''
            st.session_state.offer_11_radon_inspection_days = 0
            st.session_state.offer_11_radon_inspection_days_string = days_int_to_string(st.session_state.offer_11_radon_inspection_days)

        if st.session_state.offer_11_septic_inspection_check:
            st.session_state.offer_11_septic_inspection_value = 'Y'
            st.session_state.offer_11_septic_inspection_days = st.session_state.offer_11_septic_inspection_days
            st.session_state.offer_11_septic_inspection_days_string = st.session_state.offer_11_septic_inspection_days
        else:
            st.session_state.offer_11_septic_inspection_value = ''
            st.session_state.offer_11_septic_inspection_days = 0
            st.session_state.offer_11_septic_inspection_days_string = days_int_to_string(st.session_state.offer_11_septic_inspection_days)

        if st.session_state.offer_11_well_inspection_check:
            st.session_state.offer_11_well_inspection_value = 'Y'
            st.session_state.offer_11_well_inspection_days = st.session_state.offer_11_well_inspection_days
            st.session_state.offer_11_well_inspection_days_string = st.session_state.offer_11_well_inspection_days
        else:
            st.session_state.offer_11_well_inspection_value = ''
            st.session_state.offer_11_well_inspection_days = 0
            st.session_state.offer_11_well_inspection_days_string = days_int_to_string(st.session_state.offer_11_well_inspection_days)

        if st.session_state.offer_11_financing_contingency_check:
            st.session_state.offer_11_financing_contingency_value = 'Y'
            st.session_state.offer_11_financing_contingency_days = st.session_state.offer_11_financing_contingency_days
            st.session_state.offer_11_financing_contingency_days_string = st.session_state.offer_11_financing_contingency_days
        else:
            st.session_state.offer_11_financing_contingency_value = ''
            st.session_state.offer_11_financing_contingency_days = 0
            st.session_state.offer_11_financing_contingency_days_string = days_int_to_string(st.session_state.offer_11_financing_contingency_days)

        if st.session_state.offer_11_appraisal_contingency_check:
            st.session_state.offer_11_appraisal_contingency_value = 'Y'
            st.session_state.offer_11_appraisal_contingency_days = st.session_state.offer_11_appraisal_contingency_days
            st.session_state.offer_11_appraisal_contingency_days_string = st.session_state.offer_11_appraisal_contingency_days
        else:
            st.session_state.offer_11_appraisal_contingency_value = ''
            st.session_state.offer_11_appraisal_contingency_days = 0
            st.session_state.offer_11_appraisal_contingency_days_string = days_int_to_string(st.session_state.offer_11_appraisal_contingency_days)

        if st.session_state.offer_11_home_sale_contingency_check:
            st.session_state.offer_11_home_sale_contingency_value = 'Y'
            st.session_state.offer_11_home_sale_contingency_days = st.session_state.offer_11_home_inspection_days
            st.session_state.offer_11_home_sale_contingency_days_string = st.session_state.offer_11_home_sale_contingency_days
        else:
            st.session_state.offer_11_home_sale_contingency_value = ''
            st.session_state.offer_11_home_sale_contingency_days = 0
            st.session_state.offer_11_home_sale_contingency_days_string = days_int_to_string(
                st.session_state.offer_11_home_sale_contingency_days)

        if st.session_state.offer_11_pre_occupancy_request:
            st.session_state.offer_11_pre_occupancy_date = st.session_state.offer_11_update_pre_occupancy_date
        else:
            st.session_state.offer_11_pre_occupancy_date = ''

        if st.session_state.offer_11_post_occupancy_request:
            st.session_state.offer_11_post_occupancy_date = st.session_state.offer_11_update_post_occupancy_date
        else:
            st.session_state.offer_11_post_occupancy_date = ''
            
    def update_offer_12_info_form():
        st.session_state.offer_12_name = st.session_state.update_offer_12_name
        st.session_state.offer_12_settlement_date = st.session_state.update_offer_12_settlement_date
        st.session_state.offer_12_settlement_company = st.session_state.update_offer_12_settlement_company
        st.session_state.offer_12_amt = st.session_state.update_offer_12_amt
        st.session_state.offer_12_emd_amt = st.session_state.update_offer_12_emd_amt
        st.session_state.offer_12_down_pmt_pct = st.session_state.update_offer_12_down_pmt_pct
        st.session_state.offer_12_down_pmt_pct = st.session_state.offer_12_down_pmt_pct / 100
        st.session_state.offer_12_closing_subsidy_pct = st.session_state.offer_12_update_closing_subsidy_pct / 100
        if st.session_state.offer_12_closing_subsidy_radio == 'Percent of Offer Amt (%)':
            st.session_state.offer_12_closing_subsidy_amt = st.session_state.offer_12_closing_subsidy_pct * st.session_state.offer_12_amt
        else:
            st.session_state.offer_12_closing_subsidy_amt = st.session_state.offer_12_closing_subsidy_flat_amt

        if st.session_state.offer_12_home_inspection_check:
            st.session_state.offer_12_home_inspection_value = 'Y'
            st.session_state.offer_12_home_inspection_days = st.session_state.offer_12_home_inspection_days
            st.session_state.offer_12_home_inspection_days_string = st.session_state.offer_12_home_inspection_days
        else:
            st.session_state.offer_12_home_inspection_value = ''
            st.session_state.offer_12_home_inspection_days = 0
            st.session_state.offer_12_home_inspection_days_string = days_int_to_string(st.session_state.offer_12_home_inspection_days)

        if st.session_state.offer_12_radon_inspection_check:
            st.session_state.offer_12_radon_inspection_value = 'Y'
            st.session_state.offer_12_radon_inspection_days = st.session_state.offer_12_radon_inspection_days
            st.session_state.offer_12_radon_inspection_days_string = st.session_state.offer_12_radon_inspection_days
        else:
            st.session_state.offer_12_radon_inspection_value = ''
            st.session_state.offer_12_radon_inspection_days = 0
            st.session_state.offer_12_radon_inspection_days_string = days_int_to_string(st.session_state.offer_12_radon_inspection_days)

        if st.session_state.offer_12_septic_inspection_check:
            st.session_state.offer_12_septic_inspection_value = 'Y'
            st.session_state.offer_12_septic_inspection_days = st.session_state.offer_12_septic_inspection_days
            st.session_state.offer_12_septic_inspection_days_string = st.session_state.offer_12_septic_inspection_days
        else:
            st.session_state.offer_12_septic_inspection_value = ''
            st.session_state.offer_12_septic_inspection_days = 0
            st.session_state.offer_12_septic_inspection_days_string = days_int_to_string(st.session_state.offer_12_septic_inspection_days)

        if st.session_state.offer_12_well_inspection_check:
            st.session_state.offer_12_well_inspection_value = 'Y'
            st.session_state.offer_12_well_inspection_days = st.session_state.offer_12_well_inspection_days
            st.session_state.offer_12_well_inspection_days_string = st.session_state.offer_12_well_inspection_days
        else:
            st.session_state.offer_12_well_inspection_value = ''
            st.session_state.offer_12_well_inspection_days = 0
            st.session_state.offer_12_well_inspection_days_string = days_int_to_string(st.session_state.offer_12_well_inspection_days)

        if st.session_state.offer_12_financing_contingency_check:
            st.session_state.offer_12_financing_contingency_value = 'Y'
            st.session_state.offer_12_financing_contingency_days = st.session_state.offer_12_financing_contingency_days
            st.session_state.offer_12_financing_contingency_days_string = st.session_state.offer_12_financing_contingency_days
        else:
            st.session_state.offer_12_financing_contingency_value = ''
            st.session_state.offer_12_financing_contingency_days = 0
            st.session_state.offer_12_financing_contingency_days_string = days_int_to_string(st.session_state.offer_12_financing_contingency_days)

        if st.session_state.offer_12_appraisal_contingency_check:
            st.session_state.offer_12_appraisal_contingency_value = 'Y'
            st.session_state.offer_12_appraisal_contingency_days = st.session_state.offer_12_appraisal_contingency_days
            st.session_state.offer_12_appraisal_contingency_days_string = st.session_state.offer_12_appraisal_contingency_days
        else:
            st.session_state.offer_12_appraisal_contingency_value = ''
            st.session_state.offer_12_appraisal_contingency_days = 0
            st.session_state.offer_12_appraisal_contingency_days_string = days_int_to_string(st.session_state.offer_12_appraisal_contingency_days)

        if st.session_state.offer_12_home_sale_contingency_check:
            st.session_state.offer_12_home_sale_contingency_value = 'Y'
            st.session_state.offer_12_home_sale_contingency_days = st.session_state.offer_12_home_inspection_days
            st.session_state.offer_12_home_sale_contingency_days_string = st.session_state.offer_12_home_sale_contingency_days
        else:
            st.session_state.offer_12_home_sale_contingency_value = ''
            st.session_state.offer_12_home_sale_contingency_days = 0
            st.session_state.offer_12_home_sale_contingency_days_string = days_int_to_string(
                st.session_state.offer_12_home_sale_contingency_days)

        if st.session_state.offer_12_pre_occupancy_request:
            st.session_state.offer_12_pre_occupancy_date = st.session_state.offer_12_update_pre_occupancy_date
        else:
            st.session_state.offer_12_pre_occupancy_date = ''

        if st.session_state.offer_12_post_occupancy_request:
            st.session_state.offer_12_post_occupancy_date = st.session_state.offer_12_update_post_occupancy_date
        else:
            st.session_state.offer_12_post_occupancy_date = ''


    with intro_info_container:
        with st.expander('Introduction Data Form'):
            with st.form(key='intro_info_form'):
                st.markdown('##### **Enter Top-Level Form Data**')
                intro_info_col1, intro_info_col2 = st.columns(2)
                with intro_info_col1:
                    st.text_input('Enter the name of the agent preparing this offer comparison', key='update_preparer')
                    st.date_input('Enter the date that this offer comparison was created', key='update_prep_date')
                with intro_info_col2:
                    st.number_input('Number of Offers Being Compared', 1, 12, step=1, key='update_offer_qty')
                intro_info_submit = st.form_submit_button('Submit Information', on_click=update_intro_info_form)

    with property_container:
        with st.expander('Property Data Form'):
            with st.form(key='property_info_form'):
                st.markdown('##### **Enter Property-Related Data**')
                property_info_col1, property_info_col2 = st.columns(2)
                with property_info_col1:
                    st.text_input('Name of the Seller(s)', key='update_seller_name')
                    st.text_input('Property\'s Street Address', key='update_address')
                    st.number_input('Property\'s List Price ($)', 0, 1500000, step=1000, key='update_list_price')
                with property_info_col2:
                    st.number_input('Estimated Payoff - First Trust ($)', 0, 1000000, step=1000, key='update_payoff_amt_first_trust')
                    st.number_input('Estimated Payoff - Second Trust ($)', 0, 1000000, step=1000, key='update_payoff_amt_second_trust')
                    st.number_input('Estimated Annual Tax Amount ($)', 0, 25000, step=1, key='update_annual_tax_amt')
                    st.number_input('Estimated Annual HOA / Condo Fee Amount ($)', 0, 10000, step=1, key='update_annual_hoa_condo_fee_amt')
                property_info_submit = st.form_submit_button('Submit Property Information', on_click=update_property_info_form)

    with common_container:
        with st.expander('Common Data Form'):
            with st.form(key='common_info_form'):
                st.markdown('##### **Enter Information Common To All Offers**')
                brokerage_col, closing_cost_col, misc_col = st.columns(3)
                with brokerage_col:
                    st.markdown('###### **Brokerage Cost Data**')
                    st.number_input('Listing Company Compensation (%)', 0.0, 6.0, step=0.01, format='%.2f', key='update_listing_company_pct')
                    st.number_input('Selling Company Compensation (%)', 0.0, 6.0, step=0.01, format='%.2f', key='update_selling_company_pct')
                    st.number_input('Processing Fee ($)', 0, 10000, step=1, key='update_processing_fee')
                with closing_cost_col:
                    st.markdown('###### **Closing Cost Data**')
                    st.number_input('Settlement Fee Amount ($)', 0, 1000, step=1, key='update_settlement_fee')
                    st.number_input('Deed Preparation Fee Amount ($)', 0, 1000, step=1, key='update_deed_preparation_fee')
                    st.number_input('Release of Liens / Trusts Fee Amount ($)', 0, 1000, step=1, key='update_lien_trust_release_fee')
                    st.number_input('Quantity of Liens / Trusts to be Released', 0, 10, step=1, key='update_lien_trust_release_qty')
                with misc_col:
                    st.markdown('###### **Miscellaneous Cost Data**')
                    st.number_input('Recording Release Fee Amount ($)', 0, 250, step=1, key='update_recording_release_fee')
                    st.number_input('Quantity of Recording Releases', 0, 10, step=1, key='update_recording_release_qty')
                    st.number_input('Grantor\'s Tax Pct (%)', 0.0, 1.0, step=0.01, format='%.2f', key='update_grantors_tax_pct')
                    st.number_input('Congestion Tax Pct (%)', 0.0, 1.0, step=0.01, format='%.2f', key='update_congestion_tax_pct')
                    st.number_input('Pest Inspection Fee Amount ($)', 0, 100, step=1, key='update_pest_inspection_fee')
                    st.number_input('Power of Attorney / Condo Disclosure Fee Amount ($)', 0, 500, step=1, key='poa_condo_disclosure_fee')
                common_info_submit = st.form_submit_button('Submit Common Information', on_click=update_common_info_form)

    with offer_1_container:
        with st.expander('Offer 1 Form'):
            with st.form(key='offer_1_info_form'):
                st.markdown('##### **Enter Offer 1\'s Information**')
                offer_1_col1, offer_1_col2 = st.columns(2)
                with offer_1_col1:
                    st.text_input('Name of Offer', key='update_offer_1_name')
                    st.date_input('Settlement Date', key='update_offer_1_settlement_date')
                    st.text_input('Settlement Company', key='update_offer_1_settlement_company')
                    st.number_input('Offer Amount ($)', 0, 1500000, step=1000, key='update_offer_1_amt')
                    st.number_input('EMD Amount ($)', 0, 50000, step=100, key='update_offer_1_emd_amt')
                with offer_1_col2:
                    st.selectbox('Financing Type', financing_types, key='offer_1_finance_type')
                    st.number_input('Down Payment Pct (%)', 0.0, 100.0, step=0.01, key='update_offer_1_down_pmt_pct')
                    st.radio('Closing Cost Subsidy Radio', ['Percent of Offer Amt (%)', 'Flat $ Amount'], key='offer_1_closing_subsidy_radio')
                    st.number_input('Closing Cost Subsidy of (%):', 0.0, 100.0, step=0.01, key='offer_1_update_closing_subsidy_pct')
                    st.number_input('Closing Cost Subsidy of ($):', 0, 100000, step=50, key='offer_1_closing_subsidy_flat_amt')
                st.write('---')
                st.write('Contingencies and Clauses of the Offer')
                # st.text_input('Contingencies Waived', key='offer_1_contingencies_waved')
                offer_1_cont_col1, offer_1_cont_col2 = st.columns(2)
                with offer_1_cont_col1:
                    st.checkbox('Home Inspection', key='offer_1_home_inspection_check')
                    st.number_input('Home Inspection Days', 0, 45, step=1, key='offer_1_home_inspection_days')
                    st.checkbox('Radon Inspection', key='offer_1_radon_inspection_check')
                    st.number_input('Radon Inspection Days', 0, 45, step=1, key='offer_1_radon_inspection_days')
                    st.checkbox('Septic Inspection', key='offer_1_septic_inspection_check')
                    st.number_input('Septic Inspection Days', 0, 45, step=1, key='offer_1_septic_inspection_days')
                    st.checkbox('Well Inspection', key='offer_1_well_inspection_check')
                    st.number_input('Well Inspection Days', 0, 45, step=1, key='offer_1_well_inspection_days')
                with offer_1_cont_col2:
                    st.checkbox('Financing Contingency', key='offer_1_financing_contingency_check')
                    st.number_input('Financing Contingency Days', 0, 45, step=1, key='offer_1_financing_contingency_days')
                    st.checkbox('Appraisal Contingency', key='offer_1_appraisal_contingency_check')
                    st.number_input('Appraisal Contingency Days', 0, 45, step=1, key='offer_1_appraisal_contingency_days')
                    st.checkbox('Home Sale Contingency', key='offer_1_home_sale_contingency_check')
                    st.number_input('Home Sale Contingency Days', 0, 45, step=1, key='offer_1_home_sale_contingency_days')
                st.write('---')
                st.write('Pre Occupancy')
                offer_1_pre_occupancy_col1, offer_1_pre_occupancy_col2 = st.columns(2)
                with offer_1_pre_occupancy_col1:
                    st.checkbox('Pre Occupancy Request', key='offer_1_pre_occupancy_request')
                    st.number_input('Pre Occupancy Credit to Seller ($)', 0, 25000, step=1, key='offer_1_pre_occupancy_credit_to_seller_amt')
                with offer_1_pre_occupancy_col2:
                    st.date_input('Pre Occupancy Date', key='offer_1_update_pre_occupancy_date')
                st.write('---')
                st.write('Post Occupancy')
                offer_1_post_occupancy_col1, offer_1_post_occupancy_col2 = st.columns(2)
                with offer_1_post_occupancy_col1:
                    st.checkbox('Post Occupancy Request', key='offer_1_post_occupancy_request')
                    st.number_input('Post Occupancy Cost to Seller ($)', 0, 25000, step=1, key='offer_1_post_occupancy_cost_to_seller_amt')
                with offer_1_post_occupancy_col2:
                    st.date_input('Post Occupancy Date', key='offer_1_update_post_occupancy_date')
                offer_1_submit = st.form_submit_button('Submit Offer 1\'s Information', on_click=update_offer_1_info_form)

    with offer_2_container:
        with st.expander('Offer 2 Form'):
            with st.form(key='offer_2_info_form'):
                st.markdown('##### **Enter Offer 2\'s Information**')
                offer_2_col1, offer_2_col2 = st.columns(2)
                with offer_2_col1:
                    st.text_input('Name of Offer', key='update_offer_2_name')
                    st.date_input('Settlement Date', key='update_offer_2_settlement_date')
                    st.text_input('Settlement Company', key='update_offer_2_settlement_company')
                    st.number_input('Offer Amount ($)', 0, 1500000, step=1000, key='update_offer_2_amt')
                    st.number_input('EMD Amount ($)', 0, 50000, step=100, key='update_offer_2_emd_amt')
                with offer_2_col2:
                    st.selectbox('Financing Type', financing_types, key='offer_2_finance_type')
                    st.number_input('Down Payment Pct (%)', 0.0, 100.0, step=0.01, key='update_offer_2_down_pmt_pct')
                    st.radio('Closing Cost Subsidy Radio', ['Percent of Offer Amt (%)', 'Flat $ Amount'], key='offer_2_closing_subsidy_radio')
                    st.number_input('Closing Cost Subsidy of (%):', 0.0, 100.0, step=0.01, key='offer_2_update_closing_subsidy_pct')
                    st.number_input('Closing Cost Subsidy of ($):', 0, 100000, step=50, key='offer_2_closing_subsidy_flat_amt')
                st.write('---')
                st.write('Contingencies and Clauses of the Offer')
                # st.text_input('Contingencies Waived', key='offer_2_contingencies_waved')
                offer_2_cont_col1, offer_2_cont_col2 = st.columns(2)
                with offer_2_cont_col1:
                    st.checkbox('Home Inspection', key='offer_2_home_inspection_check')
                    st.number_input('Home Inspection Days', 0, 45, step=1, key='offer_2_home_inspection_days')
                    st.checkbox('Radon Inspection', key='offer_2_radon_inspection_check')
                    st.number_input('Radon Inspection Days', 0, 45, step=1, key='offer_2_radon_inspection_days')
                    st.checkbox('Septic Inspection', key='offer_2_septic_inspection_check')
                    st.number_input('Septic Inspection Days', 0, 45, step=1, key='offer_2_septic_inspection_days')
                    st.checkbox('Well Inspection', key='offer_2_well_inspection_check')
                    st.number_input('Well Inspection Days', 0, 45, step=1, key='offer_2_well_inspection_days')
                with offer_2_cont_col2:
                    st.checkbox('Financing Contingency', key='offer_2_financing_contingency_check')
                    st.number_input('Financing Contingency Days', 0, 45, step=1, key='offer_2_financing_contingency_days')
                    st.checkbox('Appraisal Contingency', key='offer_2_appraisal_contingency_check')
                    st.number_input('Appraisal Contingency Days', 0, 45, step=1, key='offer_2_appraisal_contingency_days')
                    st.checkbox('Home Sale Contingency', key='offer_2_home_sale_contingency_check')
                    st.number_input('Home Sale Contingency Days', 0, 45, step=1, key='offer_2_home_sale_contingency_days')
                st.write('---')
                st.write('Pre Occupancy')
                offer_2_pre_occupancy_col1, offer_2_pre_occupancy_col2 = st.columns(2)
                with offer_2_pre_occupancy_col1:
                    st.checkbox('Pre Occupancy Request', key='offer_2_pre_occupancy_request')
                    st.number_input('Pre Occupancy Credit to Seller ($)', 0, 25000, step=1,
                              key='offer_2_pre_occupancy_credit_to_seller_amt')
                with offer_2_pre_occupancy_col2:
                    st.date_input('Pre Occupancy Date', key='offer_2_update_pre_occupancy_date')
                st.write('---')
                st.write('Post Occupancy')
                offer_2_post_occupancy_col1, offer_2_post_occupancy_col2 = st.columns(2)
                with offer_2_post_occupancy_col1:
                    st.checkbox('Post Occupancy Request', key='offer_2_post_occupancy_request')
                    st.number_input('Post Occupancy Cost to Seller ($)', 0, 25000, step=1, key='offer_2_post_occupancy_cost_to_seller_amt')
                with offer_2_post_occupancy_col2:
                    st.date_input('Post Occupancy Date', key='offer_2_update_post_occupancy_date')
                offer_2_submit = st.form_submit_button('Submit Offer 2\'s Information', on_click=update_offer_2_info_form)

    with offer_3_container:
        with st.expander('Offer 3 Form'):
            with st.form(key='offer_3_info_form'):
                st.markdown('##### **Enter Offer 3\'s Information**')
                offer_3_col1, offer_3_col2 = st.columns(2)
                with offer_3_col1:
                    st.text_input('Name of Offer', key='update_offer_3_name')
                    st.date_input('Settlement Date', key='update_offer_3_settlement_date')
                    st.text_input('Settlement Company', key='update_offer_3_settlement_company')
                    st.number_input('Offer Amount ($)', 0, 1500000, step=1000, key='update_offer_3_amt')
                    st.number_input('EMD Amount ($)', 0, 50000, step=100, key='update_offer_3_emd_amt')
                with offer_3_col2:
                    st.selectbox('Financing Type', financing_types, key='offer_3_finance_type')
                    st.number_input('Down Payment Pct (%)', 0.0, 100.0, step=0.01, key='update_offer_3_down_pmt_pct')
                    st.radio('Closing Cost Subsidy Radio', ['Percent of Offer Amt (%)', 'Flat $ Amount'], key='offer_3_closing_subsidy_radio')
                    st.number_input('Closing Cost Subsidy of (%):', 0.0, 100.0, step=0.01, key='offer_3_update_closing_subsidy_pct')
                    st.number_input('Closing Cost Subsidy of ($):', 0, 100000, step=50, key='offer_3_closing_subsidy_flat_amt')
                st.write('---')
                st.write('Contingencies and Clauses of the Offer')
                # st.text_input('Contingencies Waived', key='offer_3_contingencies_waved')
                offer_3_cont_col1, offer_3_cont_col2 = st.columns(2)
                with offer_3_cont_col1:
                    st.checkbox('Home Inspection', key='offer_3_home_inspection_check')
                    st.number_input('Home Inspection Days', 0, 45, step=1, key='offer_3_home_inspection_days')
                    st.checkbox('Radon Inspection', key='offer_3_radon_inspection_check')
                    st.number_input('Radon Inspection Days', 0, 45, step=1, key='offer_3_radon_inspection_days')
                    st.checkbox('Septic Inspection', key='offer_3_septic_inspection_check')
                    st.number_input('Septic Inspection Days', 0, 45, step=1, key='offer_3_septic_inspection_days')
                    st.checkbox('Well Inspection', key='offer_3_well_inspection_check')
                    st.number_input('Well Inspection Days', 0, 45, step=1, key='offer_3_well_inspection_days')
                with offer_3_cont_col2:
                    st.checkbox('Financing Contingency', key='offer_3_financing_contingency_check')
                    st.number_input('Financing Contingency Days', 0, 45, step=1, key='offer_3_financing_contingency_days')
                    st.checkbox('Appraisal Contingency', key='offer_3_appraisal_contingency_check')
                    st.number_input('Appraisal Contingency Days', 0, 45, step=1, key='offer_3_appraisal_contingency_days')
                    st.checkbox('Home Sale Contingency', key='offer_3_home_sale_contingency_check')
                    st.number_input('Home Sale Contingency Days', 0, 45, step=1, key='offer_3_home_sale_contingency_days')
                st.write('---')
                st.write('Pre Occupancy')
                offer_3_pre_occupancy_col1, offer_3_pre_occupancy_col2 = st.columns(2)
                with offer_3_pre_occupancy_col1:
                    st.checkbox('Pre Occupancy Request', key='offer_3_pre_occupancy_request')
                    st.number_input('Pre Occupancy Credit to Seller ($)', 0, 25000, step=1,
                              key='offer_3_pre_occupancy_credit_to_seller_amt')
                with offer_3_pre_occupancy_col2:
                    st.date_input('Pre Occupancy Date', key='offer_3_update_pre_occupancy_date')
                st.write('---')
                st.write('Post Occupancy')
                offer_3_post_occupancy_col1, offer_3_post_occupancy_col2 = st.columns(2)
                with offer_3_post_occupancy_col1:
                    st.checkbox('Post Occupancy Request', key='offer_3_post_occupancy_request')
                    st.number_input('Post Occupancy Cost to Seller ($)', 0, 25000, step=1, key='offer_3_post_occupancy_cost_to_seller_amt')
                with offer_3_post_occupancy_col2:
                    st.date_input('Post Occupancy Date', key='offer_3_update_post_occupancy_date')
                offer_3_submit = st.form_submit_button('Submit Offer 3\'s Information', on_click=update_offer_3_info_form)

    with offer_4_container:
        with st.expander('Offer 4 Form'):
            with st.form(key='offer_4_info_form'):
                st.markdown('##### **Enter Offer 4\'s Information**')
                offer_4_col1, offer_4_col2 = st.columns(2)
                with offer_4_col1:
                    st.text_input('Name of Offer', key='update_offer_4_name')
                    st.date_input('Settlement Date', key='update_offer_4_settlement_date')
                    st.text_input('Settlement Company', key='update_offer_4_settlement_company')
                    st.number_input('Offer Amount ($)', 0, 1500000, step=1000, key='update_offer_4_amt')
                    st.number_input('EMD Amount ($)', 0, 50000, step=100, key='update_offer_4_emd_amt')
                with offer_4_col2:
                    st.selectbox('Financing Type', financing_types, key='offer_4_finance_type')
                    st.number_input('Down Payment Pct (%)', 0.0, 100.0, step=0.01, key='update_offer_4_down_pmt_pct')
                    st.radio('Closing Cost Subsidy Radio', ['Percent of Offer Amt (%)', 'Flat $ Amount'], key='offer_4_closing_subsidy_radio')
                    st.number_input('Closing Cost Subsidy of (%):', 0.0, 100.0, step=0.01, key='offer_4_update_closing_subsidy_pct')
                    st.number_input('Closing Cost Subsidy of ($):', 0, 100000, step=50, key='offer_4_closing_subsidy_flat_amt')
                st.write('---')
                st.write('Contingencies and Clauses of the Offer')
                # st.text_input('Contingencies Waived', key='offer_1_contingencies_waved')
                offer_4_cont_col1, offer_4_cont_col2 = st.columns(2)
                with offer_4_cont_col1:
                    st.checkbox('Home Inspection', key='offer_4_home_inspection_check')
                    st.number_input('Home Inspection Days', 0, 45, step=1, key='offer_4_home_inspection_days')
                    st.checkbox('Radon Inspection', key='offer_4_radon_inspection_check')
                    st.number_input('Radon Inspection Days', 0, 45, step=1, key='offer_4_radon_inspection_days')
                    st.checkbox('Septic Inspection', key='offer_4_septic_inspection_check')
                    st.number_input('Septic Inspection Days', 0, 45, step=1, key='offer_4_septic_inspection_days')
                    st.checkbox('Well Inspection', key='offer_4_well_inspection_check')
                    st.number_input('Well Inspection Days', 0, 45, step=1, key='offer_4_well_inspection_days')
                with offer_4_cont_col2:
                    st.checkbox('Financing Contingency', key='offer_4_financing_contingency_check')
                    st.number_input('Financing Contingency Days', 0, 45, step=1, key='offer_4_financing_contingency_days')
                    st.checkbox('Appraisal Contingency', key='offer_4_appraisal_contingency_check')
                    st.number_input('Appraisal Contingency Days', 0, 45, step=1, key='offer_4_appraisal_contingency_days')
                    st.checkbox('Home Sale Contingency', key='offer_4_home_sale_contingency_check')
                    st.number_input('Home Sale Contingency Days', 0, 45, step=1, key='offer_4_home_sale_contingency_days')
                st.write('---')
                st.write('Pre Occupancy')
                offer_4_pre_occupancy_col1, offer_4_pre_occupancy_col2 = st.columns(2)
                with offer_4_pre_occupancy_col1:
                    st.checkbox('Pre Occupancy Request', key='offer_4_pre_occupancy_request')
                    st.number_input('Pre Occupancy Credit to Seller ($)', 0, 25000, step=1, key='offer_4_pre_occupancy_credit_to_seller_amt')
                with offer_4_pre_occupancy_col2:
                    st.date_input('Pre Occupancy Date', key='offer_4_update_pre_occupancy_date')
                st.write('---')
                st.write('Post Occupancy')
                offer_4_post_occupancy_col1, offer_4_post_occupancy_col2 = st.columns(2)
                with offer_4_post_occupancy_col1:
                    st.checkbox('Post Occupancy Request', key='offer_4_post_occupancy_request')
                    st.number_input('Post Occupancy Cost to Seller ($)', 0, 25000, step=1, key='offer_4_post_occupancy_cost_to_seller_amt')
                with offer_4_post_occupancy_col2:
                    st.date_input('Post Occupancy Date', key='offer_4_update_post_occupancy_date')
                offer_4_submit = st.form_submit_button('Submit Offer 4\'s Information', on_click=update_offer_4_info_form)

    with offer_5_container:
        with st.expander('Offer 5 Form'):
            with st.form(key='offer_5_info_form'):
                st.markdown('##### **Enter Offer 5\'s Information**')
                offer_5_col1, offer_5_col2 = st.columns(2)
                with offer_5_col1:
                    st.text_input('Name of Offer', key='update_offer_5_name')
                    st.date_input('Settlement Date', key='update_offer_5_settlement_date')
                    st.text_input('Settlement Company', key='update_offer_5_settlement_company')
                    st.number_input('Offer Amount ($)', 0, 1500000, step=1000, key='update_offer_5_amt')
                    st.number_input('EMD Amount ($)', 0, 50000, step=100, key='update_offer_5_emd_amt')
                with offer_5_col2:
                    st.selectbox('Financing Type', financing_types, key='offer_5_finance_type')
                    st.number_input('Down Payment Pct (%)', 0.0, 100.0, step=0.01, key='update_offer_5_down_pmt_pct')
                    st.radio('Closing Cost Subsidy Radio', ['Percent of Offer Amt (%)', 'Flat $ Amount'], key='offer_5_closing_subsidy_radio')
                    st.number_input('Closing Cost Subsidy of (%):', 0.0, 100.0, step=0.01, key='offer_5_update_closing_subsidy_pct')
                    st.number_input('Closing Cost Subsidy of ($):', 0, 100000, step=50, key='offer_5_closing_subsidy_flat_amt')
                st.write('---')
                st.write('Contingencies and Clauses of the Offer')
                # st.text_input('Contingencies Waived', key='offer_5_contingencies_waved')
                offer_5_cont_col1, offer_5_cont_col2 = st.columns(2)
                with offer_5_cont_col1:
                    st.checkbox('Home Inspection', key='offer_5_home_inspection_check')
                    st.number_input('Home Inspection Days', 0, 45, step=1, key='offer_5_home_inspection_days')
                    st.checkbox('Radon Inspection', key='offer_5_radon_inspection_check')
                    st.number_input('Radon Inspection Days', 0, 45, step=1, key='offer_5_radon_inspection_days')
                    st.checkbox('Septic Inspection', key='offer_5_septic_inspection_check')
                    st.number_input('Septic Inspection Days', 0, 45, step=1, key='offer_5_septic_inspection_days')
                    st.checkbox('Well Inspection', key='offer_5_well_inspection_check')
                    st.number_input('Well Inspection Days', 0, 45, step=1, key='offer_5_well_inspection_days')
                with offer_5_cont_col2:
                    st.checkbox('Financing Contingency', key='offer_5_financing_contingency_check')
                    st.number_input('Financing Contingency Days', 0, 45, step=1, key='offer_5_financing_contingency_days')
                    st.checkbox('Appraisal Contingency', key='offer_5_appraisal_contingency_check')
                    st.number_input('Appraisal Contingency Days', 0, 45, step=1, key='offer_5_appraisal_contingency_days')
                    st.checkbox('Home Sale Contingency', key='offer_5_home_sale_contingency_check')
                    st.number_input('Home Sale Contingency Days', 0, 45, step=1, key='offer_5_home_sale_contingency_days')
                st.write('---')
                st.write('Pre Occupancy')
                offer_5_pre_occupancy_col1, offer_5_pre_occupancy_col2 = st.columns(2)
                with offer_5_pre_occupancy_col1:
                    st.checkbox('Pre Occupancy Request', key='offer_5_pre_occupancy_request')
                    st.number_input('Pre Occupancy Credit to Seller ($)', 0, 25000, step=1, key='offer_5_pre_occupancy_credit_to_seller_amt')
                with offer_5_pre_occupancy_col2:
                    st.date_input('Pre Occupancy Date', key='offer_5_update_pre_occupancy_date')
                st.write('---')
                st.write('Post Occupancy')
                offer_5_post_occupancy_col1, offer_5_post_occupancy_col2 = st.columns(2)
                with offer_5_post_occupancy_col1:
                    st.checkbox('Post Occupancy Request', key='offer_5_post_occupancy_request')
                    st.number_input('Post Occupancy Cost to Seller ($)', 0, 25000, step=1, key='offer_5_post_occupancy_cost_to_seller_amt')
                with offer_5_post_occupancy_col2:
                    st.date_input('Post Occupancy Date', key='offer_5_update_post_occupancy_date')
                offer_5_submit = st.form_submit_button('Submit Offer 5\'s Information', on_click=update_offer_5_info_form)

    with offer_6_container:
        with st.expander('Offer 6 Form'):
            with st.form(key='offer_6_info_form'):
                st.markdown('##### **Enter Offer 6\'s Information**')
                offer_6_col1, offer_6_col2 = st.columns(2)
                with offer_6_col1:
                    st.text_input('Name of Offer', key='update_offer_6_name')
                    st.date_input('Settlement Date', key='update_offer_6_settlement_date')
                    st.text_input('Settlement Company', key='update_offer_6_settlement_company')
                    st.number_input('Offer Amount ($)', 0, 1500000, step=1000, key='update_offer_6_amt')
                    st.number_input('EMD Amount ($)', 0, 50000, step=100, key='update_offer_6_emd_amt')
                with offer_6_col2:
                    st.selectbox('Financing Type', financing_types, key='offer_6_finance_type')
                    st.number_input('Down Payment Pct (%)', 0.0, 100.0, step=0.01, key='update_offer_6_down_pmt_pct')
                    st.radio('Closing Cost Subsidy Radio', ['Percent of Offer Amt (%)', 'Flat $ Amount'], key='offer_6_closing_subsidy_radio')
                    st.number_input('Closing Cost Subsidy of (%):', 0.0, 100.0, step=0.01, key='offer_6_update_closing_subsidy_pct')
                    st.number_input('Closing Cost Subsidy of ($):', 0, 100000, step=50, key='offer_6_closing_subsidy_flat_amt')
                st.write('---')
                st.write('Contingencies and Clauses of the Offer')
                # st.text_input('Contingencies Waived', key='offer_6_contingencies_waved')
                offer_6_cont_col1, offer_6_cont_col2 = st.columns(2)
                with offer_6_cont_col1:
                    st.checkbox('Home Inspection', key='offer_6_home_inspection_check')
                    st.number_input('Home Inspection Days', 0, 45, step=1, key='offer_6_home_inspection_days')
                    st.checkbox('Radon Inspection', key='offer_6_radon_inspection_check')
                    st.number_input('Radon Inspection Days', 0, 45, step=1, key='offer_6_radon_inspection_days')
                    st.checkbox('Septic Inspection', key='offer_6_septic_inspection_check')
                    st.number_input('Septic Inspection Days', 0, 45, step=1, key='offer_6_septic_inspection_days')
                    st.checkbox('Well Inspection', key='offer_6_well_inspection_check')
                    st.number_input('Well Inspection Days', 0, 45, step=1, key='offer_6_well_inspection_days')
                with offer_6_cont_col2:
                    st.checkbox('Financing Contingency', key='offer_6_financing_contingency_check')
                    st.number_input('Financing Contingency Days', 0, 45, step=1, key='offer_6_financing_contingency_days')
                    st.checkbox('Appraisal Contingency', key='offer_6_appraisal_contingency_check')
                    st.number_input('Appraisal Contingency Days', 0, 45, step=1, key='offer_6_appraisal_contingency_days')
                    st.checkbox('Home Sale Contingency', key='offer_6_home_sale_contingency_check')
                    st.number_input('Home Sale Contingency Days', 0, 45, step=1, key='offer_6_home_sale_contingency_days')
                st.write('---')
                st.write('Pre Occupancy')
                offer_6_pre_occupancy_col1, offer_6_pre_occupancy_col2 = st.columns(2)
                with offer_6_pre_occupancy_col1:
                    st.checkbox('Pre Occupancy Request', key='offer_6_pre_occupancy_request')
                    st.number_input('Pre Occupancy Credit to Seller ($)', 0, 25000, step=1, key='offer_6_pre_occupancy_credit_to_seller_amt')
                with offer_6_pre_occupancy_col2:
                    st.date_input('Pre Occupancy Date', key='offer_6_update_pre_occupancy_date')
                st.write('---')
                st.write('Post Occupancy')
                offer_6_post_occupancy_col1, offer_6_post_occupancy_col2 = st.columns(2)
                with offer_6_post_occupancy_col1:
                    st.checkbox('Post Occupancy Request', key='offer_6_post_occupancy_request')
                    st.number_input('Post Occupancy Cost to Seller ($)', 0, 25000, step=1, key='offer_6_post_occupancy_cost_to_seller_amt')
                with offer_6_post_occupancy_col2:
                    st.date_input('Post Occupancy Date', key='offer_6_update_post_occupancy_date')
                offer_6_submit = st.form_submit_button('Submit Offer 6\'s Information', on_click=update_offer_6_info_form)

    with offer_7_container:
        with st.expander('Offer 7 Form'):
            with st.form(key='offer_7_info_form'):
                st.markdown('##### **Enter Offer 7\'s Information**')
                offer_7_col1, offer_7_col2 = st.columns(2)
                with offer_7_col1:
                    st.text_input('Name of Offer', key='update_offer_7_name')
                    st.date_input('Settlement Date', key='update_offer_7_settlement_date')
                    st.text_input('Settlement Company', key='update_offer_7_settlement_company')
                    st.number_input('Offer Amount ($)', 0, 1500000, step=1000, key='update_offer_7_amt')
                    st.number_input('EMD Amount ($)', 0, 50000, step=100, key='update_offer_7_emd_amt')
                with offer_7_col2:
                    st.selectbox('Financing Type', financing_types, key='offer_7_finance_type')
                    st.number_input('Down Payment Pct (%)', 0.0, 100.0, step=0.01, key='update_offer_7_down_pmt_pct')
                    st.radio('Closing Cost Subsidy Radio', ['Percent of Offer Amt (%)', 'Flat $ Amount'], key='offer_7_closing_subsidy_radio')
                    st.number_input('Closing Cost Subsidy of (%):', 0.0, 100.0, step=0.01, key='offer_7_update_closing_subsidy_pct')
                    st.number_input('Closing Cost Subsidy of ($):', 0, 100000, step=50, key='offer_7_closing_subsidy_flat_amt')
                st.write('---')
                st.write('Contingencies and Clauses of the Offer')
                # st.text_input('Contingencies Waived', key='offer_7_contingencies_waved')
                offer_7_cont_col1, offer_7_cont_col2 = st.columns(2)
                with offer_7_cont_col1:
                    st.checkbox('Home Inspection', key='offer_7_home_inspection_check')
                    st.number_input('Home Inspection Days', 0, 45, step=1, key='offer_7_home_inspection_days')
                    st.checkbox('Radon Inspection', key='offer_7_radon_inspection_check')
                    st.number_input('Radon Inspection Days', 0, 45, step=1, key='offer_7_radon_inspection_days')
                    st.checkbox('Septic Inspection', key='offer_7_septic_inspection_check')
                    st.number_input('Septic Inspection Days', 0, 45, step=1, key='offer_7_septic_inspection_days')
                    st.checkbox('Well Inspection', key='offer_7_well_inspection_check')
                    st.number_input('Well Inspection Days', 0, 45, step=1, key='offer_7_well_inspection_days')
                with offer_7_cont_col2:
                    st.checkbox('Financing Contingency', key='offer_7_financing_contingency_check')
                    st.number_input('Financing Contingency Days', 0, 45, step=1, key='offer_7_financing_contingency_days')
                    st.checkbox('Appraisal Contingency', key='offer_7_appraisal_contingency_check')
                    st.number_input('Appraisal Contingency Days', 0, 45, step=1, key='offer_7_appraisal_contingency_days')
                    st.checkbox('Home Sale Contingency', key='offer_7_home_sale_contingency_check')
                    st.number_input('Home Sale Contingency Days', 0, 45, step=1, key='offer_7_home_sale_contingency_days')
                st.write('---')
                st.write('Pre Occupancy')
                offer_7_pre_occupancy_col1, offer_7_pre_occupancy_col2 = st.columns(2)
                with offer_7_pre_occupancy_col1:
                    st.checkbox('Pre Occupancy Request', key='offer_7_pre_occupancy_request')
                    st.number_input('Pre Occupancy Credit to Seller ($)', 0, 25000, step=1, key='offer_7_pre_occupancy_credit_to_seller_amt')
                with offer_7_pre_occupancy_col2:
                    st.date_input('Pre Occupancy Date', key='offer_7_update_pre_occupancy_date')
                st.write('---')
                st.write('Post Occupancy')
                offer_7_post_occupancy_col1, offer_7_post_occupancy_col2 = st.columns(2)
                with offer_7_post_occupancy_col1:
                    st.checkbox('Post Occupancy Request', key='offer_7_post_occupancy_request')
                    st.number_input('Post Occupancy Cost to Seller ($)', 0, 25000, step=1, key='offer_7_post_occupancy_cost_to_seller_amt')
                with offer_7_post_occupancy_col2:
                    st.date_input('Post Occupancy Date', key='offer_7_update_post_occupancy_date')
                offer_7_submit = st.form_submit_button('Submit Offer 7\'s Information', on_click=update_offer_7_info_form)

    with offer_8_container:
        with st.expander('Offer 8 Form'):
            with st.form(key='offer_8_info_form'):
                st.markdown('##### **Enter Offer 8\'s Information**')
                offer_8_col1, offer_8_col2 = st.columns(2)
                with offer_8_col1:
                    st.text_input('Name of Offer', key='update_offer_8_name')
                    st.date_input('Settlement Date', key='update_offer_8_settlement_date')
                    st.text_input('Settlement Company', key='update_offer_8_settlement_company')
                    st.number_input('Offer Amount ($)', 0, 1500000, step=1000, key='update_offer_8_amt')
                    st.number_input('EMD Amount ($)', 0, 50000, step=100, key='update_offer_8_emd_amt')
                with offer_8_col2:
                    st.selectbox('Financing Type', financing_types, key='offer_8_finance_type')
                    st.number_input('Down Payment Pct (%)', 0.0, 100.0, step=0.01, key='update_offer_8_down_pmt_pct')
                    st.radio('Closing Cost Subsidy Radio', ['Percent of Offer Amt (%)', 'Flat $ Amount'], key='offer_8_closing_subsidy_radio')
                    st.number_input('Closing Cost Subsidy of (%):', 0.0, 100.0, step=0.01, key='offer_8_update_closing_subsidy_pct')
                    st.number_input('Closing Cost Subsidy of ($):', 0, 100000, step=50, key='offer_8_closing_subsidy_flat_amt')
                st.write('---')
                st.write('Contingencies and Clauses of the Offer')
                # st.text_input('Contingencies Waived', key='offer_8_contingencies_waved')
                offer_8_cont_col1, offer_8_cont_col2 = st.columns(2)
                with offer_8_cont_col1:
                    st.checkbox('Home Inspection', key='offer_8_home_inspection_check')
                    st.number_input('Home Inspection Days', 0, 45, step=1, key='offer_8_home_inspection_days')
                    st.checkbox('Radon Inspection', key='offer_8_radon_inspection_check')
                    st.number_input('Radon Inspection Days', 0, 45, step=1, key='offer_8_radon_inspection_days')
                    st.checkbox('Septic Inspection', key='offer_8_septic_inspection_check')
                    st.number_input('Septic Inspection Days', 0, 45, step=1, key='offer_8_septic_inspection_days')
                    st.checkbox('Well Inspection', key='offer_8_well_inspection_check')
                    st.number_input('Well Inspection Days', 0, 45, step=1, key='offer_8_well_inspection_days')
                with offer_8_cont_col2:
                    st.checkbox('Financing Contingency', key='offer_8_financing_contingency_check')
                    st.number_input('Financing Contingency Days', 0, 45, step=1, key='offer_8_financing_contingency_days')
                    st.checkbox('Appraisal Contingency', key='offer_8_appraisal_contingency_check')
                    st.number_input('Appraisal Contingency Days', 0, 45, step=1, key='offer_8_appraisal_contingency_days')
                    st.checkbox('Home Sale Contingency', key='offer_8_home_sale_contingency_check')
                    st.number_input('Home Sale Contingency Days', 0, 45, step=1, key='offer_8_home_sale_contingency_days')
                st.write('---')
                st.write('Pre Occupancy')
                offer_8_pre_occupancy_col1, offer_8_pre_occupancy_col2 = st.columns(2)
                with offer_8_pre_occupancy_col1:
                    st.checkbox('Pre Occupancy Request', key='offer_8_pre_occupancy_request')
                    st.number_input('Pre Occupancy Credit to Seller ($)', 0, 25000, step=1, key='offer_8_pre_occupancy_credit_to_seller_amt')
                with offer_8_pre_occupancy_col2:
                    st.date_input('Pre Occupancy Date', key='offer_8_update_pre_occupancy_date')
                st.write('---')
                st.write('Post Occupancy')
                offer_8_post_occupancy_col1, offer_8_post_occupancy_col2 = st.columns(2)
                with offer_8_post_occupancy_col1:
                    st.checkbox('Post Occupancy Request', key='offer_8_post_occupancy_request')
                    st.number_input('Post Occupancy Cost to Seller ($)', 0, 25000, step=1, key='offer_8_post_occupancy_cost_to_seller_amt')
                with offer_8_post_occupancy_col2:
                    st.date_input('Post Occupancy Date', key='offer_8_update_post_occupancy_date')
                offer_8_submit = st.form_submit_button('Submit Offer 8\'s Information', on_click=update_offer_8_info_form)

    with offer_9_container:
        with st.expander('Offer 9 Form'):
            with st.form(key='offer_9_info_form'):
                st.markdown('##### **Enter Offer 9\'s Information**')
                offer_9_col1, offer_9_col2 = st.columns(2)
                with offer_9_col1:
                    st.text_input('Name of Offer', key='update_offer_9_name')
                    st.date_input('Settlement Date', key='update_offer_9_settlement_date')
                    st.text_input('Settlement Company', key='update_offer_9_settlement_company')
                    st.number_input('Offer Amount ($)', 0, 1500000, step=1000, key='update_offer_9_amt')
                    st.number_input('EMD Amount ($)', 0, 50000, step=100, key='update_offer_9_emd_amt')
                with offer_9_col2:
                    st.selectbox('Financing Type', financing_types, key='offer_9_finance_type')
                    st.number_input('Down Payment Pct (%)', 0.0, 100.0, step=0.01, key='update_offer_9_down_pmt_pct')
                    st.radio('Closing Cost Subsidy Radio', ['Percent of Offer Amt (%)', 'Flat $ Amount'], key='offer_9_closing_subsidy_radio')
                    st.number_input('Closing Cost Subsidy of (%):', 0.0, 100.0, step=0.01, key='offer_9_update_closing_subsidy_pct')
                    st.number_input('Closing Cost Subsidy of ($):', 0, 100000, step=50, key='offer_9_closing_subsidy_flat_amt')
                st.write('---')
                st.write('Contingencies and Clauses of the Offer')
                # st.text_input('Contingencies Waived', key='offer_9_contingencies_waved')
                offer_9_cont_col1, offer_9_cont_col2 = st.columns(2)
                with offer_9_cont_col1:
                    st.checkbox('Home Inspection', key='offer_9_home_inspection_check')
                    st.number_input('Home Inspection Days', 0, 45, step=1, key='offer_9_home_inspection_days')
                    st.checkbox('Radon Inspection', key='offer_9_radon_inspection_check')
                    st.number_input('Radon Inspection Days', 0, 45, step=1, key='offer_9_radon_inspection_days')
                    st.checkbox('Septic Inspection', key='offer_9_septic_inspection_check')
                    st.number_input('Septic Inspection Days', 0, 45, step=1, key='offer_9_septic_inspection_days')
                    st.checkbox('Well Inspection', key='offer_9_well_inspection_check')
                    st.number_input('Well Inspection Days', 0, 45, step=1, key='offer_9_well_inspection_days')
                with offer_9_cont_col2:
                    st.checkbox('Financing Contingency', key='offer_9_financing_contingency_check')
                    st.number_input('Financing Contingency Days', 0, 45, step=1, key='offer_9_financing_contingency_days')
                    st.checkbox('Appraisal Contingency', key='offer_9_appraisal_contingency_check')
                    st.number_input('Appraisal Contingency Days', 0, 45, step=1, key='offer_9_appraisal_contingency_days')
                    st.checkbox('Home Sale Contingency', key='offer_9_home_sale_contingency_check')
                    st.number_input('Home Sale Contingency Days', 0, 45, step=1, key='offer_9_home_sale_contingency_days')
                st.write('---')
                st.write('Pre Occupancy')
                offer_9_pre_occupancy_col1, offer_9_pre_occupancy_col2 = st.columns(2)
                with offer_9_pre_occupancy_col1:
                    st.checkbox('Pre Occupancy Request', key='offer_9_pre_occupancy_request')
                    st.number_input('Pre Occupancy Credit to Seller ($)', 0, 25000, step=1, key='offer_9_pre_occupancy_credit_to_seller_amt')
                with offer_9_pre_occupancy_col2:
                    st.date_input('Pre Occupancy Date', key='offer_9_update_pre_occupancy_date')
                st.write('---')
                st.write('Post Occupancy')
                offer_9_post_occupancy_col1, offer_9_post_occupancy_col2 = st.columns(2)
                with offer_9_post_occupancy_col1:
                    st.checkbox('Post Occupancy Request', key='offer_9_post_occupancy_request')
                    st.number_input('Post Occupancy Cost to Seller ($)', 0, 25000, step=1, key='offer_9_post_occupancy_cost_to_seller_amt')
                with offer_9_post_occupancy_col2:
                    st.date_input('Post Occupancy Date', key='offer_9_update_post_occupancy_date')
                offer_9_submit = st.form_submit_button('Submit Offer 9\'s Information', on_click=update_offer_9_info_form)

    with offer_10_container:
        with st.expander('Offer 10 Form'):
            with st.form(key='offer_10_info_form'):
                st.markdown('##### **Enter Offer 10\'s Information**')
                offer_10_col1, offer_10_col2 = st.columns(2)
                with offer_10_col1:
                    st.text_input('Name of Offer', key='update_offer_10_name')
                    st.date_input('Settlement Date', key='update_offer_10_settlement_date')
                    st.text_input('Settlement Company', key='update_offer_10_settlement_company')
                    st.number_input('Offer Amount ($)', 0, 1500000, step=1000, key='update_offer_10_amt')
                    st.number_input('EMD Amount ($)', 0, 50000, step=100, key='update_offer_10_emd_amt')
                with offer_10_col2:
                    st.selectbox('Financing Type', financing_types, key='offer_10_finance_type')
                    st.number_input('Down Payment Pct (%)', 0.0, 100.0, step=0.01, key='update_offer_10_down_pmt_pct')
                    st.radio('Closing Cost Subsidy Radio', ['Percent of Offer Amt (%)', 'Flat $ Amount'], key='offer_10_closing_subsidy_radio')
                    st.number_input('Closing Cost Subsidy of (%):', 0.0, 100.0, step=0.01, key='offer_10_update_closing_subsidy_pct')
                    st.number_input('Closing Cost Subsidy of ($):', 0, 100000, step=50, key='offer_10_closing_subsidy_flat_amt')
                st.write('---')
                st.write('Contingencies and Clauses of the Offer')
                # st.text_input('Contingencies Waived', key='offer_10_contingencies_waved')
                offer_10_cont_col1, offer_10_cont_col2 = st.columns(2)
                with offer_10_cont_col1:
                    st.checkbox('Home Inspection', key='offer_10_home_inspection_check')
                    st.number_input('Home Inspection Days', 0, 45, step=1, key='offer_10_home_inspection_days')
                    st.checkbox('Radon Inspection', key='offer_10_radon_inspection_check')
                    st.number_input('Radon Inspection Days', 0, 45, step=1, key='offer_10_radon_inspection_days')
                    st.checkbox('Septic Inspection', key='offer_10_septic_inspection_check')
                    st.number_input('Septic Inspection Days', 0, 45, step=1, key='offer_10_septic_inspection_days')
                    st.checkbox('Well Inspection', key='offer_10_well_inspection_check')
                    st.number_input('Well Inspection Days', 0, 45, step=1, key='offer_10_well_inspection_days')
                with offer_10_cont_col2:
                    st.checkbox('Financing Contingency', key='offer_10_financing_contingency_check')
                    st.number_input('Financing Contingency Days', 0, 45, step=1, key='offer_10_financing_contingency_days')
                    st.checkbox('Appraisal Contingency', key='offer_10_appraisal_contingency_check')
                    st.number_input('Appraisal Contingency Days', 0, 45, step=1, key='offer_10_appraisal_contingency_days')
                    st.checkbox('Home Sale Contingency', key='offer_10_home_sale_contingency_check')
                    st.number_input('Home Sale Contingency Days', 0, 45, step=1, key='offer_10_home_sale_contingency_days')
                st.write('---')
                st.write('Pre Occupancy')
                offer_10_pre_occupancy_col1, offer_10_pre_occupancy_col2 = st.columns(2)
                with offer_10_pre_occupancy_col1:
                    st.checkbox('Pre Occupancy Request', key='offer_10_pre_occupancy_request')
                    st.number_input('Pre Occupancy Credit to Seller ($)', 0, 25000, step=1, key='offer_10_pre_occupancy_credit_to_seller_amt')
                with offer_10_pre_occupancy_col2:
                    st.date_input('Pre Occupancy Date', key='offer_10_update_pre_occupancy_date')
                st.write('---')
                st.write('Post Occupancy')
                offer_10_post_occupancy_col1, offer_10_post_occupancy_col2 = st.columns(2)
                with offer_10_post_occupancy_col1:
                    st.checkbox('Post Occupancy Request', key='offer_10_post_occupancy_request')
                    st.number_input('Post Occupancy Cost to Seller ($)', 0, 25000, step=1, key='offer_10_post_occupancy_cost_to_seller_amt')
                with offer_10_post_occupancy_col2:
                    st.date_input('Post Occupancy Date', key='offer_10_update_post_occupancy_date')
                offer_10_submit = st.form_submit_button('Submit Offer 10\'s Information', on_click=update_offer_10_info_form)

    with offer_11_container:
        with st.expander('Offer 11 Form'):
            with st.form(key='offer_11_info_form'):
                st.markdown('##### **Enter Offer 11\'s Information**')
                offer_11_col1, offer_11_col2 = st.columns(2)
                with offer_11_col1:
                    st.text_input('Name of Offer', key='update_offer_11_name')
                    st.date_input('Settlement Date', key='update_offer_11_settlement_date')
                    st.text_input('Settlement Company', key='update_offer_11_settlement_company')
                    st.number_input('Offer Amount ($)', 0, 1500000, step=1000, key='update_offer_11_amt')
                    st.number_input('EMD Amount ($)', 0, 50000, step=100, key='update_offer_11_emd_amt')
                with offer_11_col2:
                    st.selectbox('Financing Type', financing_types, key='offer_11_finance_type')
                    st.number_input('Down Payment Pct (%)', 0.0, 100.0, step=0.01, key='update_offer_11_down_pmt_pct')
                    st.radio('Closing Cost Subsidy Radio', ['Percent of Offer Amt (%)', 'Flat $ Amount'], key='offer_11_closing_subsidy_radio')
                    st.number_input('Closing Cost Subsidy of (%):', 0.0, 100.0, step=0.01, key='offer_11_update_closing_subsidy_pct')
                    st.number_input('Closing Cost Subsidy of ($):', 0, 100000, step=50, key='offer_11_closing_subsidy_flat_amt')
                st.write('---')
                st.write('Contingencies and Clauses of the Offer')
                # st.text_input('Contingencies Waived', key='offer_11_contingencies_waved')
                offer_11_cont_col1, offer_11_cont_col2 = st.columns(2)
                with offer_11_cont_col1:
                    st.checkbox('Home Inspection', key='offer_11_home_inspection_check')
                    st.number_input('Home Inspection Days', 0, 45, step=1, key='offer_11_home_inspection_days')
                    st.checkbox('Radon Inspection', key='offer_11_radon_inspection_check')
                    st.number_input('Radon Inspection Days', 0, 45, step=1, key='offer_11_radon_inspection_days')
                    st.checkbox('Septic Inspection', key='offer_11_septic_inspection_check')
                    st.number_input('Septic Inspection Days', 0, 45, step=1, key='offer_11_septic_inspection_days')
                    st.checkbox('Well Inspection', key='offer_11_well_inspection_check')
                    st.number_input('Well Inspection Days', 0, 45, step=1, key='offer_11_well_inspection_days')
                with offer_11_cont_col2:
                    st.checkbox('Financing Contingency', key='offer_11_financing_contingency_check')
                    st.number_input('Financing Contingency Days', 0, 45, step=1, key='offer_11_financing_contingency_days')
                    st.checkbox('Appraisal Contingency', key='offer_11_appraisal_contingency_check')
                    st.number_input('Appraisal Contingency Days', 0, 45, step=1, key='offer_11_appraisal_contingency_days')
                    st.checkbox('Home Sale Contingency', key='offer_11_home_sale_contingency_check')
                    st.number_input('Home Sale Contingency Days', 0, 45, step=1, key='offer_11_home_sale_contingency_days')
                st.write('---')
                st.write('Pre Occupancy')
                offer_11_pre_occupancy_col1, offer_11_pre_occupancy_col2 = st.columns(2)
                with offer_11_pre_occupancy_col1:
                    st.checkbox('Pre Occupancy Request', key='offer_11_pre_occupancy_request')
                    st.number_input('Pre Occupancy Credit to Seller ($)', 0, 25000, step=1, key='offer_11_pre_occupancy_credit_to_seller_amt')
                with offer_11_pre_occupancy_col2:
                    st.date_input('Pre Occupancy Date', key='offer_11_update_pre_occupancy_date')
                st.write('---')
                st.write('Post Occupancy')
                offer_11_post_occupancy_col1, offer_11_post_occupancy_col2 = st.columns(2)
                with offer_11_post_occupancy_col1:
                    st.checkbox('Post Occupancy Request', key='offer_11_post_occupancy_request')
                    st.number_input('Post Occupancy Cost to Seller ($)', 0, 25000, step=1, key='offer_11_post_occupancy_cost_to_seller_amt')
                with offer_11_post_occupancy_col2:
                    st.date_input('Post Occupancy Date', key='offer_11_update_post_occupancy_date')
                offer_11_submit = st.form_submit_button('Submit Offer 11\'s Information', on_click=update_offer_11_info_form)

    with offer_12_container:
        with st.expander('Offer 12 Form'):
            with st.form(key='offer_12_info_form'):
                st.markdown('##### **Enter Offer 12\'s Information**')
                offer_12_col1, offer_12_col2 = st.columns(2)
                with offer_12_col1:
                    st.text_input('Name of Offer', key='update_offer_12_name')
                    st.date_input('Settlement Date', key='update_offer_12_settlement_date')
                    st.text_input('Settlement Company', key='update_offer_12_settlement_company')
                    st.number_input('Offer Amount ($)', 0, 1500000, step=1000, key='update_offer_12_amt')
                    st.number_input('EMD Amount ($)', 0, 50000, step=100, key='update_offer_12_emd_amt')
                with offer_12_col2:
                    st.selectbox('Financing Type', financing_types, key='offer_12_finance_type')
                    st.number_input('Down Payment Pct (%)', 0.0, 100.0, step=0.01, key='update_offer_12_down_pmt_pct')
                    st.radio('Closing Cost Subsidy Radio', ['Percent of Offer Amt (%)', 'Flat $ Amount'], key='offer_12_closing_subsidy_radio')
                    st.number_input('Closing Cost Subsidy of (%):', 0.0, 100.0, step=0.01, key='offer_12_update_closing_subsidy_pct')
                    st.number_input('Closing Cost Subsidy of ($):', 0, 100000, step=50, key='offer_12_closing_subsidy_flat_amt')
                st.write('---')
                st.write('Contingencies and Clauses of the Offer')
                # st.text_input('Contingencies Waived', key='offer_12_contingencies_waved')
                offer_12_cont_col1, offer_12_cont_col2 = st.columns(2)
                with offer_12_cont_col1:
                    st.checkbox('Home Inspection', key='offer_12_home_inspection_check')
                    st.number_input('Home Inspection Days', 0, 45, step=1, key='offer_12_home_inspection_days')
                    st.checkbox('Radon Inspection', key='offer_12_radon_inspection_check')
                    st.number_input('Radon Inspection Days', 0, 45, step=1, key='offer_12_radon_inspection_days')
                    st.checkbox('Septic Inspection', key='offer_12_septic_inspection_check')
                    st.number_input('Septic Inspection Days', 0, 45, step=1, key='offer_12_septic_inspection_days')
                    st.checkbox('Well Inspection', key='offer_12_well_inspection_check')
                    st.number_input('Well Inspection Days', 0, 45, step=1, key='offer_12_well_inspection_days')
                with offer_12_cont_col2:
                    st.checkbox('Financing Contingency', key='offer_12_financing_contingency_check')
                    st.number_input('Financing Contingency Days', 0, 45, step=1, key='offer_12_financing_contingency_days')
                    st.checkbox('Appraisal Contingency', key='offer_12_appraisal_contingency_check')
                    st.number_input('Appraisal Contingency Days', 0, 45, step=1, key='offer_12_appraisal_contingency_days')
                    st.checkbox('Home Sale Contingency', key='offer_12_home_sale_contingency_check')
                    st.number_input('Home Sale Contingency Days', 0, 45, step=1, key='offer_12_home_sale_contingency_days')
                st.write('---')
                st.write('Pre Occupancy')
                offer_12_pre_occupancy_col1, offer_12_pre_occupancy_col2 = st.columns(2)
                with offer_12_pre_occupancy_col1:
                    st.checkbox('Pre Occupancy Request', key='offer_12_pre_occupancy_request')
                    st.number_input('Pre Occupancy Credit to Seller ($)', 0, 25000, step=1, key='offer_12_pre_occupancy_credit_to_seller_amt')
                with offer_12_pre_occupancy_col2:
                    st.date_input('Pre Occupancy Date', key='offer_12_update_pre_occupancy_date')
                st.write('---')
                st.write('Post Occupancy')
                offer_12_post_occupancy_col1, offer_12_post_occupancy_col2 = st.columns(2)
                with offer_12_post_occupancy_col1:
                    st.checkbox('Post Occupancy Request', key='offer_12_post_occupancy_request')
                    st.number_input('Post Occupancy Cost to Seller ($)', 0, 25000, step=1, key='offer_12_post_occupancy_cost_to_seller_amt')
                with offer_12_post_occupancy_col2:
                    st.date_input('Post Occupancy Date', key='offer_12_update_post_occupancy_date')
                offer_12_submit = st.form_submit_button('Submit Offer 12\'s Information', on_click=update_offer_12_info_form)

    # st.write(st.session_state)

    offer_comparison_form = comparison_inputs_to_excel(
        agent=st.session_state.preparer,
        date=st.session_state.prep_date,
        offer_qty=st.session_state.offer_qty,
        seller_name=st.session_state.seller_name,
        seller_address=st.session_state.address,
        list_price=st.session_state.list_price,
        first_trust=st.session_state.payoff_amt_first_trust,
        second_trust=st.session_state.payoff_amt_second_trust,
        annual_taxes=st.session_state.annual_tax_amt,
        prorated_taxes=st.session_state.prorated_annual_tax_amt,
        annual_hoa_condo_fees=st.session_state.update_annual_hoa_condo_fee_amt,
        prorated_hoa_condo_fees=st.session_state.prorated_annual_hoa_condo_fee_amt,
        listing_company_pct=st.session_state.listing_company_pct,
        selling_company_pct=st.session_state.selling_company_pct,
        processing_fee=st.session_state.processing_fee,
        settlement_fee=st.session_state.settlement_fee,
        deed_preparation_fee=st.session_state.deed_preparation_fee,
        lien_trust_release_fee=st.session_state.lien_trust_release_fee,
        lien_trust_release_qty=st.session_state.lien_trust_release_qty,
        recording_fee=st.session_state.recording_release_fee,
        recording_trusts_liens_qty=st.session_state.recording_release_qty,
        grantors_tax_pct=st.session_state.grantors_tax_pct,
        congestion_tax_pct=st.session_state.congestion_tax_pct,
        pest_inspection_fee=st.session_state.pest_inspection_fee,
        poa_condo_disclosure_fee=st.session_state.poa_condo_disclosure_fee,
        offer_1_name=st.session_state.offer_1_name,
        offer_1_amt=st.session_state.offer_1_amt,
        offer_1_down_pmt_pct=st.session_state.offer_1_down_pmt_pct,
        offer_1_settlement_date=st.session_state.offer_1_settlement_date,
        offer_1_settlement_company=st.session_state.offer_1_settlement_company,
        offer_1_emd_amt=st.session_state.offer_1_emd_amt,
        offer_1_financing_type=st.session_state.offer_1_finance_type,
        offer_1_home_inspection_check=st.session_state.offer_1_home_inspection_value,
        offer_1_home_inspection_days=st.session_state.offer_1_home_inspection_days_string,
        offer_1_radon_inspection_check=st.session_state.offer_1_radon_inspection_value,
        offer_1_radon_inspection_days=st.session_state.offer_1_radon_inspection_days_string,
        offer_1_septic_inspection_check=st.session_state.offer_1_septic_inspection_value,
        offer_1_septic_inspection_days=st.session_state.offer_1_septic_inspection_days_string,
        offer_1_well_inspection_check=st.session_state.offer_1_well_inspection_value,
        offer_1_well_inspection_days=st.session_state.offer_1_well_inspection_days_string,
        offer_1_finance_contingency_check=st.session_state.offer_1_financing_contingency_value,
        offer_1_finance_contingency_days=st.session_state.offer_1_financing_contingency_days_string,
        offer_1_appraisal_contingency_check=st.session_state.offer_1_appraisal_contingency_value,
        offer_1_appraisal_contingency_days=st.session_state.offer_1_appraisal_contingency_days_string,
        offer_1_home_sale_contingency_check=st.session_state.offer_1_home_sale_contingency_value,
        offer_1_home_sale_contingency_days=st.session_state.offer_1_home_sale_contingency_days_string,
        offer_1_pre_occupancy_date=st.session_state.offer_1_pre_occupancy_date,
        offer_1_post_occupancy_date=st.session_state.offer_1_post_occupancy_date,
        offer_1_closing_cost_subsidy_amt=st.session_state.offer_1_closing_subsidy_amt,
        offer_1_pre_occupancy_credit_amt=st.session_state.offer_1_pre_occupancy_credit_to_seller_amt,
        offer_1_post_occupancy_cost_amt=st.session_state.offer_1_post_occupancy_cost_to_seller_amt,
        offer_2_name=st.session_state.offer_2_name,
        offer_2_amt=st.session_state.offer_2_amt,
        offer_2_down_pmt_pct=st.session_state.offer_2_down_pmt_pct,
        offer_2_settlement_date=st.session_state.offer_2_settlement_date,
        offer_2_settlement_company=st.session_state.offer_2_settlement_company,
        offer_2_emd_amt=st.session_state.offer_2_emd_amt,
        offer_2_financing_type=st.session_state.offer_2_finance_type,
        offer_2_home_inspection_check=st.session_state.offer_2_home_inspection_value,
        offer_2_home_inspection_days=st.session_state.offer_2_home_inspection_days_string,
        offer_2_radon_inspection_check=st.session_state.offer_2_radon_inspection_value,
        offer_2_radon_inspection_days=st.session_state.offer_2_radon_inspection_days_string,
        offer_2_septic_inspection_check=st.session_state.offer_2_septic_inspection_value,
        offer_2_septic_inspection_days=st.session_state.offer_2_septic_inspection_days_string,
        offer_2_well_inspection_check=st.session_state.offer_2_well_inspection_value,
        offer_2_well_inspection_days=st.session_state.offer_2_well_inspection_days_string,
        offer_2_finance_contingency_check=st.session_state.offer_2_financing_contingency_value,
        offer_2_finance_contingency_days=st.session_state.offer_2_financing_contingency_days_string,
        offer_2_appraisal_contingency_check=st.session_state.offer_2_appraisal_contingency_value,
        offer_2_appraisal_contingency_days=st.session_state.offer_2_appraisal_contingency_days_string,
        offer_2_home_sale_contingency_check=st.session_state.offer_2_home_sale_contingency_value,
        offer_2_home_sale_contingency_days=st.session_state.offer_2_home_sale_contingency_days_string,
        offer_2_pre_occupancy_date=st.session_state.offer_2_pre_occupancy_date,
        offer_2_post_occupancy_date=st.session_state.offer_2_post_occupancy_date,
        offer_2_closing_cost_subsidy_amt=st.session_state.offer_2_closing_subsidy_amt,
        offer_2_pre_occupancy_credit_amt=st.session_state.offer_2_pre_occupancy_credit_to_seller_amt,
        offer_2_post_occupancy_cost_amt=st.session_state.offer_2_post_occupancy_cost_to_seller_amt,
        offer_3_name=st.session_state.offer_3_name,
        offer_3_amt=st.session_state.offer_3_amt,
        offer_3_down_pmt_pct=st.session_state.offer_3_down_pmt_pct,
        offer_3_settlement_date=st.session_state.offer_3_settlement_date,
        offer_3_settlement_company=st.session_state.offer_3_settlement_company,
        offer_3_emd_amt=st.session_state.offer_3_emd_amt,
        offer_3_financing_type=st.session_state.offer_3_finance_type,
        offer_3_home_inspection_check=st.session_state.offer_3_home_inspection_value,
        offer_3_home_inspection_days=st.session_state.offer_3_home_inspection_days_string,
        offer_3_radon_inspection_check=st.session_state.offer_3_radon_inspection_value,
        offer_3_radon_inspection_days=st.session_state.offer_3_radon_inspection_days_string,
        offer_3_septic_inspection_check=st.session_state.offer_3_septic_inspection_value,
        offer_3_septic_inspection_days=st.session_state.offer_3_septic_inspection_days_string,
        offer_3_well_inspection_check=st.session_state.offer_3_well_inspection_value,
        offer_3_well_inspection_days=st.session_state.offer_3_well_inspection_days_string,
        offer_3_finance_contingency_check=st.session_state.offer_3_financing_contingency_value,
        offer_3_finance_contingency_days=st.session_state.offer_3_financing_contingency_days_string,
        offer_3_appraisal_contingency_check=st.session_state.offer_3_appraisal_contingency_value,
        offer_3_appraisal_contingency_days=st.session_state.offer_3_appraisal_contingency_days_string,
        offer_3_home_sale_contingency_check=st.session_state.offer_3_home_sale_contingency_value,
        offer_3_home_sale_contingency_days=st.session_state.offer_3_home_sale_contingency_days_string,
        offer_3_pre_occupancy_date=st.session_state.offer_3_pre_occupancy_date,
        offer_3_post_occupancy_date=st.session_state.offer_3_post_occupancy_date,
        offer_3_closing_cost_subsidy_amt=st.session_state.offer_3_closing_subsidy_amt,
        offer_3_pre_occupancy_credit_amt=st.session_state.offer_3_pre_occupancy_credit_to_seller_amt,
        offer_3_post_occupancy_cost_amt=st.session_state.offer_3_post_occupancy_cost_to_seller_amt,
        offer_4_name=st.session_state.offer_4_name,
        offer_4_amt=st.session_state.offer_4_amt,
        offer_4_down_pmt_pct=st.session_state.offer_4_down_pmt_pct,
        offer_4_settlement_date=st.session_state.offer_4_settlement_date,
        offer_4_settlement_company=st.session_state.offer_4_settlement_company,
        offer_4_emd_amt=st.session_state.offer_4_emd_amt,
        offer_4_financing_type=st.session_state.offer_4_finance_type,
        offer_4_home_inspection_check=st.session_state.offer_4_home_inspection_value,
        offer_4_home_inspection_days=st.session_state.offer_4_home_inspection_days_string,
        offer_4_radon_inspection_check=st.session_state.offer_4_radon_inspection_value,
        offer_4_radon_inspection_days=st.session_state.offer_4_radon_inspection_days_string,
        offer_4_septic_inspection_check=st.session_state.offer_4_septic_inspection_value,
        offer_4_septic_inspection_days=st.session_state.offer_4_septic_inspection_days_string,
        offer_4_well_inspection_check=st.session_state.offer_4_well_inspection_value,
        offer_4_well_inspection_days=st.session_state.offer_4_well_inspection_days_string,
        offer_4_finance_contingency_check=st.session_state.offer_4_financing_contingency_value,
        offer_4_finance_contingency_days=st.session_state.offer_4_financing_contingency_days_string,
        offer_4_appraisal_contingency_check=st.session_state.offer_4_appraisal_contingency_value,
        offer_4_appraisal_contingency_days=st.session_state.offer_4_appraisal_contingency_days_string,
        offer_4_home_sale_contingency_check=st.session_state.offer_4_home_sale_contingency_value,
        offer_4_home_sale_contingency_days=st.session_state.offer_4_home_sale_contingency_days_string,
        offer_4_pre_occupancy_date=st.session_state.offer_4_pre_occupancy_date,
        offer_4_post_occupancy_date=st.session_state.offer_4_post_occupancy_date,
        offer_4_closing_cost_subsidy_amt=st.session_state.offer_4_closing_subsidy_amt,
        offer_4_pre_occupancy_credit_amt=st.session_state.offer_4_pre_occupancy_credit_to_seller_amt,
        offer_4_post_occupancy_cost_amt=st.session_state.offer_4_post_occupancy_cost_to_seller_amt,
        offer_5_name=st.session_state.offer_5_name,
        offer_5_amt=st.session_state.offer_5_amt,
        offer_5_down_pmt_pct=st.session_state.offer_5_down_pmt_pct,
        offer_5_settlement_date=st.session_state.offer_5_settlement_date,
        offer_5_settlement_company=st.session_state.offer_5_settlement_company,
        offer_5_emd_amt=st.session_state.offer_5_emd_amt,
        offer_5_financing_type=st.session_state.offer_5_finance_type,
        offer_5_home_inspection_check=st.session_state.offer_5_home_inspection_value,
        offer_5_home_inspection_days=st.session_state.offer_5_home_inspection_days_string,
        offer_5_radon_inspection_check=st.session_state.offer_5_radon_inspection_value,
        offer_5_radon_inspection_days=st.session_state.offer_5_radon_inspection_days_string,
        offer_5_septic_inspection_check=st.session_state.offer_5_septic_inspection_value,
        offer_5_septic_inspection_days=st.session_state.offer_5_septic_inspection_days_string,
        offer_5_well_inspection_check=st.session_state.offer_5_well_inspection_value,
        offer_5_well_inspection_days=st.session_state.offer_5_well_inspection_days_string,
        offer_5_finance_contingency_check=st.session_state.offer_5_financing_contingency_value,
        offer_5_finance_contingency_days=st.session_state.offer_5_financing_contingency_days_string,
        offer_5_appraisal_contingency_check=st.session_state.offer_5_appraisal_contingency_value,
        offer_5_appraisal_contingency_days=st.session_state.offer_5_appraisal_contingency_days_string,
        offer_5_home_sale_contingency_check=st.session_state.offer_5_home_sale_contingency_value,
        offer_5_home_sale_contingency_days=st.session_state.offer_5_home_sale_contingency_days_string,
        offer_5_pre_occupancy_date=st.session_state.offer_5_pre_occupancy_date,
        offer_5_post_occupancy_date=st.session_state.offer_5_post_occupancy_date,
        offer_5_closing_cost_subsidy_amt=st.session_state.offer_5_closing_subsidy_amt,
        offer_5_pre_occupancy_credit_amt=st.session_state.offer_5_pre_occupancy_credit_to_seller_amt,
        offer_5_post_occupancy_cost_amt=st.session_state.offer_5_post_occupancy_cost_to_seller_amt,
        offer_6_name=st.session_state.offer_6_name,
        offer_6_amt=st.session_state.offer_6_amt,
        offer_6_down_pmt_pct=st.session_state.offer_6_down_pmt_pct,
        offer_6_settlement_date=st.session_state.offer_6_settlement_date,
        offer_6_settlement_company=st.session_state.offer_6_settlement_company,
        offer_6_emd_amt=st.session_state.offer_6_emd_amt,
        offer_6_financing_type=st.session_state.offer_6_finance_type,
        offer_6_home_inspection_check=st.session_state.offer_6_home_inspection_value,
        offer_6_home_inspection_days=st.session_state.offer_6_home_inspection_days_string,
        offer_6_radon_inspection_check=st.session_state.offer_6_radon_inspection_value,
        offer_6_radon_inspection_days=st.session_state.offer_6_radon_inspection_days_string,
        offer_6_septic_inspection_check=st.session_state.offer_6_septic_inspection_value,
        offer_6_septic_inspection_days=st.session_state.offer_6_septic_inspection_days_string,
        offer_6_well_inspection_check=st.session_state.offer_6_well_inspection_value,
        offer_6_well_inspection_days=st.session_state.offer_6_well_inspection_days_string,
        offer_6_finance_contingency_check=st.session_state.offer_6_financing_contingency_value,
        offer_6_finance_contingency_days=st.session_state.offer_6_financing_contingency_days_string,
        offer_6_appraisal_contingency_check=st.session_state.offer_6_appraisal_contingency_value,
        offer_6_appraisal_contingency_days=st.session_state.offer_6_appraisal_contingency_days_string,
        offer_6_home_sale_contingency_check=st.session_state.offer_6_home_sale_contingency_value,
        offer_6_home_sale_contingency_days=st.session_state.offer_6_home_sale_contingency_days_string,
        offer_6_pre_occupancy_date=st.session_state.offer_6_pre_occupancy_date,
        offer_6_post_occupancy_date=st.session_state.offer_6_post_occupancy_date,
        offer_6_closing_cost_subsidy_amt=st.session_state.offer_6_closing_subsidy_amt,
        offer_6_pre_occupancy_credit_amt=st.session_state.offer_6_pre_occupancy_credit_to_seller_amt,
        offer_6_post_occupancy_cost_amt=st.session_state.offer_6_post_occupancy_cost_to_seller_amt,
        offer_7_name=st.session_state.offer_7_name,
        offer_7_amt=st.session_state.offer_7_amt,
        offer_7_down_pmt_pct=st.session_state.offer_7_down_pmt_pct,
        offer_7_settlement_date=st.session_state.offer_7_settlement_date,
        offer_7_settlement_company=st.session_state.offer_7_settlement_company,
        offer_7_emd_amt=st.session_state.offer_7_emd_amt,
        offer_7_financing_type=st.session_state.offer_7_finance_type,
        offer_7_home_inspection_check=st.session_state.offer_7_home_inspection_value,
        offer_7_home_inspection_days=st.session_state.offer_7_home_inspection_days_string,
        offer_7_radon_inspection_check=st.session_state.offer_7_radon_inspection_value,
        offer_7_radon_inspection_days=st.session_state.offer_7_radon_inspection_days_string,
        offer_7_septic_inspection_check=st.session_state.offer_7_septic_inspection_value,
        offer_7_septic_inspection_days=st.session_state.offer_7_septic_inspection_days_string,
        offer_7_well_inspection_check=st.session_state.offer_7_well_inspection_value,
        offer_7_well_inspection_days=st.session_state.offer_7_well_inspection_days_string,
        offer_7_finance_contingency_check=st.session_state.offer_7_financing_contingency_value,
        offer_7_finance_contingency_days=st.session_state.offer_7_financing_contingency_days_string,
        offer_7_appraisal_contingency_check=st.session_state.offer_7_appraisal_contingency_value,
        offer_7_appraisal_contingency_days=st.session_state.offer_7_appraisal_contingency_days_string,
        offer_7_home_sale_contingency_check=st.session_state.offer_7_home_sale_contingency_value,
        offer_7_home_sale_contingency_days=st.session_state.offer_7_home_sale_contingency_days_string,
        offer_7_pre_occupancy_date=st.session_state.offer_7_pre_occupancy_date,
        offer_7_post_occupancy_date=st.session_state.offer_7_post_occupancy_date,
        offer_7_closing_cost_subsidy_amt=st.session_state.offer_7_closing_subsidy_amt,
        offer_7_pre_occupancy_credit_amt=st.session_state.offer_7_pre_occupancy_credit_to_seller_amt,
        offer_7_post_occupancy_cost_amt=st.session_state.offer_7_post_occupancy_cost_to_seller_amt,
        offer_8_name=st.session_state.offer_8_name,
        offer_8_amt=st.session_state.offer_8_amt,
        offer_8_down_pmt_pct=st.session_state.offer_8_down_pmt_pct,
        offer_8_settlement_date=st.session_state.offer_8_settlement_date,
        offer_8_settlement_company=st.session_state.offer_8_settlement_company,
        offer_8_emd_amt=st.session_state.offer_8_emd_amt,
        offer_8_financing_type=st.session_state.offer_8_finance_type,
        offer_8_home_inspection_check=st.session_state.offer_8_home_inspection_value,
        offer_8_home_inspection_days=st.session_state.offer_8_home_inspection_days_string,
        offer_8_radon_inspection_check=st.session_state.offer_8_radon_inspection_value,
        offer_8_radon_inspection_days=st.session_state.offer_8_radon_inspection_days_string,
        offer_8_septic_inspection_check=st.session_state.offer_8_septic_inspection_value,
        offer_8_septic_inspection_days=st.session_state.offer_8_septic_inspection_days_string,
        offer_8_well_inspection_check=st.session_state.offer_8_well_inspection_value,
        offer_8_well_inspection_days=st.session_state.offer_8_well_inspection_days_string,
        offer_8_finance_contingency_check=st.session_state.offer_8_financing_contingency_value,
        offer_8_finance_contingency_days=st.session_state.offer_8_financing_contingency_days_string,
        offer_8_appraisal_contingency_check=st.session_state.offer_8_appraisal_contingency_value,
        offer_8_appraisal_contingency_days=st.session_state.offer_8_appraisal_contingency_days_string,
        offer_8_home_sale_contingency_check=st.session_state.offer_8_home_sale_contingency_value,
        offer_8_home_sale_contingency_days=st.session_state.offer_8_home_sale_contingency_days_string,
        offer_8_pre_occupancy_date=st.session_state.offer_8_pre_occupancy_date,
        offer_8_post_occupancy_date=st.session_state.offer_8_post_occupancy_date,
        offer_8_closing_cost_subsidy_amt=st.session_state.offer_8_closing_subsidy_amt,
        offer_8_pre_occupancy_credit_amt=st.session_state.offer_8_pre_occupancy_credit_to_seller_amt,
        offer_8_post_occupancy_cost_amt=st.session_state.offer_8_post_occupancy_cost_to_seller_amt,
        offer_9_name=st.session_state.offer_9_name,
        offer_9_amt=st.session_state.offer_9_amt,
        offer_9_down_pmt_pct=st.session_state.offer_9_down_pmt_pct,
        offer_9_settlement_date=st.session_state.offer_9_settlement_date,
        offer_9_settlement_company=st.session_state.offer_9_settlement_company,
        offer_9_emd_amt=st.session_state.offer_9_emd_amt,
        offer_9_financing_type=st.session_state.offer_9_finance_type,
        offer_9_home_inspection_check=st.session_state.offer_9_home_inspection_value,
        offer_9_home_inspection_days=st.session_state.offer_9_home_inspection_days_string,
        offer_9_radon_inspection_check=st.session_state.offer_9_radon_inspection_value,
        offer_9_radon_inspection_days=st.session_state.offer_9_radon_inspection_days_string,
        offer_9_septic_inspection_check=st.session_state.offer_9_septic_inspection_value,
        offer_9_septic_inspection_days=st.session_state.offer_9_septic_inspection_days_string,
        offer_9_well_inspection_check=st.session_state.offer_9_well_inspection_value,
        offer_9_well_inspection_days=st.session_state.offer_9_well_inspection_days_string,
        offer_9_finance_contingency_check=st.session_state.offer_9_financing_contingency_value,
        offer_9_finance_contingency_days=st.session_state.offer_9_financing_contingency_days_string,
        offer_9_appraisal_contingency_check=st.session_state.offer_9_appraisal_contingency_value,
        offer_9_appraisal_contingency_days=st.session_state.offer_9_appraisal_contingency_days_string,
        offer_9_home_sale_contingency_check=st.session_state.offer_9_home_sale_contingency_value,
        offer_9_home_sale_contingency_days=st.session_state.offer_9_home_sale_contingency_days_string,
        offer_9_pre_occupancy_date=st.session_state.offer_9_pre_occupancy_date,
        offer_9_post_occupancy_date=st.session_state.offer_9_post_occupancy_date,
        offer_9_closing_cost_subsidy_amt=st.session_state.offer_9_closing_subsidy_amt,
        offer_9_pre_occupancy_credit_amt=st.session_state.offer_9_pre_occupancy_credit_to_seller_amt,
        offer_9_post_occupancy_cost_amt=st.session_state.offer_9_post_occupancy_cost_to_seller_amt,
        offer_10_name=st.session_state.offer_10_name,
        offer_10_amt=st.session_state.offer_10_amt,
        offer_10_down_pmt_pct=st.session_state.offer_10_down_pmt_pct,
        offer_10_settlement_date=st.session_state.offer_10_settlement_date,
        offer_10_settlement_company=st.session_state.offer_10_settlement_company,
        offer_10_emd_amt=st.session_state.offer_10_emd_amt,
        offer_10_financing_type=st.session_state.offer_10_finance_type,
        offer_10_home_inspection_check=st.session_state.offer_10_home_inspection_value,
        offer_10_home_inspection_days=st.session_state.offer_10_home_inspection_days_string,
        offer_10_radon_inspection_check=st.session_state.offer_10_radon_inspection_value,
        offer_10_radon_inspection_days=st.session_state.offer_10_radon_inspection_days_string,
        offer_10_septic_inspection_check=st.session_state.offer_10_septic_inspection_value,
        offer_10_septic_inspection_days=st.session_state.offer_10_septic_inspection_days_string,
        offer_10_well_inspection_check=st.session_state.offer_10_well_inspection_value,
        offer_10_well_inspection_days=st.session_state.offer_10_well_inspection_days_string,
        offer_10_finance_contingency_check=st.session_state.offer_10_financing_contingency_value,
        offer_10_finance_contingency_days=st.session_state.offer_10_financing_contingency_days_string,
        offer_10_appraisal_contingency_check=st.session_state.offer_10_appraisal_contingency_value,
        offer_10_appraisal_contingency_days=st.session_state.offer_10_appraisal_contingency_days_string,
        offer_10_home_sale_contingency_check=st.session_state.offer_10_home_sale_contingency_value,
        offer_10_home_sale_contingency_days=st.session_state.offer_10_home_sale_contingency_days_string,
        offer_10_pre_occupancy_date=st.session_state.offer_10_pre_occupancy_date,
        offer_10_post_occupancy_date=st.session_state.offer_10_post_occupancy_date,
        offer_10_closing_cost_subsidy_amt=st.session_state.offer_10_closing_subsidy_amt,
        offer_10_pre_occupancy_credit_amt=st.session_state.offer_10_pre_occupancy_credit_to_seller_amt,
        offer_10_post_occupancy_cost_amt=st.session_state.offer_10_post_occupancy_cost_to_seller_amt,
        offer_11_name=st.session_state.offer_11_name,
        offer_11_amt=st.session_state.offer_11_amt,
        offer_11_down_pmt_pct=st.session_state.offer_11_down_pmt_pct,
        offer_11_settlement_date=st.session_state.offer_11_settlement_date,
        offer_11_settlement_company=st.session_state.offer_11_settlement_company,
        offer_11_emd_amt=st.session_state.offer_11_emd_amt,
        offer_11_financing_type=st.session_state.offer_11_finance_type,
        offer_11_home_inspection_check=st.session_state.offer_11_home_inspection_value,
        offer_11_home_inspection_days=st.session_state.offer_11_home_inspection_days_string,
        offer_11_radon_inspection_check=st.session_state.offer_11_radon_inspection_value,
        offer_11_radon_inspection_days=st.session_state.offer_11_radon_inspection_days_string,
        offer_11_septic_inspection_check=st.session_state.offer_11_septic_inspection_value,
        offer_11_septic_inspection_days=st.session_state.offer_11_septic_inspection_days_string,
        offer_11_well_inspection_check=st.session_state.offer_11_well_inspection_value,
        offer_11_well_inspection_days=st.session_state.offer_11_well_inspection_days_string,
        offer_11_finance_contingency_check=st.session_state.offer_11_financing_contingency_value,
        offer_11_finance_contingency_days=st.session_state.offer_11_financing_contingency_days_string,
        offer_11_appraisal_contingency_check=st.session_state.offer_11_appraisal_contingency_value,
        offer_11_appraisal_contingency_days=st.session_state.offer_11_appraisal_contingency_days_string,
        offer_11_home_sale_contingency_check=st.session_state.offer_11_home_sale_contingency_value,
        offer_11_home_sale_contingency_days=st.session_state.offer_11_home_sale_contingency_days_string,
        offer_11_pre_occupancy_date=st.session_state.offer_11_pre_occupancy_date,
        offer_11_post_occupancy_date=st.session_state.offer_11_post_occupancy_date,
        offer_11_closing_cost_subsidy_amt=st.session_state.offer_11_closing_subsidy_amt,
        offer_11_pre_occupancy_credit_amt=st.session_state.offer_11_pre_occupancy_credit_to_seller_amt,
        offer_11_post_occupancy_cost_amt=st.session_state.offer_11_post_occupancy_cost_to_seller_amt,
        offer_12_name=st.session_state.offer_12_name,
        offer_12_amt=st.session_state.offer_12_amt,
        offer_12_down_pmt_pct=st.session_state.offer_12_down_pmt_pct,
        offer_12_settlement_date=st.session_state.offer_12_settlement_date,
        offer_12_settlement_company=st.session_state.offer_12_settlement_company,
        offer_12_emd_amt=st.session_state.offer_12_emd_amt,
        offer_12_financing_type=st.session_state.offer_12_finance_type,
        offer_12_home_inspection_check=st.session_state.offer_12_home_inspection_value,
        offer_12_home_inspection_days=st.session_state.offer_12_home_inspection_days_string,
        offer_12_radon_inspection_check=st.session_state.offer_12_radon_inspection_value,
        offer_12_radon_inspection_days=st.session_state.offer_12_radon_inspection_days_string,
        offer_12_septic_inspection_check=st.session_state.offer_12_septic_inspection_value,
        offer_12_septic_inspection_days=st.session_state.offer_12_septic_inspection_days_string,
        offer_12_well_inspection_check=st.session_state.offer_12_well_inspection_value,
        offer_12_well_inspection_days=st.session_state.offer_12_well_inspection_days_string,
        offer_12_finance_contingency_check=st.session_state.offer_12_financing_contingency_value,
        offer_12_finance_contingency_days=st.session_state.offer_12_financing_contingency_days_string,
        offer_12_appraisal_contingency_check=st.session_state.offer_12_appraisal_contingency_value,
        offer_12_appraisal_contingency_days=st.session_state.offer_12_appraisal_contingency_days_string,
        offer_12_home_sale_contingency_check=st.session_state.offer_12_home_sale_contingency_value,
        offer_12_home_sale_contingency_days=st.session_state.offer_12_home_sale_contingency_days_string,
        offer_12_pre_occupancy_date=st.session_state.offer_12_pre_occupancy_date,
        offer_12_post_occupancy_date=st.session_state.offer_12_post_occupancy_date,
        offer_12_closing_cost_subsidy_amt=st.session_state.offer_12_closing_subsidy_amt,
        offer_12_pre_occupancy_credit_amt=st.session_state.offer_12_pre_occupancy_credit_to_seller_amt,
        offer_12_post_occupancy_cost_amt=st.session_state.offer_12_post_occupancy_cost_to_seller_amt,
    )

    st.download_button(
        label='Download Offer Comparison Form',
        data=offer_comparison_form,
        mime='xlsx',
        file_name=f"offer_comparison_form_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    )

if __name__ == '__main__':
    main()
