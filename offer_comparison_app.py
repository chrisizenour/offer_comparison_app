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
                This application is built into separate data input forms\n
                To ensure the form is populating correctly, open the Common Data form first to check to see that the 'selling company compensation %' and 'listing_company_compensation %' sliders are at the 2.5% position\n
                If the sliders are not initialized at these positions, refresh the entire webpage, re-enter the password and again check the slider positions\n
                Sometimes the webpage doesn't load correctly, so a refresh is required\n
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
                - After all data has been updated/entered and then submitted in their respective forms, the MS Excel comparison form can be created
                - Press the 'Download Offer Comparison Form' button to generate and download the MS Excel workbook
                '''
            )

    if 'preparer' not in st.session_state:
        st.session_state['preparer'] = ''
        st.session_state['prep_date'] = date.today()
        # st.session_state['update_prep_date'] = date.today()
        st.session_state['offer_qty'] = 1

        st.session_state['seller_name'] = ''
        st.session_state['address'] = ''
        st.session_state['list_price'] = 0
        st.session_state['payoff_amt_first_trust'] = 0
        st.session_state['payoff_amt_second_trust'] = 0
        st.session_state['annual_tax_amt'] = 0
        st.session_state['prorated_annual_tax_amt'] = 0.0
        st.session_state['annual_hoa_condo_fee_amt'] = 0
        st.session_state['update_annual_hoa_condo_fee_amt'] = 0
        st.session_state['prorated_annual_hoa_condo_fee_amt'] = 0.0

        st.session_state['update_listing_company_pct'] = 2.5
        st.session_state['listing_company_pct'] = 0.025
        st.session_state['update_selling_company_pct'] = 2.5
        st.session_state['selling_company_pct'] = 0.025
        st.session_state['processing_fee'] = 0
        st.session_state['settlement_fee'] = 450
        st.session_state['deed_preparation_fee'] = 150
        st.session_state['lien_trust_release_fee'] = 100
        st.session_state['lien_trust_release_qty'] = 1
        st.session_state['recording_release_fee'] = 38
        st.session_state['recording_release_qty'] = 1
        st.session_state['update_grantors_tax_pct'] = 0.1
        st.session_state['grantors_tax_pct'] = 0.001
        st.session_state['update_congestion_tax_pct'] = 0.2
        st.session_state['congestion_tax_pct'] = 0.002
        st.session_state['pest_inspection_fee'] = 50
        st.session_state['poa_condo_disclosure_fee'] = 350

        st.session_state['offer_1_name'] = 'Offer 1'
        st.session_state['offer_1_settlement_date'] = date.today()
        # st.session_state['update_offer_1_settlement_date'] = date.today()
        st.session_state['offer_1_settlement_company'] = ''
        st.session_state['offer_1_amt'] = 0
        st.session_state['offer_1_emd_amt'] = 0
        st.session_state['offer_1_finance_type'] = 'Select Financing Type'
        st.session_state['offer_1_down_pmt_pct'] = 0.0
        st.session_state['offer_1_closing_subsidy_radio'] = 'Percent of Offer Amt (%)'
        st.session_state['offer_1_update_closing_subsidy_pct'] = 0.0
        st.session_state['offer_1_closing_subsidy_flat_amt'] = 0
        st.session_state['offer_1_closing_subsidy_amt'] = 0.0
        # st.session_state['offer_1_contingencies_waived'] = ''
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

        st.session_state['offer_2_name'] = 'Offer 2'
        st.session_state['offer_2_settlement_date'] = date.today()
        # st.session_state['update_offer_2_settlement_date'] = date.today()
        st.session_state['offer_2_settlement_company'] = ''
        st.session_state['offer_2_amt'] = 0
        st.session_state['offer_2_emd_amt'] = 0
        st.session_state['offer_2_finance_type'] = 'Select Financing Type'
        st.session_state['offer_2_down_pmt_pct'] = 0.0
        st.session_state['offer_2_closing_subsidy_radio'] = 'Percent of Offer Amt (%)'
        st.session_state['offer_2_update_closing_subsidy_pct'] = 0.0
        st.session_state['offer_2_closing_subsidy_flat_amt'] = 0
        st.session_state['offer_2_closing_subsidy_amt'] = 0.0
        # st.session_state['offer_2_contingencies_waived'] = ''
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

        st.session_state['offer_3_name'] = 'Offer 3'
        st.session_state['offer_3_settlement_date'] = date.today()
        # st.session_state['update_offer_3_settlement_date'] = date.today()
        st.session_state['offer_3_settlement_company'] = ''
        st.session_state['offer_3_amt'] = 0
        st.session_state['offer_3_emd_amt'] = 0
        st.session_state['offer_3_finance_type'] = 'Select Financing Type'
        st.session_state['offer_3_down_pmt_pct'] = 0.0
        st.session_state['offer_3_closing_subsidy_radio'] = 'Percent of Offer Amt (%)'
        st.session_state['offer_3_update_closing_subsidy_pct'] = 0.0
        st.session_state['offer_3_closing_subsidy_flat_amt'] = 0
        st.session_state['offer_3_closing_subsidy_amt'] = 0.0
        # st.session_state['offer_3_contingencies_waived'] = ''
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


    contingencies = ['Home Inspection', 'Financing', 'Appraisal', 'Pest Inspection']
    financing_types = ['Select Financing Type', 'Cash', 'Conventional', 'FHA', 'VA', 'USDA', 'Other']


    def update_intro_info_form():
        st.session_state.prep_date = st.session_state.update_prep_date


    def update_property_info_form():
        st.session_state.prorated_annual_tax_amt = st.session_state.annual_tax_amt / 12 * 3
        st.session_state.prorated_annual_hoa_condo_fee_amt = st.session_state.update_annual_hoa_condo_fee_amt / 12 * 3


    def update_common_info_form():
        st.session_state.listing_company_pct = st.session_state.update_listing_company_pct / 100
        st.session_state.selling_company_pct = st.session_state.update_selling_company_pct / 100
        st.session_state.grantors_tax_pct = st.session_state.update_grantors_tax_pct / 100
        st.session_state.congestion_tax_pct = st.session_state.update_congestion_tax_pct / 100


    def days_int_to_string(x):
        if x == 0:
            string_value = ''
        return string_value

    def days_int_to_string_false(x):
        pass


    def update_offer_1_info_form():
        st.session_state.offer_1_name = st.session_state.offer_1_name
        st.session_state.offer_1_settlement_date = st.session_state.update_offer_1_settlement_date
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
        st.session_state.offer_2_name = st.session_state.offer_2_name
        st.session_state.offer_2_settlement_date = st.session_state.update_offer_2_settlement_date
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
        st.session_state.offer_3_name = st.session_state.offer_3_name
        st.session_state.offer_3_settlement_date = st.session_state.update_offer_3_settlement_date
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


    with intro_info_container:
        with st.expander('Introduction Data Form'):
            with st.form(key='intro_info_form'):
                st.markdown('##### **Enter Top-Level Form Data**')
                intro_info_col1, intro_info_col2 = st.columns(2)
                with intro_info_col1:
                    st.text_input('Enter the name of the agent preparing this offer comparison', key='preparer')
                    st.date_input('Enter the date that this offer comparison was created', key='update_prep_date')
                with intro_info_col2:
                    st.slider('Number of Offers Being Compared', 1, 10, step=1, key='offer_qty')
                intro_info_submit = st.form_submit_button('Submit Information', on_click=update_intro_info_form)

    with property_container:
        with st.expander('Property Data Form'):
            with st.form(key='property_info_form'):
                st.markdown('##### **Enter Property-Related Data**')
                property_info_col1, property_info_col2 = st.columns(2)
                with property_info_col1:
                    st.text_input('Name of the Seller(s)', key='seller_name')
                    st.text_input('Property\'s Street Address', key='address')
                    st.slider('Property\'s List Price ($)', 0, 1500000, step=1000, key='list_price')
                with property_info_col2:
                    st.slider('Estimated Payoff - First Trust ($)', 0, 1000000, step=1000, key='payoff_amt_first_trust')
                    st.slider('Estimated Payoff - Second Trust ($)', 0, 1000000, step=1000, key='payoff_amt_second_trust')
                    st.slider('Estimated Annual Tax Amount ($)', 0, 25000, step=1, key='annual_tax_amt')
                    st.slider('Estimated Annual HOA / Condo Fee Amount ($)', 0, 10000, step=1, key='update_annual_hoa_condo_fee_amt')
                property_info_submit = st.form_submit_button('Submit Property Information', on_click=update_property_info_form)

    with common_container:
        with st.expander('Common Data Form'):
            with st.form(key='common_info_form'):
                st.markdown('##### **Enter Information Common To All Offers**')
                brokerage_col, closing_cost_col, misc_col = st.columns(3)
                with brokerage_col:
                    st.markdown('###### **Brokerage Cost Data**')
                    st.slider('Listing Company Compensation (%)', 0.0, 6.0, step=0.01, format='%.2f', key='update_listing_company_pct')
                    st.slider('Selling Company Compensation (%)', 0.0, 6.0, step=0.01, format='%.2f', key='update_selling_company_pct')
                    st.slider('Processing Fee ($)', 0, 10000, step=1, key='processing_fee')
                with closing_cost_col:
                    st.markdown('###### **Closing Cost Data**')
                    st.slider('Settlement Fee Amount ($)', 0, 1000, step=1, key='settlement_fee')
                    st.slider('Deed Preparation Fee Amount ($)', 0, 1000, step=1, key='deed_preparation_fee')
                    st.slider('Release of Liens / Trusts Fee Amount ($)', 0, 1000, step=1, key='lien_trust_release_fee')
                    st.slider('Quantity of Liens / Trusts to be Released', 0, 10, step=1, key='lien_trust_release_qty')
                with misc_col:
                    st.markdown('###### **Miscellaneous Cost Data**')
                    st.slider('Recording Release Fee Amount ($)', 0, 250, step=1, key='recording_release_fee')
                    st.slider('Quantity of Recording Releases', 0, 10, step=1, key='recording_release_qty')
                    st.slider('Grantor\'s Tax Pct (%)', 0.0, 1.0, step=0.01, format='%.2f', key='update_grantors_tax_pct')
                    st.slider('Congestion Tax Pct (%)', 0.0, 1.0, step=0.01, format='%.2f', key='update_congestion_tax_pct')
                    st.slider('Pest Inspection Fee Amount ($)', 0, 100, step=1, key='pest_inspection_fee')
                    st.slider('Power of Attorney / Condo Disclosure Fee Amount ($)', 0, 500, step=1, key='poa_condo_disclosure_fee')
                common_info_submit = st.form_submit_button('Submit Common Information', on_click=update_common_info_form)

    with offer_1_container:
        with st.expander('Offer 1 Form'):
            with st.form(key='offer_1_info_form'):
                st.markdown('##### **Enter Offer 1\'s Information**')
                offer_1_col1, offer_1_col2 = st.columns(2)
                with offer_1_col1:
                    st.text_input('Name of Offer', key='offer_1_name')
                    st.date_input('Settlement Date', key='update_offer_1_settlement_date')
                    st.text_input('Settlement Company', key='offer_1_settlement_company')
                    st.slider('Offer Amount ($)', 0, 1500000, step=1000, key='offer_1_amt')
                    st.slider('EMD Amount ($)', 0, 50000, step=100, key='offer_1_emd_amt')
                with offer_1_col2:
                    st.selectbox('Financing Type', financing_types, key='offer_1_finance_type')
                    st.slider('Down Payment Pct (%)', 0.0, 100.0, step=0.01, key='offer_1_down_pmt_pct')
                    st.radio('Closing Cost Subsidy Radio', ['Percent of Offer Amt (%)', 'Flat $ Amount'], key='offer_1_closing_subsidy_radio')
                    st.slider('Closing Cost Subsidy of (%):', 0.0, 100.0, step=0.01, key='offer_1_update_closing_subsidy_pct')
                    st.slider('Closing Cost Subsidy of ($):', 0, 100000, step=50, key='offer_1_closing_subsidy_flat_amt')
                st.write('---')
                st.write('Contingencies and Clauses of the Offer')
                # st.text_input('Contingencies Waived', key='offer_1_contingencies_waved')
                offer_1_cont_col1, offer_1_cont_col2 = st.columns(2)
                with offer_1_cont_col1:
                    st.checkbox('Home Inspection', key='offer_1_home_inspection_check')
                    st.slider('Home Inspection Days', 0, 45, step=1, key='offer_1_home_inspection_days')
                    st.checkbox('Radon Inspection', key='offer_1_radon_inspection_check')
                    st.slider('Radon Inspection Days', 0, 45, step=1, key='offer_1_radon_inspection_days')
                    st.checkbox('Septic Inspection', key='offer_1_septic_inspection_check')
                    st.slider('Septic Inspection Days', 0, 45, step=1, key='offer_1_septic_inspection_days')
                    st.checkbox('Well Inspection', key='offer_1_well_inspection_check')
                    st.slider('Well Inspection Days', 0, 45, step=1, key='offer_1_well_inspection_days')
                with offer_1_cont_col2:
                    st.checkbox('Financing Contingency', key='offer_1_financing_contingency_check')
                    st.slider('Financing Contingency Days', 0, 45, step=1, key='offer_1_financing_contingency_days')
                    st.checkbox('Appraisal Contingency', key='offer_1_appraisal_contingency_check')
                    st.slider('Appraisal Contingency Days', 0, 45, step=1, key='offer_1_appraisal_contingency_days')
                    st.checkbox('Home Sale Contingency', key='offer_1_home_sale_contingency_check')
                    st.slider('Home Sale Contingency Days', 0, 45, step=1, key='offer_1_home_sale_contingency_days')
                st.write('---')
                st.write('Pre Occupancy')
                offer_1_pre_occupancy_col1, offer_1_pre_occupancy_col2 = st.columns(2)
                with offer_1_pre_occupancy_col1:
                    st.checkbox('Pre Occupancy Request', key='offer_1_pre_occupancy_request')
                    st.slider('Pre Occupancy Credit to Seller ($)', 0, 25000, step=1, key='offer_1_pre_occupancy_credit_to_seller_amt')
                with offer_1_pre_occupancy_col2:
                    st.date_input('Pre Occupancy Date', key='offer_1_update_pre_occupancy_date')
                st.write('---')
                st.write('Post Occupancy')
                offer_1_post_occupancy_col1, offer_1_post_occupancy_col2 = st.columns(2)
                with offer_1_post_occupancy_col1:
                    st.checkbox('Post Occupancy Request', key='offer_1_post_occupancy_request')
                    st.slider('Post Occupancy Cost to Seller ($)', 0, 25000, step=1, key='offer_1_post_occupancy_cost_to_seller_amt')
                with offer_1_post_occupancy_col2:
                    st.date_input('Post Occupancy Date', key='offer_1_update_post_occupancy_date')
                offer_1_submit = st.form_submit_button('Submit Offer 1\'s Information', on_click=update_offer_1_info_form)


    with offer_2_container:
        with st.expander('Offer 2 Form'):
            with st.form(key='offer_2_info_form'):
                st.markdown('##### **Enter Offer 2\'s Information**')
                offer_2_col1, offer_2_col2 = st.columns(2)
                with offer_2_col1:
                    st.text_input('Name of Offer', key='offer_2_name')
                    st.date_input('Settlement Date', key='update_offer_2_settlement_date')
                    st.text_input('Settlement Company', key='offer_2_settlement_company')
                    st.slider('Offer Amount ($)', 0, 1500000, step=1000, key='offer_2_amt')
                    st.slider('EMD Amount ($)', 0, 50000, step=100, key='offer_2_emd_amt')
                with offer_2_col2:
                    st.selectbox('Financing Type', financing_types, key='offer_2_finance_type')
                    st.slider('Down Payment Pct (%)', 0.0, 100.0, step=0.01, key='offer_2_down_pmt_pct')
                    st.radio('Closing Cost Subsidy Radio', ['Percent of Offer Amt (%)', 'Flat $ Amount'], key='offer_2_closing_subsidy_radio')
                    st.slider('Closing Cost Subsidy of (%):', 0.0, 100.0, step=0.01, key='offer_2_update_closing_subsidy_pct')
                    st.slider('Closing Cost Subsidy of ($):', 0, 100000, step=50, key='offer_2_closing_subsidy_flat_amt')
                st.write('---')
                st.write('Contingencies and Clauses of the Offer')
                # st.text_input('Contingencies Waived', key='offer_2_contingencies_waved')
                offer_2_cont_col1, offer_2_cont_col2 = st.columns(2)
                with offer_2_cont_col1:
                    st.checkbox('Home Inspection', key='offer_2_home_inspection_check')
                    st.slider('Home Inspection Days', 0, 45, step=1, key='offer_2_home_inspection_days')
                    st.checkbox('Radon Inspection', key='offer_2_radon_inspection_check')
                    st.slider('Radon Inspection Days', 0, 45, step=1, key='offer_2_radon_inspection_days')
                    st.checkbox('Septic Inspection', key='offer_2_septic_inspection_check')
                    st.slider('Septic Inspection Days', 0, 45, step=1, key='offer_2_septic_inspection_days')
                    st.checkbox('Well Inspection', key='offer_2_well_inspection_check')
                    st.slider('Well Inspection Days', 0, 45, step=1, key='offer_2_well_inspection_days')
                with offer_2_cont_col2:
                    st.checkbox('Financing Contingency', key='offer_2_financing_contingency_check')
                    st.slider('Financing Contingency Days', 0, 45, step=1, key='offer_2_financing_contingency_days')
                    st.checkbox('Appraisal Contingency', key='offer_2_appraisal_contingency_check')
                    st.slider('Appraisal Contingency Days', 0, 45, step=1, key='offer_2_appraisal_contingency_days')
                    st.checkbox('Home Sale Contingency', key='offer_2_home_sale_contingency_check')
                    st.slider('Home Sale Contingency Days', 0, 45, step=1, key='offer_2_home_sale_contingency_days')
                st.write('---')
                st.write('Pre Occupancy')
                offer_2_pre_occupancy_col1, offer_2_pre_occupancy_col2 = st.columns(2)
                with offer_2_pre_occupancy_col1:
                    st.checkbox('Pre Occupancy Request', key='offer_2_pre_occupancy_request')
                    st.slider('Pre Occupancy Credit to Seller ($)', 0, 25000, step=1,
                              key='offer_2_pre_occupancy_credit_to_seller_amt')
                with offer_2_pre_occupancy_col2:
                    st.date_input('Pre Occupancy Date', key='offer_2_update_pre_occupancy_date')
                st.write('---')
                st.write('Post Occupancy')
                offer_2_post_occupancy_col1, offer_2_post_occupancy_col2 = st.columns(2)
                with offer_2_post_occupancy_col1:
                    st.checkbox('Post Occupancy Request', key='offer_2_post_occupancy_request')
                    st.slider('Post Occupancy Cost to Seller ($)', 0, 25000, step=1, key='offer_2_post_occupancy_cost_to_seller_amt')
                with offer_2_post_occupancy_col2:
                    st.date_input('Post Occupancy Date', key='offer_2_update_post_occupancy_date')
                offer_2_submit = st.form_submit_button('Submit Offer 2\'s Information', on_click=update_offer_2_info_form)

    with offer_3_container:
        with st.expander('Offer 3 Form'):
            with st.form(key='offer_3_info_form'):
                st.markdown('##### **Enter Offer 3\'s Information**')
                offer_3_col1, offer_3_col2 = st.columns(2)
                with offer_3_col1:
                    st.text_input('Name of Offer', key='offer_3_name')
                    st.date_input('Settlement Date', key='update_offer_3_settlement_date')
                    st.text_input('Settlement Company', key='offer_3_settlement_company')
                    st.slider('Offer Amount ($)', 0, 1500000, step=1000, key='offer_3_amt')
                    st.slider('EMD Amount ($)', 0, 50000, step=100, key='offer_3_emd_amt')
                with offer_3_col2:
                    st.selectbox('Financing Type', financing_types, key='offer_3_finance_type')
                    st.slider('Down Payment Pct (%)', 0.0, 100.0, step=0.01, key='offer_3_down_pmt_pct')
                    st.radio('Closing Cost Subsidy Radio', ['Percent of Offer Amt (%)', 'Flat $ Amount'], key='offer_3_closing_subsidy_radio')
                    st.slider('Closing Cost Subsidy of (%):', 0.0, 100.0, step=0.01, key='offer_3_update_closing_subsidy_pct')
                    st.slider('Closing Cost Subsidy of ($):', 0, 100000, step=50, key='offer_3_closing_subsidy_flat_amt')
                st.write('---')
                st.write('Contingencies and Clauses of the Offer')
                # st.text_input('Contingencies Waived', key='offer_3_contingencies_waved')
                offer_3_cont_col1, offer_3_cont_col2 = st.columns(2)
                with offer_3_cont_col1:
                    st.checkbox('Home Inspection', key='offer_3_home_inspection_check')
                    st.slider('Home Inspection Days', 0, 45, step=1, key='offer_3_home_inspection_days')
                    st.checkbox('Radon Inspection', key='offer_3_radon_inspection_check')
                    st.slider('Radon Inspection Days', 0, 45, step=1, key='offer_3_radon_inspection_days')
                    st.checkbox('Septic Inspection', key='offer_3_septic_inspection_check')
                    st.slider('Septic Inspection Days', 0, 45, step=1, key='offer_3_septic_inspection_days')
                    st.checkbox('Well Inspection', key='offer_3_well_inspection_check')
                    st.slider('Well Inspection Days', 0, 45, step=1, key='offer_3_well_inspection_days')
                with offer_3_cont_col2:
                    st.checkbox('Financing Contingency', key='offer_3_financing_contingency_check')
                    st.slider('Financing Contingency Days', 0, 45, step=1, key='offer_3_financing_contingency_days')
                    st.checkbox('Appraisal Contingency', key='offer_3_appraisal_contingency_check')
                    st.slider('Appraisal Contingency Days', 0, 45, step=1, key='offer_3_appraisal_contingency_days')
                    st.checkbox('Home Sale Contingency', key='offer_3_home_sale_contingency_check')
                    st.slider('Home Sale Contingency Days', 0, 45, step=1, key='offer_3_home_sale_contingency_days')
                st.write('---')
                st.write('Pre Occupancy')
                offer_3_pre_occupancy_col1, offer_3_pre_occupancy_col2 = st.columns(2)
                with offer_3_pre_occupancy_col1:
                    st.checkbox('Pre Occupancy Request', key='offer_3_pre_occupancy_request')
                    st.slider('Pre Occupancy Credit to Seller ($)', 0, 25000, step=1,
                              key='offer_3_pre_occupancy_credit_to_seller_amt')
                with offer_3_pre_occupancy_col2:
                    st.date_input('Pre Occupancy Date', key='offer_3_update_pre_occupancy_date')
                st.write('---')
                st.write('Post Occupancy')
                offer_3_post_occupancy_col1, offer_3_post_occupancy_col2 = st.columns(2)
                with offer_3_post_occupancy_col1:
                    st.checkbox('Post Occupancy Request', key='offer_3_post_occupancy_request')
                    st.slider('Post Occupancy Cost to Seller ($)', 0, 25000, step=1, key='offer_3_post_occupancy_cost_to_seller_amt')
                with offer_3_post_occupancy_col2:
                    st.date_input('Post Occupancy Date', key='offer_3_update_post_occupancy_date')
                offer_3_submit = st.form_submit_button('Submit Offer 3\'s Information', on_click=update_offer_3_info_form)

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
    )

    st.download_button(
        label='Download Offer Comparison Form',
        data=offer_comparison_form,
        mime='xlsx',
        file_name=f"offer_comparison_form_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    )

if __name__ == '__main__':
    main()
