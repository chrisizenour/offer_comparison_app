import pandas as pd
from io import BytesIO
from tempfile import NamedTemporaryFile
from datetime import datetime
import openpyxl
from openpyxl.workbook import Workbook
from openpyxl.writer.excel import save_virtual_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.properties import WorksheetProperties, PageSetupProperties
from openpyxl.worksheet.dimensions import ColumnDimension, SheetDimension, SheetFormatProperties
from openpyxl.worksheet.pagebreak import Break, PageBreak
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.worksheet.page import PageMargins
from openpyxl.styles import PatternFill, NamedStyle, Border, Side, Alignment, Protection, Font, Color, colors, numbers, DEFAULT_FONT
from openpyxl import drawing

def comparison_inputs_to_excel(
        agent,
        date,
        offer_qty,
        seller_name,
        seller_address,
        list_price,
        first_trust,
        second_trust,
        annual_taxes,
        prorated_taxes,
        annual_hoa_condo_fees,
        prorated_hoa_condo_fees,
        listing_company_pct,
        selling_company_pct,
        processing_fee,
        settlement_fee,
        deed_preparation_fee,
        lien_trust_release_fee,
        lien_trust_release_qty,
        recording_fee,
        recording_trusts_liens_qty,
        grantors_tax_pct,
        congestion_tax_pct,
        pest_inspection_fee,
        poa_condo_disclosure_fee,
        offer_1_name,
        offer_1_amt,
        offer_1_down_pmt_pct,
        offer_1_settlement_date,
        offer_1_settlement_company,
        offer_1_emd_amt,
        offer_1_financing_type,
        offer_1_contingencies_waived,
        offer_1_post_occupancy_date,
        offer_1_closing_cost_subsidy_amt,
        offer_1_post_occupancy_cost_amt,
        offer_2_name,
        offer_2_amt,
        offer_2_down_pmt_pct,
        offer_2_settlement_date,
        offer_2_settlement_company,
        offer_2_emd_amt,
        offer_2_financing_type,
        offer_2_contingencies_waived,
        offer_2_post_occupancy_date,
        offer_2_closing_cost_subsidy_amt,
        offer_2_post_occupancy_cost_amt,
        offer_3_name,
        offer_3_amt,
        offer_3_down_pmt_pct,
        offer_3_settlement_date,
        offer_3_settlement_company,
        offer_3_emd_amt,
        offer_3_financing_type,
        offer_3_contingencies_waived,
        offer_3_post_occupancy_date,
        offer_3_closing_cost_subsidy_amt,
        offer_3_post_occupancy_cost_amt,):

    wb = Workbook()
    dest_filename = f"offer_comparison_form_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

    ws1 = wb.active
    ws1.title = 'offer_comparison'
    ws1.print_area = 'B2:H56'
    ws1.set_printer_settings(paper_size=1, orientation='landscape')
    # ws1.print_options.horizontalCentered = True
    # ws1.print_options.verticalCentered = True
    ws1.sheet_properties.pageSetUpPr.fitToPage = True

    white_fill = '00FFFFFF'
    yellow_fill = '00FFFF00'
    black_fill = '00000000'
    font_size = 12
    thick = Side(border_style='thick')
    thin = Side(border_style='thin')
    hair = Side(border_style='hair')
    DEFAULT_FONT.size = font_size
    ws1.column_dimensions['B'].width = 1.5
    ws1.column_dimensions['C'].width = 39.0
    ws1.column_dimensions['D'].width = 39.0
    ws1.column_dimensions['E'].width = 12.0
    ws1.column_dimensions['F'].width = 12.0
    ws1.column_dimensions['G'].width = 12.0
    ws1.column_dimensions['H'].width = 1.5

    acct_fmt = '_($* #,##0_);[Red]_($* (#,##0);_($* "-"??_);_(@_)'
    pct_fmt = '0.00%'
    # date_fmt = NamedStyle(name='date', number_format='DD/MM/YYYY')

    for row in ws1.iter_rows(min_row=1, max_row=70, min_col=1, max_col=15):
        for cell in row:
            cell.fill = PatternFill(start_color=white_fill, end_color=white_fill, fill_type='solid')

    # Build Black Border
    ws1.merge_cells('A1:I1')
    top_left_border_one = ws1['A1']
    top_left_border_one.fill = PatternFill('solid', fgColor=black_fill)

    ws1.merge_cells('A2:A57')
    top_left_border_two = ws1['A2']
    top_left_border_two.fill = PatternFill('solid', fgColor=black_fill)

    ws1.merge_cells('B57:I57')
    top_left_border_three = ws1['B57']
    top_left_border_three.fill = PatternFill('solid', fgColor=black_fill)

    ws1.merge_cells('I2:I56')
    top_left_border_four = ws1['I2']
    top_left_border_four.fill = PatternFill('solid', fgColor=black_fill)

    # Build Header
    ws1.merge_cells('C3:G3')
    top_left_cell_one = ws1['C3']
    top_left_cell_one.value = 'Seller\'s Total Net Proceeds For Different Offers'
    top_left_cell_one.font = Font(bold=True)
    top_left_cell_one.alignment = Alignment(horizontal='center')

    ws1.merge_cells('C4:G4')
    top_left_cell_two = ws1['C4']
    top_left_cell_two.value = f'{seller_name} - {seller_address}'
    top_left_cell_two.font = Font(bold=True)
    top_left_cell_two.alignment = Alignment(horizontal='center')

    ws1.merge_cells('C5:G5')
    top_left_cell_three = ws1['C5']
    top_left_cell_three.value = f'Date Prepared: {date}'
    top_left_cell_three.font = Font(bold=True)
    top_left_cell_three.alignment = Alignment(horizontal='center')

    ws1.merge_cells('C6:G6')
    top_left_cell_four = ws1['C6']
    top_left_cell_four.value = f'List Price: ${list_price:,.2f}'
    top_left_cell_four.font = Font(bold=True)
    top_left_cell_four.alignment = Alignment(horizontal='center')

    e7 = ws1['E7']
    e7.value = offer_1_name
    e7.font = Font(bold=True)
    e7.border = Border(bottom=thin)
    e7.alignment = Alignment(horizontal='center', wrap_text=True)

    f7 = ws1['F7']
    f7.value = offer_2_name
    f7.font = Font(bold=True)
    f7.border = Border(bottom=thin)
    f7.alignment = Alignment(horizontal='center', wrap_text=True)

    g7 = ws1['G7']
    g7.value = offer_3_name
    g7.font = Font(bold=True)
    g7.border = Border(bottom=thin)
    g7.alignment = Alignment(horizontal='center', wrap_text=True)

    d8 = ws1['D8']
    d8.value = 'Offer Amt. ($)'
    d8.font = Font(bold=True)
    d8.alignment = Alignment(horizontal='right', vertical='center')

    e8 = ws1['E8']
    e8.value = offer_1_amt
    e8.font = Font(bold=True)
    e8.number_format = acct_fmt
    e8.border = Border(top=thin, bottom=hair, left=thin, right=thin)

    f8 = ws1['F8']
    f8.value = offer_2_amt
    f8.font = Font(bold=True)
    f8.number_format = acct_fmt
    f8.border = Border(top=thin, bottom=hair, left=thin, right=thin)

    g8 = ws1['G8']
    g8.value = offer_3_amt
    g8.font = Font(bold=True)
    g8.number_format = acct_fmt
    g8.border = Border(top=thin, bottom=hair, right=thin, left=thin)

    d9 = ws1['D9']
    d9.value = 'Down Pmt (%)'
    d9.font = Font(bold=True)
    d9.alignment = Alignment(horizontal='right', vertical='center')

    e9 = ws1['E9']
    e9.value = offer_1_down_pmt_pct
    e9.font = Font(bold=True)
    e9.number_format = pct_fmt
    e9.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    f9 = ws1['F9']
    f9.value = offer_2_down_pmt_pct
    f9.font = Font(bold=True)
    f9.number_format = pct_fmt
    f9.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    g9 = ws1['G9']
    g9.value = offer_3_down_pmt_pct
    g9.font = Font(bold=True)
    g9.number_format = pct_fmt
    g9.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    d10 = ws1['D10']
    d10.value = 'Settlement Date'
    d10.font = Font(bold=True)
    d10.alignment = Alignment(horizontal='right', vertical='center')

    e10 = ws1['E10']
    e10.value = offer_1_settlement_date
    e10.font = Font(bold=True)
    e10.alignment = Alignment(horizontal='right')
    e10.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    f10 = ws1['F10']
    f10.value = offer_2_settlement_date
    f10.font = Font(bold=True)
    f10.alignment = Alignment(horizontal='right')
    f10.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    g10 = ws1['G10']
    g10.value = offer_3_settlement_date
    g10.font = Font(bold=True)
    g10.alignment = Alignment(horizontal='right')
    g10.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    d11 = ws1['D11']
    d11.value = 'Settlement Company'
    d11.font = Font(bold=True)
    d11.alignment = Alignment(horizontal='right', vertical='center')

    e11 = ws1['E11']
    e11.value = offer_1_settlement_company
    e11.font = Font(bold=True)
    e11.alignment = Alignment(horizontal='center', vertical='top', wrap_text=True)
    e11.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    f11 = ws1['F11']
    f11.value = offer_2_settlement_company
    f11.font = Font(bold=True)
    f11.alignment = Alignment(horizontal='center', vertical='top', wrap_text=True)
    f11.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    g11 = ws1['G11']
    g11.value = offer_3_settlement_company
    g11.font = Font(bold=True)
    g11.alignment = Alignment(horizontal='center', vertical='top', wrap_text=True)
    g11.border = Border(top=hair, bottom=hair, right=thin, left=thin)

    d12 = ws1['D12']
    d12.value = 'EMD Amt. ($)'
    d12.font = Font(bold=True)
    d12.alignment = Alignment(horizontal='right', vertical='center')

    e12 = ws1['E12']
    e12.value = offer_1_emd_amt
    e12.font = Font(bold=True)
    e12.number_format = acct_fmt
    e12.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    f12 = ws1['F12']
    f12.value = offer_2_emd_amt
    f12.font = Font(bold=True)
    f12.number_format = acct_fmt
    f12.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    g12 = ws1['G12']
    g12.value = offer_3_emd_amt
    g12.font = Font(bold=True)
    g12.number_format = acct_fmt
    g12.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    d13 = ws1['D13']
    d13.value = 'Financing Type'
    d13.font = Font(bold=True)
    d13.alignment = Alignment(horizontal='right', vertical='center')

    e13 = ws1['E13']
    e13.value = offer_1_financing_type
    e13.font = Font(bold=True)
    e13.alignment = Alignment(horizontal='center')
    e13.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    f13 = ws1['F13']
    f13.value = offer_2_financing_type
    f13.font = Font(bold=True)
    f13.alignment = Alignment(horizontal='center')
    f13.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    g13 = ws1['G13']
    g13.value = offer_3_financing_type
    g13.font = Font(bold=True)
    g13.alignment = Alignment(horizontal='center')
    g13.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    d14 = ws1['D14']
    d14.value = 'Contingencies Waived'
    d14.font = Font(bold=True)
    d14.alignment = Alignment(horizontal='right', vertical='center')

    e14 = ws1['E14']
    e14.value = offer_1_contingencies_waived
    e14.font = Font(bold=True)
    e14.alignment = Alignment(horizontal='center', vertical='top', wrap_text=True)
    e14.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    f14 = ws1['F14']
    f14.value = offer_2_contingencies_waived
    f14.font = Font(bold=True)
    f14.alignment = Alignment(horizontal='center', vertical='top', wrap_text=True)
    f14.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    g14 = ws1['G14']
    g14.value = offer_3_contingencies_waived
    g14.font = Font(bold=True)
    g14.alignment = Alignment(horizontal='center', vertical='top', wrap_text = True)
    g14.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    d15 = ws1['D15']
    d15.value = 'Post Occupancy Thru Date'
    d15.font = Font(bold=True)
    d15.alignment = Alignment(horizontal='right', vertical='center')

    e15 = ws1['E15']
    e15.value = offer_1_post_occupancy_date
    e15.font = Font(bold=True)
    e15.alignment = Alignment(horizontal='right')
    e15.border = Border(top=hair, bottom=thin, left=thin, right=thin)

    f15 = ws1['F15']
    f15.value = offer_2_post_occupancy_date
    f15.font = Font(bold=True)
    f15.alignment = Alignment(horizontal='right')
    f15.border = Border(top=hair, bottom=thin, left=thin, right=thin)

    g15 = ws1['G15']
    g15.value = offer_3_post_occupancy_date
    g15.font = Font(bold=True)
    g15.alignment = Alignment(horizontal='right')
    g15.border = Border(top=hair, bottom=thin, left=thin, right=thin)

    # Housing Related Cost Table
    c17 = ws1['C17']
    c17.value = 'HOUSING-RELATED COSTS'
    c17.font = Font(bold=True)
    c17.border = Border(bottom=thin)

    d17 = ws1['D17']
    d17.value = 'Calculation Description'
    d17.font = Font(bold=True)
    d17.border = Border(bottom=thin)

    c18 = ws1['C18']
    c18.value = 'Estimated Payoff - 1st Trust'
    c18.border = Border(top=thin, bottom=hair, left=thin)

    d18 = ws1['D18']
    d18.value = 'Principal Balance of Loan'
    d18.border = Border(top=thin, bottom=hair)

    e18 = ws1['E18']
    e18.value = first_trust
    e18.number_format = acct_fmt
    e18.border = Border(top=thin, bottom=hair, left=thin, right=thin)

    f18 = ws1['F18']
    f18.value = first_trust
    f18.number_format = acct_fmt
    f18.border = Border(top=thin, bottom=hair, left=thin, right=thin)

    g18 = ws1['G18']
    g18.value = first_trust
    g18.number_format = acct_fmt
    g18.border = Border(top=thin, bottom=hair, left=thin, right=thin)

    c19 = ws1['C19']
    c19.value = 'Estimated Payoff - 2nd Trust'
    c19.border = Border(top=hair, bottom=hair, left=thin)

    d19 = ws1['D19']
    d19.value = 'Principal Balance of Loan'
    d19.border = Border(top=hair, bottom=hair)

    e19 = ws1['E19']
    e19.value = second_trust
    e19.number_format = acct_fmt
    e19.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    f19 = ws1['F19']
    f19.value = second_trust
    f19.number_format = acct_fmt
    f19.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    g19 = ws1['G19']
    g19.value = second_trust
    g19.number_format = acct_fmt
    g19.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    c20 = ws1['C20']
    c20.value = 'Purchaser Closing Cost / Contract'
    c20.border = Border(top=hair, bottom=hair, left=thin)

    d20 = ws1['D20']
    d20.value = 'Negotiated Into Contract'
    d20.border = Border(top=hair, bottom=hair)

    e20 = ws1['E20']
    e20.value = offer_1_closing_cost_subsidy_amt
    e20.number_format = acct_fmt
    e20.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    f20 = ws1['F20']
    f20.value = offer_2_closing_cost_subsidy_amt
    f20.number_format = acct_fmt
    f20.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    g20 = ws1['G20']
    g20.value = offer_3_closing_cost_subsidy_amt
    g20.number_format = acct_fmt
    g20.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    c21 = ws1['C21']
    c21.value = 'Prorated Taxes / Assessments'
    c21.border = Border(top=hair, bottom=hair, left=thin)

    d21 = ws1['D21']
    d21.value = '1 Year of Taxes Divided by 12 Multiplied by 3'
    d21.border = Border(top=hair, bottom=hair)

    e21 = ws1['E21']
    e21.value = prorated_taxes
    e21.number_format = acct_fmt
    e21.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    f21 = ws1['F21']
    f21.value = prorated_taxes
    f21.number_format = acct_fmt
    f21.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    g21 = ws1['G21']
    g21.value = prorated_taxes
    g21.number_format = acct_fmt
    g21.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    c22 = ws1['C22']
    c22.value = 'Prorated HOA / Condo Dues'
    c22.border = Border(top=hair, bottom=thin, left=thin)

    d22 = ws1['D22']
    d22.value = '1 Year of Dues Divided by 12 Multiplied by 3'
    d22.border = Border(top=hair, bottom=thin)

    e22 = ws1['E22']
    e22.value = prorated_hoa_condo_fees
    e22.number_format = acct_fmt
    e22.border = Border(top=hair, bottom=thin, left=thin, right=thin)

    f22 = ws1['F22']
    f22.value = prorated_hoa_condo_fees
    f22.number_format = acct_fmt
    f22.border = Border(top=hair, bottom=thin, left=thin, right=thin)

    g22 = ws1['G22']
    g22.value = prorated_hoa_condo_fees
    g22.number_format = acct_fmt
    g22.border = Border(top=hair, bottom=thin, left=thin, right=thin)

    e23 = ws1['E23']
    e23.value = '=SUM(E18:E22)'
    e23.font = Font(bold=True)
    e23.number_format = acct_fmt
    e23.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    f23 = ws1['F23']
    f23.value = '=SUM(F18:F22)'
    f23.font = Font(bold=True)
    f23.number_format = acct_fmt
    f23.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    g23 = ws1['G23']
    g23.value = '=SUM(G18:G22)'
    g23.font = Font(bold=True)
    g23.number_format = acct_fmt
    g23.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    # Broker and Financing Costs Table
    c24 = ws1['C24']
    c24.value = 'BROKERAGE & FINANCING COSTS'
    c24.font = Font(bold=True)
    c24.border = Border(bottom=thin)

    c25 = ws1['C25']
    c25.value = 'Listing Company Compensation'
    c25.border = Border(top=thin, bottom=hair, left=thin)

    d25 = ws1['D25']
    d25.value = '% from Listing Agreement * Offer Amount ($)'
    d25.border = Border(top=thin, bottom=hair)

    e25 = ws1['E25']
    e25.value = listing_company_pct * offer_1_amt
    e25.number_format = acct_fmt
    e25.border = Border(top=thin, bottom=hair, left=thin, right=thin)

    f25 = ws1['F25']
    f25.value = listing_company_pct * offer_2_amt
    f25.number_format = acct_fmt
    f25.border = Border(top=thin, bottom=hair, left=thin, right=thin)

    g25 = ws1['G25']
    g25.value = listing_company_pct * offer_3_amt
    g25.number_format = acct_fmt
    g25.border = Border(top=thin, bottom=hair, left=thin, right=thin)

    c26 = ws1['C26']
    c26.value = 'Selling Company Compensation'
    c26.border = Border(top=hair, bottom=hair, left=thin)

    d26 = ws1['D26']
    d26.value = '% from Listing Agreement * Offer Amount ($)'
    d26.border = Border(top=hair, bottom=hair)

    e26 = ws1['E26']
    e26.value = selling_company_pct * offer_1_amt
    e26.number_format = acct_fmt
    e26.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    f26 = ws1['F26']
    f26.value = selling_company_pct * offer_2_amt
    f26.number_format = acct_fmt
    f26.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    g26 = ws1['G26']
    g26.value = selling_company_pct * offer_3_amt
    g26.number_format = acct_fmt
    g26.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    c27 = ws1['C27']
    c27.value = 'Processing Fee'
    c27.border = Border(top=hair, bottom=thin, left=thin)

    d27 = ws1['D27']
    d27.border = Border(top=hair, bottom=thin)

    e27 = ws1['E27']
    e27.value = processing_fee
    e27.number_format = acct_fmt
    e27.border = Border(top=hair, bottom=thin, left=thin, right=thin)

    f27 = ws1['F27']
    f27.value = processing_fee
    f27.number_format = acct_fmt
    f27.border = Border(top=hair, bottom=thin, left=thin, right=thin)

    g27 = ws1['G27']
    g27.value = processing_fee
    g27.number_format = acct_fmt
    g27.border = Border(top=hair, bottom=thin, left=thin, right=thin)

    e28 = ws1['E28']
    e28.value = '=SUM(E25:E27)'
    e28.font = Font(bold=True)
    e28.number_format = acct_fmt
    e28.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    f28 = ws1['F28']
    f28.value = '=SUM(F25:F27)'
    f28.font = Font(bold=True)
    f28.number_format = acct_fmt
    f28.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    g28 = ws1['G28']
    g28.value = '=SUM(G25:G27)'
    g28.font = Font(bold=True)
    g28.number_format = acct_fmt
    g28.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    # Estimated Closing Costs Table
    c29 = ws1['C29']
    c29.value = 'ESTIMATED CLOSING COSTS'
    c29.font = Font(bold=True)
    c29.border = Border(bottom=thin)

    c30 = ws1['C30']
    c30.value = 'Settlement Fee'
    c30.border = Border(top=thin, bottom=hair, left=thin)

    d30 = ws1['D30']
    d30.value = 'Commonly Used Fee'
    d30.border = Border(top=thin, bottom=hair)

    e30 = ws1['E30']
    e30.value = settlement_fee
    e30.number_format = acct_fmt
    e30.border = Border(top=thin, bottom=hair, left=thin, right=thin)

    f30 = ws1['F30']
    f30.value = settlement_fee
    f30.number_format = acct_fmt
    f30.border = Border(top=thin, bottom=hair, left=thin, right=thin)

    g30 = ws1['G30']
    g30.value = settlement_fee
    g30.number_format = acct_fmt
    g30.border = Border(top=thin, bottom=hair, left=thin, right=thin)

    c31 = ws1['C31']
    c31.value = 'Deed Preparation'
    c31.border = Border(top=hair, bottom=hair, left=thin)

    d31 = ws1['D31']
    d31.value = 'Commonly Used Fee'
    d31.border = Border(top=hair, bottom=hair)

    e31 = ws1['E31']
    e31.value = deed_preparation_fee
    e31.number_format = acct_fmt
    e31.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    f31 = ws1['F31']
    f31.value = deed_preparation_fee
    f31.number_format = acct_fmt
    f31.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    g31 = ws1['G31']
    g31.value = deed_preparation_fee
    g31.number_format = acct_fmt
    g31.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    c32 = ws1['C32']
    c32.value = 'Release of Liens / Trusts'
    c32.border = Border(top=hair, bottom=thin, left=thin)

    d32 = ws1['D32']
    d32.value = 'Commonly Used Fee * Qty of Trusts or Liens'
    d32.border = Border(top=hair, bottom=thin)

    e32 = ws1['E32']
    e32.value = lien_trust_release_fee * lien_trust_release_qty
    e32.number_format = acct_fmt
    e32.border = Border(top=hair, bottom=thin, left=thin, right=thin)

    f32 = ws1['F32']
    f32.value = lien_trust_release_fee * lien_trust_release_qty
    f32.number_format = acct_fmt
    f32.border = Border(top=hair, bottom=thin, left=thin, right=thin)

    g32 = ws1['G32']
    g32.value = lien_trust_release_fee * lien_trust_release_qty
    g32.number_format = acct_fmt
    g32.border = Border(top=hair, bottom=thin, left=thin, right=thin)

    e33 = ws1['E33']
    e33.value = '=SUM(E30:E32)'
    e33.font = Font(bold=True)
    e33.number_format = acct_fmt
    e33.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    f33 = ws1['F33']
    f33.value = '=SUM(F30:F32)'
    f33.font = Font(bold=True)
    f33.number_format = acct_fmt
    f33.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    g33 = ws1['G33']
    g33.value = '=SUM(G30:G32)'
    g33.font = Font(bold=True)
    g33.number_format = acct_fmt
    g33.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    # Misc Costs Table
    c34 = ws1['C34']
    c34.value = 'MISCELLANEOUS COSTS'
    c34.font = Font(bold=True)
    c34.border = Border(bottom=thin)

    c35 = ws1['C35']
    c35.value = 'Recording Release(s)'
    c35.border = Border(top=thin, bottom=hair, left=thin)

    d35 = ws1['D35']
    d35.value = 'Commonly Used Fee * Qty of Trusts Recorded'
    d35.border = Border(top=thin, bottom=hair)

    e35 = ws1['E35']
    e35.value = recording_fee * recording_trusts_liens_qty
    e35.number_format = acct_fmt
    e35.border = Border(top=thin, bottom=hair, left=thin, right=thin)

    f35 = ws1['F35']
    f35.value = recording_fee * recording_trusts_liens_qty
    f35.number_format = acct_fmt
    f35.border = Border(top=thin, bottom=hair, left=thin, right=thin)

    g35 = ws1['G35']
    g35.value = recording_fee * recording_trusts_liens_qty
    g35.number_format = acct_fmt
    g35.border = Border(top=thin, bottom=hair, left=thin, right=thin)

    c36 = ws1['C36']
    c36.value = 'Grantor\'s Tax'
    c36.border = Border(top=hair, bottom=hair, left=thin)

    d36 = ws1['D36']
    d36.value = '% of Offer Amount ($)'
    d36.border = Border(top=hair, bottom=hair)

    e36 = ws1['E36']
    e36.value = grantors_tax_pct * offer_1_amt
    e36.number_format = acct_fmt
    e36.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    f36 = ws1['F36']
    f36.value = grantors_tax_pct * offer_2_amt
    f36.number_format = acct_fmt
    f36.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    g36 = ws1['G36']
    g36.value = grantors_tax_pct * offer_2_amt
    g36.number_format = acct_fmt
    g36.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    c37 = ws1['C37']
    c37.value = 'Congestion Relief Tax'
    c37.border = Border(top=hair, bottom=hair, left=thin)

    d37 = ws1['D37']
    d37.value = '% of Offer Amount ($)'
    d37.border = Border(top=hair, bottom=hair)

    e37 = ws1['E37']
    e37.value = congestion_tax_pct * offer_1_amt
    e37.number_format = acct_fmt
    e37.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    f37 = ws1['F37']
    f37.value = congestion_tax_pct * offer_2_amt
    f37.number_format = acct_fmt
    f37.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    g37 = ws1['G37']
    g37.value = congestion_tax_pct * offer_3_amt
    g37.number_format = acct_fmt
    g37.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    c38 = ws1['C38']
    c38.value = 'Pest Inspection'
    c38.border = Border(top=hair, bottom=hair, left=thin)

    d38 = ws1['D38']
    d38.value = 'Commonly Used Fee'
    d38.border = Border(top=hair, bottom=hair)

    e38 = ws1['E38']
    e38.value = pest_inspection_fee
    e38.number_format = acct_fmt
    e38.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    f38 = ws1['F38']
    f38.value = pest_inspection_fee
    f38.number_format = acct_fmt
    f38.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    g38 = ws1['G38']
    g38.value = pest_inspection_fee
    g38.number_format = acct_fmt
    g38.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    c39 = ws1['C39']
    c39.value = 'POA / Condo Disclosures'
    c39.border = Border(top=hair, bottom=hair, left=thin)

    d39 = ws1['D39']
    d39.value = 'Commonly Used Fee'
    d39.border = Border(top=hair, bottom=hair)

    e39 = ws1['E39']
    e39.value = poa_condo_disclosure_fee
    e39.number_format = acct_fmt
    e39.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    f39 = ws1['F39']
    f39.value = poa_condo_disclosure_fee
    f39.number_format = acct_fmt
    f39.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    g39 = ws1['G39']
    g39.value = poa_condo_disclosure_fee
    g39.number_format = acct_fmt
    g39.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    c40 = ws1['C40']
    c40.value = 'Post Occupancy Cost to Seller'
    c40.border = Border(top=hair, bottom=thin, left=thin)

    d40 = ws1['D40']
    d40.value = 'Negotiated Into Contract'
    d40.border = Border(top=hair, bottom=thin)

    e40 = ws1['E40']
    e40.value = offer_1_post_occupancy_cost_amt
    e40.number_format = acct_fmt
    e40.border = Border(top=hair, bottom=thin, left=thin, right=thin)

    f40 = ws1['F40']
    f40.value = offer_2_post_occupancy_cost_amt
    f40.number_format = acct_fmt
    f40.border = Border(top=hair, bottom=thin, left=thin, right=thin)

    g40 = ws1['G40']
    g40.value = offer_3_post_occupancy_cost_amt
    g40.number_format = acct_fmt
    g40.border = Border(top=hair, bottom=thin, left=thin, right=thin)

    e41 = ws1['E41']
    e41.value = '=SUM(E35:E40)'
    e41.font = Font(bold=True)
    e41.number_format = acct_fmt
    e41.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    f41 = ws1['F41']
    f41.value = '=SUM(F35:F40)'
    f41.font = Font(bold=True)
    f41.number_format = acct_fmt
    f41.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    g41 = ws1['G41']
    g41.value = '=SUM(G35:G40)'
    g41.font = Font(bold=True)
    g41.number_format = acct_fmt
    g41.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    # Subtotals section
    ws1.merge_cells('C43:D43')
    top_left_cell_four = ws1['C43']
    top_left_cell_four.value = 'TOTAL ESTIMATED COST OF SETTLEMENT'
    top_left_cell_four.font = Font(bold=True)
    top_left_cell_four.alignment = Alignment(horizontal='right')

    e43 = ws1['E43']
    e43.value = '=SUM(E23,E28,E33,E41)'
    e43.font = Font(bold=True)
    e43.number_format = acct_fmt
    e43.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    f43 = ws1['F43']
    f43.value = '=SUM(F23,F28,F33,F41)'
    f43.font = Font(bold=True)
    f43.number_format = acct_fmt
    f43.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    g43 = ws1['G43']
    g43.value = '=SUM(G23,G28,G33,G41)'
    g43.font = Font(bold=True)
    g43.number_format = acct_fmt
    g43.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    ws1.merge_cells('C44:D44')
    top_left_cell_five = ws1['C44']
    top_left_cell_five.value = 'Offer Amount ($)'
    top_left_cell_five.font = Font(bold=True)
    top_left_cell_five.alignment = Alignment(horizontal='right')

    e44 = ws1['E44']
    e44.value = offer_1_amt
    e44.font = Font(bold=True)
    e44.number_format = acct_fmt
    e44.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    f44 = ws1['F44']
    f44.value = offer_2_amt
    f44.font = Font(bold=True)
    f44.number_format = acct_fmt
    f44.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    g44 = ws1['G44']
    g44.value = offer_3_amt
    g44.font = Font(bold=True)
    g44.number_format = acct_fmt
    g44.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    ws1.merge_cells('C45:D45')
    top_left_cell_six = ws1['C45']
    top_left_cell_six.value = 'LESS: Total Estimated Cost of Settlement'
    top_left_cell_six.font = Font(bold=True)
    top_left_cell_six.alignment = Alignment(horizontal='right')

    e45 = ws1['E45']
    e45.value = '=-SUM(E23,E28,E33,E41)'
    e45.font = Font(bold=True)
    e45.number_format = acct_fmt
    e45.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    f45 = ws1['F45']
    f45.value = '=-SUM(F23,F28,F33,F41)'
    f45.font = Font(bold=True)
    f45.number_format = acct_fmt
    f45.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    g45 = ws1['G45']
    g45.value = '=-SUM(G23,G28,G33,G41)'
    g45.font = Font(bold=True)
    g45.number_format = acct_fmt
    g45.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    ws1.merge_cells('C46:D46')
    top_left_cell_seven = ws1['C46']
    top_left_cell_seven.value = 'ESTIMATED TOTAL NET PROCEEDS'
    top_left_cell_seven.font = Font(bold=True)
    top_left_cell_seven.alignment = Alignment(horizontal='right')

    e46 = ws1['E46']
    e46.value = '=SUM(E44:E45)'
    e46.font = Font(bold=True)
    e46.number_format = acct_fmt
    e46.fill = PatternFill(start_color=yellow_fill, end_color=yellow_fill, fill_type='solid')
    e46.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    f46 = ws1['F46']
    f46.value = '=SUM(F44:F45)'
    f46.font = Font(bold=True)
    f46.number_format = acct_fmt
    f46.fill = PatternFill(start_color=yellow_fill, end_color=yellow_fill, fill_type='solid')
    f46.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    g46 = ws1['G46']
    g46.value = '=SUM(G44:G45)'
    g46.font = Font(bold=True)
    g46.number_format = acct_fmt
    g46.fill = PatternFill(start_color=yellow_fill, end_color=yellow_fill, fill_type='solid')
    g46.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    # Disclosure Statement
    ws1.merge_cells('C47:G49')
    top_left_cell_eight = ws1['C47']
    top_left_cell_eight.value = '''These estimates are not guaranteed and may not include escrows. Escrow balances are reimbursed by the existing lender. Taxes, rents & association dues are pro-rated at settlement. Under Virginia Law, the seller's proceeds may not be available for up to 2 business days following recording of the deed. Seller acknowledges receipt of this statement.'''
    top_left_cell_eight.font = Font(italic=True)
    top_left_cell_eight.alignment = Alignment(wrapText=True)

    # Signature Block
    c51 = ws1['C51']
    c51.value = 'PREPARED BY:'

    c52 = ws1['C52']
    c52.value = agent

    d51 = ws1['D51']
    d51.value = 'SELLER:'

    d52 = ws1['D52']
    d52.value = seller_name

    # Freedom Logo
    # c53 = ws1['C53']
    freedom_logo = drawing.image.Image('freedom_logo.png')
    freedom_logo.height = 80
    freedom_logo.width = 115
    freedom_logo.anchor = 'C53'
    ws1.add_image(freedom_logo)

    if offer_qty == 1:
        ws1.column_dimensions['F'].hidden = True
        ws1.column_dimensions['G'].hidden = True
    elif offer_qty == 2:
        ws1.column_dimensions['G'].hidden = True

    # wb.save(filename=dest_filename)

    with NamedTemporaryFile() as tmp:
        wb.save(tmp.name)
        data = BytesIO(tmp.read())

    return data

# comparison_inputs_to_excel(
#     agent='Tiffany Izenour, Principal Broker & Owner',
#     date='2022-04-10',
#     offer_qty=1,
#     seller_name='Michael and Patrica Tracey',
#     seller_address='9517 Basilwood Dr., Manassas, VA 20110',
#     list_price=550000,
#     first_trust=296000,
#     second_trust=0,
#     annual_taxes=4776,
#     prorated_taxes=4776/12*3,
#     annual_hoa_condo_fees=0,
#     prorated_hoa_condo_fees=0,
#     listing_company_pct=0.025,
#     selling_company_pct=0.025,
#     processing_fee=0,
#     settlement_fee=450,
#     deed_preparation_fee=150,
#     lien_trust_release_fee=100,
#     lien_trust_release_qty=1,
#     recording_fee=38,
#     recording_trusts_liens_qty=1,
#     grantors_tax_pct=0.001,
#     congestion_tax_pct=0.002,
#     pest_inspection_fee=50,
#     poa_condo_disclosure_fee=350,
#     offer_1_name='Audi/Romero',
#     offer_1_amt=545000,
#     offer_1_down_pmt_pct=0.03,
#     offer_1_settlement_date='2022-05-11',
#     offer_1_settlement_company='Title co.',
#     offer_1_emd_amt=5000,
#     offer_1_financing_type='VA',
#     offer_1_contingencies_waived='financing appraisal',
#     offer_1_post_occupancy_thru_date='7/3/2022',
#     offer_1_closing_cost_subsidy_amt=0,
#     offer_1_post_occupancy_cost_amt=1000,
#     offer_2_name='Gallichio',
#     offer_2_amt=560000,
#     offer_2_down_pmt_pct=0.02,
#     offer_2_settlement_date='2022-05-15',
#     offer_2_settlement_company='Title co.',
#     offer_2_emd_amt=5000,
#     offer_2_financing_type='VA',
#     offer_2_contingencies_waived='financing',
#     offer_2_post_occupancy_thru_date='7/3/2022',
#     offer_2_closing_cost_subsidy_amt=6000,
#     offer_2_post_occupancy_cost_amt=0,
#     offer_3_name='Hilliard',
#     offer_3_amt=540000,
#     offer_3_down_pmt_pct=0.05,
#     offer_3_settlement_date='2022-06-11',
#     offer_3_settlement_company='Title co.',
#     offer_3_emd_amt=5000,
#     offer_3_financing_type='VA',
#     offer_3_contingencies_waived='financing',
#     offer_3_post_occupancy_thru_date='7/3/2022',
#     offer_3_closing_cost_subsidy_amt=1000,
#     offer_3_post_occupancy_cost_amt=3000,
# )
