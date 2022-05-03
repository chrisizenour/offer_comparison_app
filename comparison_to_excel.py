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
        offer_1_home_inspection_check,
        offer_1_home_inspection_days,
        offer_1_radon_inspection_check,
        offer_1_radon_inspection_days,
        offer_1_septic_inspection_check,
        offer_1_septic_inspection_days,
        offer_1_well_inspection_check,
        offer_1_well_inspection_days,
        offer_1_finance_contingency_check,
        offer_1_finance_contingency_days,
        offer_1_appraisal_contingency_check,
        offer_1_appraisal_contingency_days,
        offer_1_home_sale_contingency_check,
        offer_1_home_sale_contingency_days,
        offer_1_pre_occupancy_date,
        offer_1_post_occupancy_date,
        offer_1_closing_cost_subsidy_amt,
        offer_1_pre_occupancy_credit_amt,
        offer_1_post_occupancy_cost_amt,
        offer_2_name,
        offer_2_amt,
        offer_2_down_pmt_pct,
        offer_2_settlement_date,
        offer_2_settlement_company,
        offer_2_emd_amt,
        offer_2_financing_type,
        offer_2_home_inspection_check,
        offer_2_home_inspection_days,
        offer_2_radon_inspection_check,
        offer_2_radon_inspection_days,
        offer_2_septic_inspection_check,
        offer_2_septic_inspection_days,
        offer_2_well_inspection_check,
        offer_2_well_inspection_days,
        offer_2_finance_contingency_check,
        offer_2_finance_contingency_days,
        offer_2_appraisal_contingency_check,
        offer_2_appraisal_contingency_days,
        offer_2_home_sale_contingency_check,
        offer_2_home_sale_contingency_days,
        offer_2_pre_occupancy_date,
        offer_2_post_occupancy_date,
        offer_2_closing_cost_subsidy_amt,
        offer_2_pre_occupancy_credit_amt,
        offer_2_post_occupancy_cost_amt,
        offer_3_name,
        offer_3_amt,
        offer_3_down_pmt_pct,
        offer_3_settlement_date,
        offer_3_settlement_company,
        offer_3_emd_amt,
        offer_3_financing_type,
        offer_3_home_inspection_check,
        offer_3_home_inspection_days,
        offer_3_radon_inspection_check,
        offer_3_radon_inspection_days,
        offer_3_septic_inspection_check,
        offer_3_septic_inspection_days,
        offer_3_well_inspection_check,
        offer_3_well_inspection_days,
        offer_3_finance_contingency_check,
        offer_3_finance_contingency_days,
        offer_3_appraisal_contingency_check,
        offer_3_appraisal_contingency_days,
        offer_3_home_sale_contingency_check,
        offer_3_home_sale_contingency_days,
        offer_3_pre_occupancy_date,
        offer_3_post_occupancy_date,
        offer_3_closing_cost_subsidy_amt,
        offer_3_pre_occupancy_credit_amt,
        offer_3_post_occupancy_cost_amt,
):

    wb = Workbook()
    dest_filename = f"offer_comparison_form_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

    ws1 = wb.active
    ws1.title = 'offer_comparison'
    ws1.print_area = 'B2:H75'
    ws1.set_printer_settings(paper_size=1, orientation='portrait')

    ws1.page_margins.top = 0.5
    ws1.page_margins.bottom = 0.5
    ws1.page_margins.left = 0.5
    ws1.page_margins.right = 0.5
    ws1.sheet_properties.pageSetUpPr.fitToPage = True
    ws1.print_options.horizontalCentered = True
    ws1.print_options.verticalCentered = True


    white_fill = '00FFFFFF'
    yellow_fill = '00FFFF00'
    black_fill = '00000000'
    font_size = 12
    thick = Side(border_style='thick')
    thin = Side(border_style='thin')
    hair = Side(border_style='hair')
    DEFAULT_FONT.size = font_size
    ws1.column_dimensions['B'].width = 1.5
    ws1.column_dimensions['C'].width = 40.83
    ws1.column_dimensions['D'].width = 40.83
    ws1.column_dimensions['E'].width = 20.83
    ws1.column_dimensions['F'].width = 20.83
    ws1.column_dimensions['G'].width = 20.83
    ws1.column_dimensions['H'].width = 1.5

    acct_fmt = '_($* #,##0_);[Red]_($* (#,##0);_($* "-"??_);_(@_)'
    pct_fmt = '0.00%'
    # date_fmt = NamedStyle(name='date', number_format='DD/MM/YYYY')

    for row in ws1.iter_rows(min_row=1, max_row=100, min_col=1, max_col=20):
        for cell in row:
            cell.fill = PatternFill(start_color=white_fill, end_color=white_fill, fill_type='solid')

    # Build Black Border
    ws1.merge_cells('A1:I1')
    top_left_border_one = ws1['A1']
    top_left_border_one.fill = PatternFill('solid', fgColor=black_fill)

    ws1.merge_cells('A2:A76')
    top_left_border_two = ws1['A2']
    top_left_border_two.fill = PatternFill('solid', fgColor=black_fill)

    ws1.merge_cells('B76:I76')
    top_left_border_three = ws1['B76']
    top_left_border_three.fill = PatternFill('solid', fgColor=black_fill)

    ws1.merge_cells('I2:I75')
    top_left_border_four = ws1['I2']
    top_left_border_four.fill = PatternFill('solid', fgColor=black_fill)

    # Build Header
    ws1.merge_cells('C2:G2')
    top_left_cell_one = ws1['C2']
    top_left_cell_one.value = 'Seller\'s Total Net Proceeds For Different Offers'
    top_left_cell_one.font = Font(bold=True)
    top_left_cell_one.alignment = Alignment(horizontal='center')

    ws1.merge_cells('C3:G3')
    top_left_cell_two = ws1['C3']
    top_left_cell_two.value = f'{seller_name} - {seller_address}'
    top_left_cell_two.font = Font(bold=True)
    top_left_cell_two.alignment = Alignment(horizontal='center')

    ws1.merge_cells('C4:G4')
    top_left_cell_three = ws1['C4']
    top_left_cell_three.value = f'Date Prepared: {date}'
    top_left_cell_three.font = Font(bold=True)
    top_left_cell_three.alignment = Alignment(horizontal='center')

    ws1.merge_cells('C5:G5')
    top_left_cell_four = ws1['C5']
    top_left_cell_four.value = f'List Price: ${list_price:,.2f}'
    top_left_cell_four.font = Font(bold=True)
    top_left_cell_four.alignment = Alignment(horizontal='center')

    c7 = ws1['C7']
    c7.value = 'OFFER SUMMARY FEATURES'
    c7.font = Font(bold=True)
    c7.border = Border(bottom=thin)

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

    c8 = ws1['C8']
    c8.value = 'Offer Amt. ($)'
    # c8.font = Font(bold=True)
    c8.alignment = Alignment(horizontal='left', vertical='center')
    c8.border = Border(top=thin, bottom=hair, left=thin)

    d8 = ws1['D8']
    d8.border = Border(top=thin, bottom=hair)

    e8 = ws1['E8']
    e8.value = offer_1_amt
    # e8.font = Font(bold=True)
    e8.number_format = acct_fmt
    e8.border = Border(top=thin, bottom=hair, left=thin, right=thin)

    f8 = ws1['F8']
    f8.value = offer_2_amt
    # f8.font = Font(bold=True)
    f8.number_format = acct_fmt
    f8.border = Border(top=thin, bottom=hair, left=thin, right=thin)

    g8 = ws1['G8']
    g8.value = offer_3_amt
    # g8.font = Font(bold=True)
    g8.number_format = acct_fmt
    g8.border = Border(top=thin, bottom=hair, right=thin, left=thin)

    c9 = ws1['C9']
    c9.value = 'Down Pmt (%)'
    # c9.font = Font(bold=True)
    c9.alignment = Alignment(horizontal='left', vertical='center')
    c9.border = Border(top=hair, bottom=hair, left=thin)

    d9 = ws1['D9']
    d9.border = Border(top=hair, bottom=hair)

    e9 = ws1['E9']
    e9.value = offer_1_down_pmt_pct
    # e9.font = Font(bold=True)
    e9.number_format = pct_fmt
    e9.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    f9 = ws1['F9']
    f9.value = offer_2_down_pmt_pct
    # f9.font = Font(bold=True)
    f9.number_format = pct_fmt
    f9.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    g9 = ws1['G9']
    g9.value = offer_3_down_pmt_pct
    # g9.font = Font(bold=True)
    g9.number_format = pct_fmt
    g9.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    c10 = ws1['C10']
    c10.value = 'Settlement Date'
    # c10.font = Font(bold=True)
    c10.alignment = Alignment(horizontal='left', vertical='center')
    c10.border = Border(top=hair, bottom=hair, left=thin)

    d10 = ws1['D10']
    d10.border = Border(top=hair, bottom=hair)

    e10 = ws1['E10']
    e10.value = offer_1_settlement_date
    # e10.font = Font(bold=True)
    e10.alignment = Alignment(horizontal='right')
    e10.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    f10 = ws1['F10']
    f10.value = offer_2_settlement_date
    # f10.font = Font(bold=True)
    f10.alignment = Alignment(horizontal='right')
    f10.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    g10 = ws1['G10']
    g10.value = offer_3_settlement_date
    # g10.font = Font(bold=True)
    g10.alignment = Alignment(horizontal='right')
    g10.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    c11 = ws1['C11']
    c11.value = 'Settlement Company'
    # c11.font = Font(bold=True)
    c11.alignment = Alignment(horizontal='left', vertical='center')
    c11.border = Border(top=hair, bottom=hair, left=thin)

    d11 = ws1['D11']
    d11.border = Border(top=hair, bottom=hair)

    e11 = ws1['E11']
    e11.value = offer_1_settlement_company
    # e11.font = Font(bold=True)
    e11.alignment = Alignment(horizontal='center', vertical='top', wrap_text=True)
    e11.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    f11 = ws1['F11']
    f11.value = offer_2_settlement_company
    # f11.font = Font(bold=True)
    f11.alignment = Alignment(horizontal='center', vertical='top', wrap_text=True)
    f11.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    g11 = ws1['G11']
    g11.value = offer_3_settlement_company
    # g11.font = Font(bold=True)
    g11.alignment = Alignment(horizontal='center', vertical='top', wrap_text=True)
    g11.border = Border(top=hair, bottom=hair, right=thin, left=thin)

    c12 = ws1['C12']
    c12.value = 'EMD Amt. ($)'
    # c12.font = Font(bold=True)
    c12.alignment = Alignment(horizontal='left', vertical='center')
    c12.border = Border(top=hair, bottom=hair, left=thin)

    d12 = ws1['D12']
    d12.border = Border(top=hair, bottom=hair)

    e12 = ws1['E12']
    e12.value = offer_1_emd_amt
    # e12.font = Font(bold=True)
    e12.number_format = acct_fmt
    e12.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    f12 = ws1['F12']
    f12.value = offer_2_emd_amt
    # f12.font = Font(bold=True)
    f12.number_format = acct_fmt
    f12.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    g12 = ws1['G12']
    g12.value = offer_3_emd_amt
    # g12.font = Font(bold=True)
    g12.number_format = acct_fmt
    g12.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    c13 = ws1['C13']
    c13.value = 'Financing Type'
    # c13.font = Font(bold=True)
    c13.alignment = Alignment(horizontal='left', vertical='center')
    c13.border = Border(top=hair, bottom=hair, left=thin)

    d13 = ws1['D13']
    d13.border = Border(top=hair, bottom=hair)

    e13 = ws1['E13']
    e13.value = offer_1_financing_type
    # e13.font = Font(bold=True)
    e13.alignment = Alignment(horizontal='center')
    e13.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    f13 = ws1['F13']
    f13.value = offer_2_financing_type
    # f13.font = Font(bold=True)
    f13.alignment = Alignment(horizontal='center')
    f13.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    g13 = ws1['G13']
    g13.value = offer_3_financing_type
    # g13.font = Font(bold=True)
    g13.alignment = Alignment(horizontal='center')
    g13.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    ws1.merge_cells('C14:C15')
    top_left_home_insp_cont = ws1['C14']
    top_left_home_insp_cont.value = 'Home Inspection Contingency'
    # top_left_home_insp_cont.font = Font(bold=True)
    top_left_home_insp_cont.alignment = Alignment(horizontal='left', vertical='center')
    top_left_home_insp_cont.border = Border(top=hair, bottom=hair, left=thin)
    c14 = ws1['C14']
    c14.border = Border(top=hair, left=thin)
    d14 = ws1['D14']
    d14.border = Border(top=hair)
    c15 = ws1['C15']
    c15.border = Border(bottom=hair, left=thin)
    d15 = ws1['D15']
    d15.border = Border(bottom=hair)

    e14 = ws1['E14']
    e14.value = offer_1_home_inspection_check
    # e14.font = Font(bold=True)
    e14.alignment = Alignment(horizontal='center')
    e14.border = Border(top=hair, left=thin, right=thin)

    f14 = ws1['F14']
    f14.value = offer_2_home_inspection_check
    # f14.font = Font(bold=True)
    f14.alignment = Alignment(horizontal='center')
    f14.border = Border(top=hair, left=thin, right=thin)

    g14 = ws1['G14']
    g14.value = offer_3_home_inspection_check
    # g14.font = Font(bold=True)
    g14.alignment = Alignment(horizontal='center')
    g14.border = Border(top=hair, left=thin, right=thin)

    e15 = ws1['E15']
    e15.value = offer_1_home_inspection_days
    # e15.font = Font(bold=True)
    e15.alignment = Alignment(horizontal='center')
    e15.border = Border(bottom=hair, left=thin, right=thin)

    f15 = ws1['F15']
    f15.value = offer_2_home_inspection_days
    # f15.font = Font(bold=True)
    f15.alignment = Alignment(horizontal='center')
    f15.border = Border(bottom=hair, left=thin, right=thin)

    g15 = ws1['G15']
    g15.value = offer_3_home_inspection_days
    # g15.font = Font(bold=True)
    g15.alignment = Alignment(horizontal='center')
    g15.border = Border(bottom=hair, left=thin, right=thin)

    ws1.merge_cells('C16:C17')
    top_left_radon_insp_cont = ws1['C16']
    top_left_radon_insp_cont.value = 'Radon Inspection Contingency'
    # top_left_radon_insp_cont.font = Font(bold=True)
    top_left_radon_insp_cont.alignment = Alignment(horizontal='left', vertical='center')
    top_left_radon_insp_cont.border = Border(top=hair, bottom=hair, left=thin)
    c16 = ws1['C16']
    c16.border = Border(top=hair, left=thin)
    d16 = ws1['D16']
    d16.border = Border(top=hair)
    c17 = ws1['C17']
    c17.border = Border(bottom=hair, left=thin)
    d17 = ws1['D17']
    d17.border = Border(bottom=hair)

    e16 = ws1['E16']
    e16.value = offer_1_radon_inspection_check
    # e16.font = Font(bold=True)
    e16.alignment = Alignment(horizontal='center')
    e16.border = Border(top=hair, left=thin, right=thin)

    f16 = ws1['F16']
    f16.value = offer_2_radon_inspection_check
    # f16.font = Font(bold=True)
    f16.alignment = Alignment(horizontal='center')
    f16.border = Border(top=hair, left=thin, right=thin)

    g16 = ws1['G16']
    g16.value = offer_3_radon_inspection_check
    # g16.font = Font(bold=True)
    g16.alignment = Alignment(horizontal='center')
    g16.border = Border(top=hair, left=thin, right=thin)

    e17 = ws1['E17']
    e17.value = offer_1_radon_inspection_days
    # e17.font = Font(bold=True)
    e17.alignment = Alignment(horizontal='center')
    e17.border = Border(bottom=hair, left=thin, right=thin)

    f17 = ws1['F17']
    f17.value = offer_2_radon_inspection_days
    # f17.font = Font(bold=True)
    f17.alignment = Alignment(horizontal='center')
    f17.border = Border(bottom=hair, left=thin, right=thin)

    g17 = ws1['G17']
    g17.value = offer_3_radon_inspection_days
    # g17.font = Font(bold=True)
    g17.alignment = Alignment(horizontal='center')
    g17.border = Border(bottom=hair, left=thin, right=thin)

    ws1.merge_cells('C18:C19')
    top_left_septic_insp_cont = ws1['C18']
    top_left_septic_insp_cont.value = 'Septic Inspection Contingency'
    # top_left_septic_insp_cont.font = Font(bold=True)
    top_left_septic_insp_cont.alignment = Alignment(horizontal='left', vertical='center')
    top_left_septic_insp_cont.border = Border(top=hair, bottom=hair, left=thin)
    c18 = ws1['C18']
    c18.border = Border(top=hair, left=thin)
    d18 = ws1['D18']
    d18.border = Border(top=hair)
    c19 = ws1['C19']
    c19.border = Border(bottom=hair, left=thin)
    d19 = ws1['D19']
    d19.border = Border(bottom=hair)

    e18 = ws1['E18']
    e18.value = offer_1_septic_inspection_check
    # e18.font = Font(bold=True)
    e18.alignment = Alignment(horizontal='center')
    e18.border = Border(top=hair, left=thin, right=thin)

    f18 = ws1['F18']
    f18.value = offer_2_septic_inspection_check
    # f18.font = Font(bold=True)
    f18.alignment = Alignment(horizontal='center')
    f18.border = Border(top=hair, left=thin, right=thin)

    g18 = ws1['G18']
    g18.value = offer_3_septic_inspection_check
    # g18.font = Font(bold=True)
    g18.alignment = Alignment(horizontal='center')
    g18.border = Border(top=hair, left=thin, right=thin)

    e19 = ws1['E19']
    e19.value = offer_1_septic_inspection_days
    # e19.font = Font(bold=True)
    e19.alignment = Alignment(horizontal='center')
    e19.border = Border(bottom=hair, left=thin, right=thin)

    f19 = ws1['F19']
    f19.value = offer_2_septic_inspection_days
    # f19.font = Font(bold=True)
    f19.alignment = Alignment(horizontal='center')
    f19.border = Border(bottom=hair, left=thin, right=thin)

    g19 = ws1['G19']
    g19.value = offer_3_septic_inspection_days
    # g19.font = Font(bold=True)
    g19.alignment = Alignment(horizontal='center')
    g19.border = Border(bottom=hair, left=thin, right=thin)

    ws1.merge_cells('C20:C21')
    top_left_well_insp_cont = ws1['C20']
    top_left_well_insp_cont.value = 'Well Inspection Contingency'
    # top_left_well_insp_cont.font = Font(bold=True)
    top_left_well_insp_cont.alignment = Alignment(horizontal='left', vertical='center')
    top_left_well_insp_cont.border = Border(top=hair, bottom=hair, left=thin)
    c20 = ws1['C20']
    c20.border = Border(top=hair, left=thin)
    d20 = ws1['D20']
    d20.border = Border(top=hair)
    c21 = ws1['C21']
    c21.border = Border(bottom=hair, left=thin)
    d21 = ws1['D21']
    d21.border = Border(bottom=hair)

    e20 = ws1['E20']
    e20.value = offer_1_well_inspection_check
    # e20.font = Font(bold=True)
    e20.alignment = Alignment(horizontal='center')
    e20.border = Border(top=hair, left=thin, right=thin)

    f20 = ws1['F20']
    f20.value = offer_2_well_inspection_check
    # f20.font = Font(bold=True)
    f20.alignment = Alignment(horizontal='center')
    f20.border = Border(top=hair, left=thin, right=thin)

    g20 = ws1['G20']
    g20.value = offer_3_well_inspection_check
    # g20.font = Font(bold=True)
    g20.alignment = Alignment(horizontal='center')
    g20.border = Border(top=hair, left=thin, right=thin)

    e21 = ws1['E21']
    e21.value = offer_1_well_inspection_days
    # e21.font = Font(bold=True)
    e21.alignment = Alignment(horizontal='center')
    e21.border = Border(bottom=hair, left=thin, right=thin)

    f21 = ws1['F21']
    f21.value = offer_2_well_inspection_days
    # f21.font = Font(bold=True)
    f21.alignment = Alignment(horizontal='center')
    f21.border = Border(bottom=hair, left=thin, right=thin)

    g21 = ws1['G21']
    g21.value = offer_3_well_inspection_days
    # g21.font = Font(bold=True)
    g21.alignment = Alignment(horizontal='center')
    g21.border = Border( bottom=hair, left=thin, right=thin)

    ws1.merge_cells('C22:C23')
    top_left_finance_cont = ws1['C22']
    top_left_finance_cont.value = 'Finance Contingency'
    # top_left_finance_cont.font = Font(bold=True)
    top_left_finance_cont.alignment = Alignment(horizontal='left', vertical='center')
    top_left_finance_cont.border = Border(top=hair, bottom=hair, left=thin)
    c22 = ws1['C22']
    c22.border = Border(top=hair, left=thin)
    d22 = ws1['D22']
    d22.border = Border(top=hair)
    c23 = ws1['C23']
    c23.border = Border(bottom=hair, left=thin)
    d23 = ws1['D23']
    d23.border = Border(bottom=hair)

    e22 = ws1['E22']
    e22.value = offer_1_finance_contingency_check
    # e22.font = Font(bold=True)
    e22.alignment = Alignment(horizontal='center')
    e22.border = Border(top=hair, left=thin, right=thin)

    f22 = ws1['F22']
    f22.value = offer_2_finance_contingency_check
    # f22.font = Font(bold=True)
    f22.alignment = Alignment(horizontal='center')
    f22.border = Border(top=hair, left=thin, right=thin)

    g22 = ws1['G22']
    g22.value = offer_3_finance_contingency_check
    # g22.font = Font(bold=True)
    g22.alignment = Alignment(horizontal='center')
    g22.border = Border(top=hair, left=thin, right=thin)

    e23 = ws1['E23']
    e23.value = offer_1_finance_contingency_days
    # e23.font = Font(bold=True)
    e23.alignment = Alignment(horizontal='center')
    e23.border = Border(bottom=hair, left=thin, right=thin)

    f23 = ws1['F23']
    f23.value = offer_2_finance_contingency_days
    # f23.font = Font(bold=True)
    f23.alignment = Alignment(horizontal='center')
    f23.border = Border(bottom=hair, left=thin, right=thin)

    g23 = ws1['G23']
    g23.value = offer_3_finance_contingency_days
    # g23.font = Font(bold=True)
    g23.alignment = Alignment(horizontal='center')
    g23.border = Border(bottom=hair, left=thin, right=thin)

    ws1.merge_cells('C24:C25')
    top_left_appraisal_cont = ws1['C24']
    top_left_appraisal_cont.value = 'Appraisal Contingency'
    # top_left_appraisal_cont.font = Font(bold=True)
    top_left_appraisal_cont.alignment = Alignment(horizontal='left', vertical='center')
    top_left_appraisal_cont.border = Border(top=hair, bottom=hair, left=thin)
    c24 = ws1['C24']
    c24.border = Border(top=hair, left=thin)
    d24 = ws1['D24']
    d24.border = Border(top=hair)
    c25 = ws1['C25']
    c25.border = Border(bottom=hair, left=thin)
    d25 = ws1['D25']
    d25.border = Border(bottom=hair)

    e24 = ws1['E24']
    e24.value = offer_1_appraisal_contingency_check
    # e24.font = Font(bold=True)
    e24.alignment = Alignment(horizontal='center')
    e24.border = Border(top=hair, left=thin, right=thin)

    f24 = ws1['F24']
    f24.value = offer_2_appraisal_contingency_check
    # f24.font = Font(bold=True)
    f24.alignment = Alignment(horizontal='center')
    f24.border = Border(top=hair, left=thin, right=thin)

    g24 = ws1['G24']
    g24.value = offer_3_appraisal_contingency_check
    # g24.font = Font(bold=True)
    g24.alignment = Alignment(horizontal='center')
    g24.border = Border(top=hair, left=thin, right=thin)

    e25 = ws1['E25']
    e25.value = offer_1_appraisal_contingency_days
    # e25.font = Font(bold=True)
    e25.alignment = Alignment(horizontal='center')
    e25.border = Border(bottom=hair, left=thin, right=thin)

    f25 = ws1['F25']
    f25.value = offer_2_appraisal_contingency_days
    # f25.font = Font(bold=True)
    f25.alignment = Alignment(horizontal='center')
    f25.border = Border(bottom=hair, left=thin, right=thin)

    g25 = ws1['G25']
    g25.value = offer_3_appraisal_contingency_days
    # g25.font = Font(bold=True)
    g25.alignment = Alignment(horizontal='center')
    g25.border = Border(bottom=hair, left=thin, right=thin)

    ws1.merge_cells('C26:C27')
    top_left_home_sale_cont = ws1['C26']
    top_left_home_sale_cont.value = 'Home Sale Contingency'
    # top_left_home_sale_cont.font = Font(bold=True)
    top_left_home_sale_cont.alignment = Alignment(horizontal='left', vertical='center')
    top_left_home_sale_cont.border = Border(top=hair, bottom=hair, left=thin)
    c26 = ws1['C26']
    c26.border = Border(top=hair, left=thin)
    d26 = ws1['D26']
    d26.border = Border(top=hair)
    c27 = ws1['C27']
    c27.border = Border(bottom=hair, left=thin)
    d27 = ws1['D27']
    d27.border = Border(bottom=hair)

    e26 = ws1['E26']
    e26.value = offer_1_home_sale_contingency_check
    # e26.font = Font(bold=True)
    e26.alignment = Alignment(horizontal='center')
    e26.border = Border(top=hair, left=thin, right=thin)

    f26 = ws1['F26']
    f26.value = offer_2_home_sale_contingency_check
    # f26.font = Font(bold=True)
    f26.alignment = Alignment(horizontal='center')
    f26.border = Border(top=hair, left=thin, right=thin)

    g26 = ws1['G26']
    g26.value = offer_3_home_sale_contingency_check
    # g26.font = Font(bold=True)
    g26.alignment = Alignment(horizontal='center')
    g26.border = Border(top=hair, left=thin, right=thin)

    e27 = ws1['E27']
    e27.value = offer_1_home_sale_contingency_days
    # e27.font = Font(bold=True)
    e27.alignment = Alignment(horizontal='center')
    e27.border = Border(bottom=hair, left=thin, right=thin)

    f27 = ws1['F27']
    f27.value = offer_2_home_sale_contingency_days
    # f27.font = Font(bold=True)
    f27.alignment = Alignment(horizontal='center')
    f27.border = Border(bottom=hair, left=thin, right=thin)

    g27 = ws1['G27']
    g27.value = offer_3_home_sale_contingency_days
    # g27.font = Font(bold=True)
    g27.alignment = Alignment(horizontal='center')
    g27.border = Border(bottom=hair, left=thin, right=thin)

    c28 = ws1['C28']
    c28.value = 'Pre Occupancy Start Date'
    # c28.font = Font(bold=True)
    c28.alignment = Alignment(horizontal='left', vertical='center')
    c28.border = Border(top=hair, bottom=hair, left=thin)

    d28 = ws1['D28']
    d28.border = Border(top=hair, bottom=hair)

    e28 = ws1['E28']
    e28.value = offer_1_pre_occupancy_date
    # e28.font = Font(bold=True)
    e28.alignment = Alignment(horizontal='right')
    e28.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    f28 = ws1['F28']
    f28.value = offer_2_pre_occupancy_date
    # f28.font = Font(bold=True)
    f28.alignment = Alignment(horizontal='right')
    f28.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    g28 = ws1['G28']
    g28.value = offer_3_pre_occupancy_date
    # g28.font = Font(bold=True)
    g28.alignment = Alignment(horizontal='right')
    g28.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    c29 = ws1['C29']
    c29.value = 'Post Occupancy Thru Date'
    # c29.font = Font(bold=True)
    c29.alignment = Alignment(horizontal='left', vertical='center')
    c29.border = Border(top=hair, bottom=thin, left=thin)

    d29 = ws1['D29']
    d29.border = Border(top=hair, bottom=thin)

    e29 = ws1['E29']
    e29.value = offer_1_post_occupancy_date
    # e29.font = Font(bold=True)
    e29.alignment = Alignment(horizontal='right')
    e29.border = Border(top=hair, bottom=thin, left=thin, right=thin)

    f29 = ws1['F29']
    f29.value = offer_2_post_occupancy_date
    # f29.font = Font(bold=True)
    f29.alignment = Alignment(horizontal='right')
    f29.border = Border(top=hair, bottom=thin, left=thin, right=thin)

    g29 = ws1['G29']
    g29.value = offer_3_post_occupancy_date
    # g29.font = Font(bold=True)
    g29.alignment = Alignment(horizontal='right')
    g29.border = Border(top=hair, bottom=thin, left=thin, right=thin)

    # Housing Related Cost Table
    c31 = ws1['C31']
    c31.value = 'HOUSING-RELATED COSTS'
    c31.font = Font(bold=True)
    c31.border = Border(bottom=thin)

    d31 = ws1['D31']
    d31.value = 'Calculation Description'
    d31.font = Font(bold=True)
    d31.border = Border(bottom=thin)

    c32 = ws1['C32']
    c32.value = 'Estimated Payoff - 1st Trust'
    c32.border = Border(top=thin, bottom=hair, left=thin)

    d32 = ws1['D32']
    d32.value = 'Principal Balance of Loan'
    d32.border = Border(top=thin, bottom=hair)

    e32 = ws1['E32']
    e32.value = first_trust
    e32.number_format = acct_fmt
    e32.border = Border(top=thin, bottom=hair, left=thin, right=thin)

    f32 = ws1['F32']
    f32.value = first_trust
    f32.number_format = acct_fmt
    f32.border = Border(top=thin, bottom=hair, left=thin, right=thin)

    g32 = ws1['G32']
    g32.value = first_trust
    g32.number_format = acct_fmt
    g32.border = Border(top=thin, bottom=hair, left=thin, right=thin)

    c33 = ws1['C33']
    c33.value = 'Estimated Payoff - 2nd Trust'
    c33.border = Border(top=hair, bottom=hair, left=thin)

    d33 = ws1['D33']
    d33.value = 'Principal Balance of Loan'
    d33.border = Border(top=hair, bottom=hair)

    e33 = ws1['E33']
    e33.value = second_trust
    e33.number_format = acct_fmt
    e33.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    f33 = ws1['F33']
    f33.value = second_trust
    f33.number_format = acct_fmt
    f33.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    g33 = ws1['G33']
    g33.value = second_trust
    g33.number_format = acct_fmt
    g33.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    c34 = ws1['C34']
    c34.value = 'Purchaser Closing Cost / Contract'
    c34.border = Border(top=hair, bottom=hair, left=thin)

    d34 = ws1['D34']
    d34.value = 'Negotiated Into Contract'
    d34.border = Border(top=hair, bottom=hair)

    e34 = ws1['E34']
    e34.value = offer_1_closing_cost_subsidy_amt
    e34.number_format = acct_fmt
    e34.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    f34 = ws1['F34']
    f34.value = offer_2_closing_cost_subsidy_amt
    f34.number_format = acct_fmt
    f34.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    g34 = ws1['G34']
    g34.value = offer_3_closing_cost_subsidy_amt
    g34.number_format = acct_fmt
    g34.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    c35 = ws1['C35']
    c35.value = 'Prorated Taxes / Assessments'
    c35.border = Border(top=hair, bottom=hair, left=thin)

    d35 = ws1['D35']
    d35.value = '1 Year of Taxes Divided by 12 Multiplied by 3'
    d35.border = Border(top=hair, bottom=hair)

    e35 = ws1['E35']
    e35.value = prorated_taxes
    e35.number_format = acct_fmt
    e35.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    f35 = ws1['F35']
    f35.value = prorated_taxes
    f35.number_format = acct_fmt
    f35.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    g35 = ws1['G35']
    g35.value = prorated_taxes
    g35.number_format = acct_fmt
    g35.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    c36 = ws1['C36']
    c36.value = 'Prorated HOA / Condo Dues'
    c36.border = Border(top=hair, bottom=thin, left=thin)

    d36 = ws1['D36']
    d36.value = '1 Year of Dues Divided by 12 Multiplied by 3'
    d36.border = Border(top=hair, bottom=thin)

    e36 = ws1['E36']
    e36.value = prorated_hoa_condo_fees
    e36.number_format = acct_fmt
    e36.border = Border(top=hair, bottom=thin, left=thin, right=thin)

    f36 = ws1['F36']
    f36.value = prorated_hoa_condo_fees
    f36.number_format = acct_fmt
    f36.border = Border(top=hair, bottom=thin, left=thin, right=thin)

    g36 = ws1['G36']
    g36.value = prorated_hoa_condo_fees
    g36.number_format = acct_fmt
    g36.border = Border(top=hair, bottom=thin, left=thin, right=thin)

    e37 = ws1['E37']
    e37.value = '=SUM(E32:E36)'
    e37.font = Font(bold=True)
    e37.number_format = acct_fmt
    e37.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    f37 = ws1['F37']
    f37.value = '=SUM(F32:F36)'
    f37.font = Font(bold=True)
    f37.number_format = acct_fmt
    f37.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    g37 = ws1['G37']
    g37.value = '=SUM(G32:G36)'
    g37.font = Font(bold=True)
    g37.number_format = acct_fmt
    g37.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    # Broker and Financing Costs Table
    c38 = ws1['C38']
    c38.value = 'BROKERAGE & FINANCING COSTS'
    c38.font = Font(bold=True)
    c38.border = Border(bottom=thin)

    c39 = ws1['C39']
    c39.value = 'Listing Company Compensation'
    c39.border = Border(top=thin, bottom=hair, left=thin)

    d39 = ws1['D39']
    d39.value = '% from Listing Agreement * Offer Amount ($)'
    d39.border = Border(top=thin, bottom=hair)

    e39 = ws1['E39']
    e39.value = listing_company_pct * offer_1_amt
    e39.number_format = acct_fmt
    e39.border = Border(top=thin, bottom=hair, left=thin, right=thin)

    f39 = ws1['F39']
    f39.value = listing_company_pct * offer_2_amt
    f39.number_format = acct_fmt
    f39.border = Border(top=thin, bottom=hair, left=thin, right=thin)

    g39 = ws1['G39']
    g39.value = listing_company_pct * offer_3_amt
    g39.number_format = acct_fmt
    g39.border = Border(top=thin, bottom=hair, left=thin, right=thin)

    c40 = ws1['C40']
    c40.value = 'Selling Company Compensation'
    c40.border = Border(top=hair, bottom=hair, left=thin)

    d40 = ws1['D40']
    d40.value = '% from Listing Agreement * Offer Amount ($)'
    d40.border = Border(top=hair, bottom=hair)

    e40 = ws1['E40']
    e40.value = selling_company_pct * offer_1_amt
    e40.number_format = acct_fmt
    e40.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    f40 = ws1['F40']
    f40.value = selling_company_pct * offer_2_amt
    f40.number_format = acct_fmt
    f40.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    g40 = ws1['G40']
    g40.value = selling_company_pct * offer_3_amt
    g40.number_format = acct_fmt
    g40.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    c41 = ws1['C41']
    c41.value = 'Processing Fee'
    c41.border = Border(top=hair, bottom=thin, left=thin)

    d41 = ws1['D41']
    d41.border = Border(top=hair, bottom=thin)

    e41 = ws1['E41']
    e41.value = processing_fee
    e41.number_format = acct_fmt
    e41.border = Border(top=hair, bottom=thin, left=thin, right=thin)

    f41 = ws1['F41']
    f41.value = processing_fee
    f41.number_format = acct_fmt
    f41.border = Border(top=hair, bottom=thin, left=thin, right=thin)

    g41 = ws1['G41']
    g41.value = processing_fee
    g41.number_format = acct_fmt
    g41.border = Border(top=hair, bottom=thin, left=thin, right=thin)

    e42 = ws1['E42']
    e42.value = '=SUM(E39:E41)'
    e42.font = Font(bold=True)
    e42.number_format = acct_fmt
    e42.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    f42 = ws1['F42']
    f42.value = '=SUM(F39:F41)'
    f42.font = Font(bold=True)
    f42.number_format = acct_fmt
    f42.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    g42 = ws1['G42']
    g42.value = '=SUM(G39:G41)'
    g42.font = Font(bold=True)
    g42.number_format = acct_fmt
    g42.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    # Estimated Closing Costs Table
    c43 = ws1['C43']
    c43.value = 'ESTIMATED CLOSING COSTS'
    c43.font = Font(bold=True)
    c43.border = Border(bottom=thin)

    c44 = ws1['C44']
    c44.value = 'Settlement Fee'
    c44.border = Border(top=thin, bottom=hair, left=thin)

    d44 = ws1['D44']
    d44.value = 'Commonly Used Fee'
    d44.border = Border(top=thin, bottom=hair)

    e44 = ws1['E44']
    e44.value = settlement_fee
    e44.number_format = acct_fmt
    e44.border = Border(top=thin, bottom=hair, left=thin, right=thin)

    f44 = ws1['F44']
    f44.value = settlement_fee
    f44.number_format = acct_fmt
    f44.border = Border(top=thin, bottom=hair, left=thin, right=thin)

    g44 = ws1['G44']
    g44.value = settlement_fee
    g44.number_format = acct_fmt
    g44.border = Border(top=thin, bottom=hair, left=thin, right=thin)

    c45 = ws1['C45']
    c45.value = 'Deed Preparation'
    c45.border = Border(top=hair, bottom=hair, left=thin)

    d45 = ws1['D45']
    d45.value = 'Commonly Used Fee'
    d45.border = Border(top=hair, bottom=hair)

    e45 = ws1['E45']
    e45.value = deed_preparation_fee
    e45.number_format = acct_fmt
    e45.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    f45 = ws1['F45']
    f45.value = deed_preparation_fee
    f45.number_format = acct_fmt
    f45.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    g45 = ws1['G45']
    g45.value = deed_preparation_fee
    g45.number_format = acct_fmt
    g45.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    c46 = ws1['C46']
    c46.value = 'Release of Liens / Trusts'
    c46.border = Border(top=hair, bottom=thin, left=thin)

    d46 = ws1['D46']
    d46.value = 'Commonly Used Fee * Qty of Trusts or Liens'
    d46.border = Border(top=hair, bottom=thin)

    e46 = ws1['E46']
    e46.value = lien_trust_release_fee * lien_trust_release_qty
    e46.number_format = acct_fmt
    e46.border = Border(top=hair, bottom=thin, left=thin, right=thin)

    f46 = ws1['F46']
    f46.value = lien_trust_release_fee * lien_trust_release_qty
    f46.number_format = acct_fmt
    f46.border = Border(top=hair, bottom=thin, left=thin, right=thin)

    g46 = ws1['G46']
    g46.value = lien_trust_release_fee * lien_trust_release_qty
    g46.number_format = acct_fmt
    g46.border = Border(top=hair, bottom=thin, left=thin, right=thin)

    e47 = ws1['E47']
    e47.value = '=SUM(E44:E46)'
    e47.font = Font(bold=True)
    e47.number_format = acct_fmt
    e47.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    f47 = ws1['F47']
    f47.value = '=SUM(F44:F46)'
    f47.font = Font(bold=True)
    f47.number_format = acct_fmt
    f47.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    g47 = ws1['G47']
    g47.value = '=SUM(G44:G46)'
    g47.font = Font(bold=True)
    g47.number_format = acct_fmt
    g47.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    # Misc Costs Table
    c48 = ws1['C48']
    c48.value = 'MISCELLANEOUS COSTS'
    c48.font = Font(bold=True)
    c48.border = Border(bottom=thin)

    c49 = ws1['C49']
    c49.value = 'Recording Release(s)'
    c49.border = Border(top=thin, bottom=hair, left=thin)

    d49 = ws1['D49']
    d49.value = 'Commonly Used Fee * Qty of Trusts Recorded'
    d49.border = Border(top=thin, bottom=hair)

    e49 = ws1['E49']
    e49.value = recording_fee * recording_trusts_liens_qty
    e49.number_format = acct_fmt
    e49.border = Border(top=thin, bottom=hair, left=thin, right=thin)

    f49 = ws1['F49']
    f49.value = recording_fee * recording_trusts_liens_qty
    f49.number_format = acct_fmt
    f49.border = Border(top=thin, bottom=hair, left=thin, right=thin)

    g49 = ws1['G49']
    g49.value = recording_fee * recording_trusts_liens_qty
    g49.number_format = acct_fmt
    g49.border = Border(top=thin, bottom=hair, left=thin, right=thin)

    c50 = ws1['C50']
    c50.value = 'Grantor\'s Tax'
    c50.border = Border(top=hair, bottom=hair, left=thin)

    d50 = ws1['D50']
    d50.value = '% of Offer Amount ($)'
    d50.border = Border(top=hair, bottom=hair)

    e50 = ws1['E50']
    e50.value = grantors_tax_pct * offer_1_amt
    e50.number_format = acct_fmt
    e50.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    f50 = ws1['F50']
    f50.value = grantors_tax_pct * offer_2_amt
    f50.number_format = acct_fmt
    f50.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    g50 = ws1['G50']
    g50.value = grantors_tax_pct * offer_2_amt
    g50.number_format = acct_fmt
    g50.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    c51 = ws1['C51']
    c51.value = 'Congestion Relief Tax'
    c51.border = Border(top=hair, bottom=hair, left=thin)

    d51 = ws1['D51']
    d51.value = '% of Offer Amount ($)'
    d51.border = Border(top=hair, bottom=hair)

    e51 = ws1['E51']
    e51.value = congestion_tax_pct * offer_1_amt
    e51.number_format = acct_fmt
    e51.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    f51 = ws1['F51']
    f51.value = congestion_tax_pct * offer_2_amt
    f51.number_format = acct_fmt
    f51.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    g51 = ws1['G51']
    g51.value = congestion_tax_pct * offer_3_amt
    g51.number_format = acct_fmt
    g51.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    c52 = ws1['C52']
    c52.value = 'Pest Inspection'
    c52.border = Border(top=hair, bottom=hair, left=thin)

    d52 = ws1['D52']
    d52.value = 'Commonly Used Fee'
    d52.border = Border(top=hair, bottom=hair)

    e52 = ws1['E52']
    e52.value = pest_inspection_fee
    e52.number_format = acct_fmt
    e52.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    f52 = ws1['F52']
    f52.value = pest_inspection_fee
    f52.number_format = acct_fmt
    f52.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    g52 = ws1['G52']
    g52.value = pest_inspection_fee
    g52.number_format = acct_fmt
    g52.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    c53 = ws1['C53']
    c53.value = 'POA / Condo Disclosures'
    c53.border = Border(top=hair, bottom=hair, left=thin)

    d53 = ws1['D53']
    d53.value = 'Commonly Used Fee'
    d53.border = Border(top=hair, bottom=hair)

    e53 = ws1['E53']
    e53.value = poa_condo_disclosure_fee
    e53.number_format = acct_fmt
    e53.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    f53 = ws1['F53']
    f53.value = poa_condo_disclosure_fee
    f53.number_format = acct_fmt
    f53.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    g53 = ws1['G53']
    g53.value = poa_condo_disclosure_fee
    g53.number_format = acct_fmt
    g53.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    c54 = ws1['C54']
    c54.value = 'Pre Occupancy Credit to Seller'
    c54.border = Border(top=hair, bottom=hair, left=thin)

    d54 = ws1['D54']
    d54.value = 'Negotiated Into Contract'
    d54.border = Border(top=hair, bottom=hair)

    e54 = ws1['E54']
    e54.value = offer_1_pre_occupancy_credit_amt
    e54.number_format = acct_fmt
    e54.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    f54 = ws1['F54']
    f54.value = offer_2_pre_occupancy_credit_amt
    f54.number_format = acct_fmt
    f54.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    g54 = ws1['G54']
    g54.value = offer_3_pre_occupancy_credit_amt
    g54.number_format = acct_fmt
    g54.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    c55 = ws1['C55']
    c55.value = 'Post Occupancy Cost to Seller'
    c55.border = Border(top=hair, bottom=thin, left=thin)

    d55 = ws1['D55']
    d55.value = 'Negotiated Into Contract'
    d55.border = Border(top=hair, bottom=thin)

    e55 = ws1['E55']
    e55.value = offer_1_post_occupancy_cost_amt
    e55.number_format = acct_fmt
    e55.border = Border(top=hair, bottom=thin, left=thin, right=thin)

    f55 = ws1['F55']
    f55.value = offer_2_post_occupancy_cost_amt
    f55.number_format = acct_fmt
    f55.border = Border(top=hair, bottom=thin, left=thin, right=thin)

    g55 = ws1['G55']
    g55.value = offer_3_post_occupancy_cost_amt
    g55.number_format = acct_fmt
    g55.border = Border(top=hair, bottom=thin, left=thin, right=thin)

    e56 = ws1['E56']
    e56.value = '=SUM(E49:E53,E55)-E54'
    e56.font = Font(bold=True)
    e56.number_format = acct_fmt
    e56.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    f56 = ws1['F56']
    f56.value = '=SUM(F49:F53,F55)-F54'
    f56.font = Font(bold=True)
    f56.number_format = acct_fmt
    f56.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    g56 = ws1['G56']
    g56.value = '=SUM(G49:G53,G55)-G54'
    g56.font = Font(bold=True)
    g56.number_format = acct_fmt
    g56.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    # Subtotals section
    ws1.merge_cells('C58:D58')
    top_left_cell_four = ws1['C58']
    top_left_cell_four.value = 'TOTAL ESTIMATED COST OF SETTLEMENT'
    top_left_cell_four.font = Font(bold=True)
    top_left_cell_four.alignment = Alignment(horizontal='right')

    e58 = ws1['E58']
    e58.value = '=SUM(E37,E42,E47,E56)'
    e58.font = Font(bold=True)
    e58.number_format = acct_fmt
    e58.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    f58 = ws1['F58']
    f58.value = '=SUM(F37,F42,F47,F56)'
    f58.font = Font(bold=True)
    f58.number_format = acct_fmt
    f58.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    g58 = ws1['G58']
    g58.value = '=SUM(G37,G42,G47,G56)'
    g58.font = Font(bold=True)
    g58.number_format = acct_fmt
    g58.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    ws1.merge_cells('C59:D59')
    top_left_cell_five = ws1['C59']
    top_left_cell_five.value = 'Offer Amount ($)'
    top_left_cell_five.font = Font(bold=True)
    top_left_cell_five.alignment = Alignment(horizontal='right')

    e59 = ws1['E59']
    e59.value = offer_1_amt
    e59.font = Font(bold=True)
    e59.number_format = acct_fmt
    e59.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    f59 = ws1['F59']
    f59.value = offer_2_amt
    f59.font = Font(bold=True)
    f59.number_format = acct_fmt
    f59.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    g59 = ws1['G59']
    g59.value = offer_3_amt
    g59.font = Font(bold=True)
    g59.number_format = acct_fmt
    g59.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    ws1.merge_cells('C60:D60')
    top_left_cell_six = ws1['C60']
    top_left_cell_six.value = 'LESS: Total Estimated Cost of Settlement'
    top_left_cell_six.font = Font(bold=True)
    top_left_cell_six.alignment = Alignment(horizontal='right')

    e60 = ws1['E60']
    e60.value = '=-SUM(E37,E42,E47,E56)'
    e60.font = Font(bold=True)
    e60.number_format = acct_fmt
    e60.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    f60 = ws1['F60']
    f60.value = '=-SUM(F37,F42,F47,F56)'
    f60.font = Font(bold=True)
    f60.number_format = acct_fmt
    f60.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    g60 = ws1['G60']
    g60.value = '=-SUM(G37,G42,G47,G56)'
    g60.font = Font(bold=True)
    g60.number_format = acct_fmt
    g60.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    ws1.merge_cells('C61:D61')
    top_left_cell_seven = ws1['C61']
    top_left_cell_seven.value = 'ESTIMATED TOTAL NET PROCEEDS'
    top_left_cell_seven.font = Font(bold=True)
    top_left_cell_seven.alignment = Alignment(horizontal='right')

    e61 = ws1['E61']
    e61.value = '=SUM(E59:E60)'
    e61.font = Font(bold=True)
    e61.number_format = acct_fmt
    e61.fill = PatternFill(start_color=yellow_fill, end_color=yellow_fill, fill_type='solid')
    e61.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    f61 = ws1['F61']
    f61.value = '=SUM(F59:F60)'
    f61.font = Font(bold=True)
    f61.number_format = acct_fmt
    f61.fill = PatternFill(start_color=yellow_fill, end_color=yellow_fill, fill_type='solid')
    f61.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    g61 = ws1['G61']
    g61.value = '=SUM(G59:G60)'
    g61.font = Font(bold=True)
    g61.number_format = acct_fmt
    g61.fill = PatternFill(start_color=yellow_fill, end_color=yellow_fill, fill_type='solid')
    g61.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    # Signature Block
    c71 = ws1['C63']
    c71.value = 'PREPARED BY:'

    c72 = ws1['C64']
    c72.value = agent

    e71 = ws1['E63']
    e71.value = 'SELLER:'

    e72 = ws1['E64']
    e72.value = seller_name

    # Freedom Logo
    # c53 = ws1['C53']
    freedom_logo = drawing.image.Image('freedom_logo.png')
    freedom_logo.height = 80
    freedom_logo.width = 115
    freedom_logo.anchor = 'C66'
    ws1.add_image(freedom_logo)

    # Disclosure Statement
    ws1.merge_cells('C70:G74')
    top_left_cell_eight = ws1['C70']
    top_left_cell_eight.value = '''These estimates are not guaranteed and may not include escrows. Escrow balances are reimbursed by the existing lender. Taxes, rents & association dues are pro-rated at settlement. Under Virginia Law, the seller's proceeds may not be available for up to 2 business days following recording of the deed. Seller acknowledges receipt of this statement.'''
    top_left_cell_eight.font = Font(italic=True)
    top_left_cell_eight.alignment = Alignment(horizontal='left', vertical='top', wrapText=True)

    # hide columns
    ws1.column_dimensions['D'].hidden = True

    # hide unused columns based on number of offers
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
