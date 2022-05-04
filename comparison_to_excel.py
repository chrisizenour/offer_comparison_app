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
        offer_4_name,
        offer_4_amt,
        offer_4_down_pmt_pct,
        offer_4_settlement_date,
        offer_4_settlement_company,
        offer_4_emd_amt,
        offer_4_financing_type,
        offer_4_home_inspection_check,
        offer_4_home_inspection_days,
        offer_4_radon_inspection_check,
        offer_4_radon_inspection_days,
        offer_4_septic_inspection_check,
        offer_4_septic_inspection_days,
        offer_4_well_inspection_check,
        offer_4_well_inspection_days,
        offer_4_finance_contingency_check,
        offer_4_finance_contingency_days,
        offer_4_appraisal_contingency_check,
        offer_4_appraisal_contingency_days,
        offer_4_home_sale_contingency_check,
        offer_4_home_sale_contingency_days,
        offer_4_pre_occupancy_date,
        offer_4_post_occupancy_date,
        offer_4_closing_cost_subsidy_amt,
        offer_4_pre_occupancy_credit_amt,
        offer_4_post_occupancy_cost_amt,
):

    wb = Workbook()
    dest_filename = f"offer_comparison_form_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

    ws1 = wb.active
    ws1.title = 'offer_comparison_pg_1'
    ws1.print_area = 'B2:I75'
    ws1.set_printer_settings(paper_size=1, orientation='portrait')

    ws1.page_margins.top = 0.5
    ws1.page_margins.bottom = 0.5
    ws1.page_margins.left = 0.5
    ws1.page_margins.right = 0.5
    ws1.sheet_properties.pageSetUpPr.fitToPage = True
    ws1.print_options.horizontalCentered = True
    ws1.print_options.verticalCentered = True

    ws2 = wb.create_sheet(title='offer_comparison_pg_2')
    ws2.print_area = 'B2:I75'
    ws2.set_printer_settings(paper_size=1, orientation='portrait')

    ws2.page_margins.top = 0.5
    ws2.page_margins.bottom = 0.5
    ws2.page_margins.left = 0.5
    ws2.page_margins.right = 0.5
    ws2.sheet_properties.pageSetUpPr.fitToPage = True
    ws2.print_options.horizontalCentered = True
    ws2.print_options.verticalCentered = True

    ws3 = wb.create_sheet(title='offer_comparison_pg_3')
    ws3.print_area = 'B2:I75'
    ws3.set_printer_settings(paper_size=1, orientation='portrait')

    ws3.page_margins.top = 0.5
    ws3.page_margins.bottom = 0.5
    ws3.page_margins.left = 0.5
    ws3.page_margins.right = 0.5
    ws3.sheet_properties.pageSetUpPr.fitToPage = True
    ws3.print_options.horizontalCentered = True
    ws3.print_options.verticalCentered = True


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
    ws1.column_dimensions['H'].width = 20.83
    ws1.column_dimensions['I'].width = 1.5

    ws2.column_dimensions['B'].width = 1.5
    ws2.column_dimensions['C'].width = 40.83
    ws2.column_dimensions['D'].width = 40.83
    ws2.column_dimensions['E'].width = 20.83
    ws2.column_dimensions['F'].width = 20.83
    ws2.column_dimensions['G'].width = 20.83
    ws2.column_dimensions['H'].width = 20.83
    ws2.column_dimensions['I'].width = 1.5

    ws3.column_dimensions['B'].width = 1.5
    ws3.column_dimensions['C'].width = 40.83
    ws3.column_dimensions['D'].width = 40.83
    ws3.column_dimensions['E'].width = 20.83
    ws3.column_dimensions['F'].width = 20.83
    ws3.column_dimensions['G'].width = 20.83
    ws3.column_dimensions['H'].width = 20.83
    ws3.column_dimensions['I'].width = 1.5

    acct_fmt = '_($* #,##0_);[Red]_($* (#,##0);_($* "-"??_);_(@_)'
    pct_fmt = '0.00%'
    # date_fmt = NamedStyle(name='date', number_format='DD/MM/YYYY')

    for row in ws1.iter_rows(min_row=1, max_row=100, min_col=1, max_col=20):
        for cell in row:
            cell.fill = PatternFill(start_color=white_fill, end_color=white_fill, fill_type='solid')

    for row in ws2.iter_rows(min_row=1, max_row=100, min_col=1, max_col=20):
        for cell in row:
            cell.fill = PatternFill(start_color=white_fill, end_color=white_fill, fill_type='solid')

    for row in ws3.iter_rows(min_row=1, max_row=100, min_col=1, max_col=20):
        for cell in row:
            cell.fill = PatternFill(start_color=white_fill, end_color=white_fill, fill_type='solid')

    # Build Black Border for pages 1, 2, and 3
    ws1.merge_cells('A1:J1')
    top_left_border_one_1 = ws1['A1']
    top_left_border_one_1.fill = PatternFill('solid', fgColor=black_fill)

    ws1.merge_cells('A2:A76')
    top_left_border_two_1 = ws1['A2']
    top_left_border_two_1.fill = PatternFill('solid', fgColor=black_fill)

    ws1.merge_cells('B76:J76')
    top_left_border_three_1 = ws1['B76']
    top_left_border_three_1.fill = PatternFill('solid', fgColor=black_fill)

    ws1.merge_cells('J2:J75')
    top_left_border_four_1 = ws1['J2']
    top_left_border_four_1.fill = PatternFill('solid', fgColor=black_fill)

    ws2.merge_cells('A1:J1')
    top_left_border_one_2 = ws2['A1']
    top_left_border_one_2.fill = PatternFill('solid', fgColor=black_fill)

    ws2.merge_cells('A2:A76')
    top_left_border_two_2 = ws2['A2']
    top_left_border_two_2.fill = PatternFill('solid', fgColor=black_fill)

    ws2.merge_cells('B76:J76')
    top_left_border_three_2 = ws2['B76']
    top_left_border_three_2.fill = PatternFill('solid', fgColor=black_fill)

    ws2.merge_cells('J2:J75')
    top_left_border_four_2 = ws2['J2']
    top_left_border_four_2.fill = PatternFill('solid', fgColor=black_fill)

    ws3.merge_cells('A1:J1')
    top_left_border_one_3 = ws3['A1']
    top_left_border_one_3.fill = PatternFill('solid', fgColor=black_fill)

    ws3.merge_cells('A2:A76')
    top_left_border_two_3 = ws3['A2']
    top_left_border_two_3.fill = PatternFill('solid', fgColor=black_fill)

    ws3.merge_cells('B76:J76')
    top_left_border_three_3 = ws3['B76']
    top_left_border_three_3.fill = PatternFill('solid', fgColor=black_fill)

    ws3.merge_cells('J2:J75')
    top_left_border_four_3 = ws3['J2']
    top_left_border_four_3.fill = PatternFill('solid', fgColor=black_fill)

    # page 1
    # Build Header
    ws1.merge_cells('C2:H2')
    top_left_cell_one_1 = ws1['C2']
    top_left_cell_one_1.value = 'Seller\'s Total Net Proceeds For Different Offers'
    top_left_cell_one_1.font = Font(bold=True)
    top_left_cell_one_1.alignment = Alignment(horizontal='center')

    ws1.merge_cells('C3:H3')
    top_left_cell_two_1 = ws1['C3']
    top_left_cell_two_1.value = f'{seller_name} - {seller_address}'
    top_left_cell_two_1.font = Font(bold=True)
    top_left_cell_two_1.alignment = Alignment(horizontal='center')

    ws1.merge_cells('C4:H4')
    top_left_cell_three_1 = ws1['C4']
    top_left_cell_three_1.value = f'Date Prepared: {date}'
    top_left_cell_three_1.font = Font(bold=True)
    top_left_cell_three_1.alignment = Alignment(horizontal='center')

    ws1.merge_cells('C5:H5')
    top_left_cell_four_1 = ws1['C5']
    top_left_cell_four_1.value = f'List Price: ${list_price:,.2f}'
    top_left_cell_four_1.font = Font(bold=True)
    top_left_cell_four_1.alignment = Alignment(horizontal='center')

    c7_1 = ws1['C7']
    c7_1.value = 'OFFER SUMMARY FEATURES'
    c7_1.font = Font(bold=True)
    c7_1.border = Border(bottom=thin)

    e7_1 = ws1['E7']
    e7_1.value = offer_1_name
    e7_1.font = Font(bold=True)
    e7_1.border = Border(bottom=thin)
    e7_1.alignment = Alignment(horizontal='center', wrap_text=True)

    f7_1 = ws1['F7']
    f7_1.value = offer_2_name
    f7_1.font = Font(bold=True)
    f7_1.border = Border(bottom=thin)
    f7_1.alignment = Alignment(horizontal='center', wrap_text=True)

    g7_1 = ws1['G7']
    g7_1.value = offer_3_name
    g7_1.font = Font(bold=True)
    g7_1.border = Border(bottom=thin)
    g7_1.alignment = Alignment(horizontal='center', wrap_text=True)

    h7_1 = ws1['H7']
    h7_1.value = offer_4_name
    h7_1.font = Font(bold=True)
    h7_1.border = Border(bottom=thin)
    h7_1.alignment = Alignment(horizontal='center', wrap_text=True)

    c8_1 = ws1['C8']
    c8_1.value = 'Offer Amt. ($)'
    # c8_1.font = Font(bold=True)
    c8_1.alignment = Alignment(horizontal='left', vertical='center')
    c8_1.border = Border(top=thin, bottom=hair, left=thin)

    d8_1 = ws1['D8']
    d8_1.border = Border(top=thin, bottom=hair)

    e8_1 = ws1['E8']
    e8_1.value = offer_1_amt
    # e8_1.font = Font(bold=True)
    e8_1.number_format = acct_fmt
    e8_1.border = Border(top=thin, bottom=hair, left=thin, right=thin)

    f8_1 = ws1['F8']
    f8_1.value = offer_2_amt
    # f8_1.font = Font(bold=True)
    f8_1.number_format = acct_fmt
    f8_1.border = Border(top=thin, bottom=hair, left=thin, right=thin)

    g8_1 = ws1['G8']
    g8_1.value = offer_3_amt
    # g8_1.font = Font(bold=True)
    g8_1.number_format = acct_fmt
    g8_1.border = Border(top=thin, bottom=hair, right=thin, left=thin)

    h8_1 = ws1['H8']
    h8_1.value = offer_4_amt
    # h8_1.font = Font(bold=True)
    h8_1.number_format = acct_fmt
    h8_1.border = Border(top=thin, bottom=hair, right=thin, left=thin)

    c9_1 = ws1['C9']
    c9_1.value = 'Down Pmt (%)'
    # c9_1.font = Font(bold=True)
    c9_1.alignment = Alignment(horizontal='left', vertical='center')
    c9_1.border = Border(top=hair, bottom=hair, left=thin)

    d9_1 = ws1['D9']
    d9_1.border = Border(top=hair, bottom=hair)

    e9_1 = ws1['E9']
    e9_1.value = offer_1_down_pmt_pct
    # e9_1.font = Font(bold=True)
    e9_1.number_format = pct_fmt
    e9_1.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    f9_1 = ws1['F9']
    f9_1.value = offer_2_down_pmt_pct
    # f9_1.font = Font(bold=True)
    f9_1.number_format = pct_fmt
    f9_1.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    g9_1 = ws1['G9']
    g9_1.value = offer_3_down_pmt_pct
    # g9_1.font = Font(bold=True)
    g9_1.number_format = pct_fmt
    g9_1.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    h9_1 = ws1['H9']
    h9_1.value = offer_4_down_pmt_pct
    # h9_1.font = Font(bold=True)
    h9_1.number_format = pct_fmt
    h9_1.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    c10_1 = ws1['C10']
    c10_1.value = 'Settlement Date'
    # c10_1.font = Font(bold=True)
    c10_1.alignment = Alignment(horizontal='left', vertical='center')
    c10_1.border = Border(top=hair, bottom=hair, left=thin)

    d10_1 = ws1['D10']
    d10_1.border = Border(top=hair, bottom=hair)

    e10_1 = ws1['E10']
    e10_1.value = offer_1_settlement_date
    # e10_1.font = Font(bold=True)
    e10_1.alignment = Alignment(horizontal='right')
    e10_1.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    f10_1 = ws1['F10']
    f10_1.value = offer_2_settlement_date
    # f10_1.font = Font(bold=True)
    f10_1.alignment = Alignment(horizontal='right')
    f10_1.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    g10_1 = ws1['G10']
    g10_1.value = offer_3_settlement_date
    # g10_1.font = Font(bold=True)
    g10_1.alignment = Alignment(horizontal='right')
    g10_1.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    h10_1 = ws1['H10']
    h10_1.value = offer_4_settlement_date
    # h10_1.font = Font(bold=True)
    h10_1.alignment = Alignment(horizontal='right')
    h10_1.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    c11_1 = ws1['C11']
    c11_1.value = 'Settlement Company'
    # c11_1.font = Font(bold=True)
    c11_1.alignment = Alignment(horizontal='left', vertical='center')
    c11_1.border = Border(top=hair, bottom=hair, left=thin)

    d11_1 = ws1['D11']
    d11_1.border = Border(top=hair, bottom=hair)

    e11_1 = ws1['E11']
    e11_1.value = offer_1_settlement_company
    # e11_1.font = Font(bold=True)
    e11_1.alignment = Alignment(horizontal='center', vertical='top', wrap_text=True)
    e11_1.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    f11_1 = ws1['F11']
    f11_1.value = offer_2_settlement_company
    # f11_1.font = Font(bold=True)
    f11_1.alignment = Alignment(horizontal='center', vertical='top', wrap_text=True)
    f11_1.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    g11_1 = ws1['G11']
    g11_1.value = offer_3_settlement_company
    # g11_1.font = Font(bold=True)
    g11_1.alignment = Alignment(horizontal='center', vertical='top', wrap_text=True)
    g11_1.border = Border(top=hair, bottom=hair, right=thin, left=thin)

    h11_1 = ws1['H11']
    h11_1.value = offer_4_settlement_company
    # h11_1.font = Font(bold=True)
    h11_1.alignment = Alignment(horizontal='center', vertical='top', wrap_text=True)
    h11_1.border = Border(top=hair, bottom=hair, right=thin, left=thin)

    c12_1 = ws1['C12']
    c12_1.value = 'EMD Amt. ($)'
    # c12_1.font = Font(bold=True)
    c12_1.alignment = Alignment(horizontal='left', vertical='center')
    c12_1.border = Border(top=hair, bottom=hair, left=thin)

    d12_1 = ws1['D12']
    d12_1.border = Border(top=hair, bottom=hair)

    e12_1 = ws1['E12']
    e12_1.value = offer_1_emd_amt
    # e12_1.font = Font(bold=True)
    e12_1.number_format = acct_fmt
    e12_1.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    f12_1 = ws1['F12']
    f12_1.value = offer_2_emd_amt
    # f12_1.font = Font(bold=True)
    f12_1.number_format = acct_fmt
    f12_1.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    g12_1 = ws1['G12']
    g12_1.value = offer_3_emd_amt
    # g12_1.font = Font(bold=True)
    g12_1.number_format = acct_fmt
    g12_1.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    h12_1 = ws1['H12']
    h12_1.value = offer_4_emd_amt
    # h12_1.font = Font(bold=True)
    h12_1.number_format = acct_fmt
    h12_1.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    c13_1 = ws1['C13']
    c13_1.value = 'Financing Type'
    # c13_1.font = Font(bold=True)
    c13_1.alignment = Alignment(horizontal='left', vertical='center')
    c13_1.border = Border(top=hair, bottom=hair, left=thin)

    d13_1 = ws1['D13']
    d13_1.border = Border(top=hair, bottom=hair)

    e13_1 = ws1['E13']
    e13_1.value = offer_1_financing_type
    # e13_1.font = Font(bold=True)
    e13_1.alignment = Alignment(horizontal='center')
    e13_1.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    f13_1 = ws1['F13']
    f13_1.value = offer_2_financing_type
    # f13_1.font = Font(bold=True)
    f13_1.alignment = Alignment(horizontal='center')
    f13_1.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    g13_1 = ws1['G13']
    g13_1.value = offer_3_financing_type
    # g13_1.font = Font(bold=True)
    g13_1.alignment = Alignment(horizontal='center')
    g13_1.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    h13_1 = ws1['H13']
    h13_1.value = offer_4_financing_type
    # h13_1.font = Font(bold=True)
    h13_1.alignment = Alignment(horizontal='center')
    h13_1.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    ws1.merge_cells('C14:C15')
    top_left_home_insp_cont_1 = ws1['C14']
    top_left_home_insp_cont_1.value = 'Home Inspection Contingency'
    # top_left_home_insp_cont_1.font = Font(bold=True)
    top_left_home_insp_cont_1.alignment = Alignment(horizontal='left', vertical='center')
    top_left_home_insp_cont_1.border = Border(top=hair, bottom=hair, left=thin)
    c14_1 = ws1['C14']
    c14_1.border = Border(top=hair, left=thin)
    d14_1 = ws1['D14']
    d14_1.border = Border(top=hair)
    c15_1 = ws1['C15']
    c15_1.border = Border(bottom=hair, left=thin)
    d15_1 = ws1['D15']
    d15_1.border = Border(bottom=hair)

    e14_1 = ws1['E14']
    e14_1.value = offer_1_home_inspection_check
    # e14_1.font = Font(bold=True)
    e14_1.alignment = Alignment(horizontal='center')
    e14_1.border = Border(top=hair, left=thin, right=thin)

    f14_1 = ws1['F14']
    f14_1.value = offer_2_home_inspection_check
    # f14_1.font = Font(bold=True)
    f14_1.alignment = Alignment(horizontal='center')
    f14_1.border = Border(top=hair, left=thin, right=thin)

    g14_1 = ws1['G14']
    g14_1.value = offer_3_home_inspection_check
    # g14_1.font = Font(bold=True)
    g14_1.alignment = Alignment(horizontal='center')
    g14_1.border = Border(top=hair, left=thin, right=thin)

    h14_1 = ws1['H14']
    h14_1.value = offer_4_home_inspection_check
    # h14_1.font = Font(bold=True)
    h14_1.alignment = Alignment(horizontal='center')
    h14_1.border = Border(top=hair, left=thin, right=thin)

    e15_1 = ws1['E15']
    e15_1.value = offer_1_home_inspection_days
    # e15_1.font = Font(bold=True)
    e15_1.alignment = Alignment(horizontal='center')
    e15_1.border = Border(bottom=hair, left=thin, right=thin)

    f15_1 = ws1['F15']
    f15_1.value = offer_2_home_inspection_days
    # f15_1.font = Font(bold=True)
    f15_1.alignment = Alignment(horizontal='center')
    f15_1.border = Border(bottom=hair, left=thin, right=thin)

    g15_1 = ws1['G15']
    g15_1.value = offer_3_home_inspection_days
    # g15_1.font = Font(bold=True)
    g15_1.alignment = Alignment(horizontal='center')
    g15_1.border = Border(bottom=hair, left=thin, right=thin)

    h15_1 = ws1['H15']
    h15_1.value = offer_4_home_inspection_days
    # h15_1.font = Font(bold=True)
    h15_1.alignment = Alignment(horizontal='center')
    h15_1.border = Border(bottom=hair, left=thin, right=thin)

    ws1.merge_cells('C16:C17')
    top_left_radon_insp_cont_1 = ws1['C16']
    top_left_radon_insp_cont_1.value = 'Radon Inspection Contingency'
    # top_left_radon_insp_cont_1.font = Font(bold=True)
    top_left_radon_insp_cont_1.alignment = Alignment(horizontal='left', vertical='center')
    top_left_radon_insp_cont_1.border = Border(top=hair, bottom=hair, left=thin)
    c16_1 = ws1['C16']
    c16_1.border = Border(top=hair, left=thin)
    d16_1 = ws1['D16']
    d16_1.border = Border(top=hair)
    c17_1 = ws1['C17']
    c17_1.border = Border(bottom=hair, left=thin)
    d17_1 = ws1['D17']
    d17_1.border = Border(bottom=hair)

    e16_1 = ws1['E16']
    e16_1.value = offer_1_radon_inspection_check
    # e16_1.font = Font(bold=True)
    e16_1.alignment = Alignment(horizontal='center')
    e16_1.border = Border(top=hair, left=thin, right=thin)

    f16_1 = ws1['F16']
    f16_1.value = offer_2_radon_inspection_check
    # f16_1.font = Font(bold=True)
    f16_1.alignment = Alignment(horizontal='center')
    f16_1.border = Border(top=hair, left=thin, right=thin)

    g16_1 = ws1['G16']
    g16_1.value = offer_3_radon_inspection_check
    # g16_1.font = Font(bold=True)
    g16_1.alignment = Alignment(horizontal='center')
    g16_1.border = Border(top=hair, left=thin, right=thin)

    h16_1 = ws1['H16']
    h16_1.value = offer_4_radon_inspection_check
    # h16_1.font = Font(bold=True)
    h16_1.alignment = Alignment(horizontal='center')
    h16_1.border = Border(top=hair, left=thin, right=thin)

    e17_1 = ws1['E17']
    e17_1.value = offer_1_radon_inspection_days
    # e17_1.font = Font(bold=True)
    e17_1.alignment = Alignment(horizontal='center')
    e17_1.border = Border(bottom=hair, left=thin, right=thin)

    f17_1 = ws1['F17']
    f17_1.value = offer_2_radon_inspection_days
    # f17_1.font = Font(bold=True)
    f17_1.alignment = Alignment(horizontal='center')
    f17_1.border = Border(bottom=hair, left=thin, right=thin)

    g17_1 = ws1['G17']
    g17_1.value = offer_3_radon_inspection_days
    # g17_1.font = Font(bold=True)
    g17_1.alignment = Alignment(horizontal='center')
    g17_1.border = Border(bottom=hair, left=thin, right=thin)

    h17_1 = ws1['H17']
    h17_1.value = offer_4_radon_inspection_days
    # h17_1.font = Font(bold=True)
    h17_1.alignment = Alignment(horizontal='center')
    h17_1.border = Border(bottom=hair, left=thin, right=thin)

    ws1.merge_cells('C18:C19')
    top_left_septic_insp_cont_1 = ws1['C18']
    top_left_septic_insp_cont_1.value = 'Septic Inspection Contingency'
    # top_left_septic_insp_cont_1.font = Font(bold=True)
    top_left_septic_insp_cont_1.alignment = Alignment(horizontal='left', vertical='center')
    top_left_septic_insp_cont_1.border = Border(top=hair, bottom=hair, left=thin)
    c18_1 = ws1['C18']
    c18_1.border = Border(top=hair, left=thin)
    d18_1 = ws1['D18']
    d18_1.border = Border(top=hair)
    c19_1 = ws1['C19']
    c19_1.border = Border(bottom=hair, left=thin)
    d19_1 = ws1['D19']
    d19_1.border = Border(bottom=hair)

    e18_1 = ws1['E18']
    e18_1.value = offer_1_septic_inspection_check
    # e18_1.font = Font(bold=True)
    e18_1.alignment = Alignment(horizontal='center')
    e18_1.border = Border(top=hair, left=thin, right=thin)

    f18_1 = ws1['F18']
    f18_1.value = offer_2_septic_inspection_check
    # f18_1.font = Font(bold=True)
    f18_1.alignment = Alignment(horizontal='center')
    f18_1.border = Border(top=hair, left=thin, right=thin)

    g18_1 = ws1['G18']
    g18_1.value = offer_3_septic_inspection_check
    # g18_1.font = Font(bold=True)
    g18_1.alignment = Alignment(horizontal='center')
    g18_1.border = Border(top=hair, left=thin, right=thin)

    h18_1 = ws1['H18']
    h18_1.value = offer_4_septic_inspection_check
    # h18_1.font = Font(bold=True)
    h18_1.alignment = Alignment(horizontal='center')
    h18_1.border = Border(top=hair, left=thin, right=thin)

    e19_1 = ws1['E19']
    e19_1.value = offer_1_septic_inspection_days
    # e19_1.font = Font(bold=True)
    e19_1.alignment = Alignment(horizontal='center')
    e19_1.border = Border(bottom=hair, left=thin, right=thin)

    f19_1 = ws1['F19']
    f19_1.value = offer_2_septic_inspection_days
    # f19_1.font = Font(bold=True)
    f19_1.alignment = Alignment(horizontal='center')
    f19_1.border = Border(bottom=hair, left=thin, right=thin)

    g19_1 = ws1['G19']
    g19_1.value = offer_3_septic_inspection_days
    # g19_1.font = Font(bold=True)
    g19_1.alignment = Alignment(horizontal='center')
    g19_1.border = Border(bottom=hair, left=thin, right=thin)

    h19_1 = ws1['H19']
    h19_1.value = offer_4_septic_inspection_days
    # h19_1.font = Font(bold=True)
    h19_1.alignment = Alignment(horizontal='center')
    h19_1.border = Border(bottom=hair, left=thin, right=thin)

    ws1.merge_cells('C20:C21')
    top_left_well_insp_cont_1 = ws1['C20']
    top_left_well_insp_cont_1.value = 'Well Inspection Contingency'
    # top_left_well_insp_cont_1.font = Font(bold=True)
    top_left_well_insp_cont_1.alignment = Alignment(horizontal='left', vertical='center')
    top_left_well_insp_cont_1.border = Border(top=hair, bottom=hair, left=thin)
    c20_1 = ws1['C20']
    c20_1.border = Border(top=hair, left=thin)
    d20_1 = ws1['D20']
    d20_1.border = Border(top=hair)
    c21_1 = ws1['C21']
    c21_1.border = Border(bottom=hair, left=thin)
    d21_1 = ws1['D21']
    d21_1.border = Border(bottom=hair)

    e20_1 = ws1['E20']
    e20_1.value = offer_1_well_inspection_check
    # e20_1.font = Font(bold=True)
    e20_1.alignment = Alignment(horizontal='center')
    e20_1.border = Border(top=hair, left=thin, right=thin)

    f20_1 = ws1['F20']
    f20_1.value = offer_2_well_inspection_check
    # f20_1.font = Font(bold=True)
    f20_1.alignment = Alignment(horizontal='center')
    f20_1.border = Border(top=hair, left=thin, right=thin)

    g20_1 = ws1['G20']
    g20_1.value = offer_3_well_inspection_check
    # g20_1.font = Font(bold=True)
    g20_1.alignment = Alignment(horizontal='center')
    g20_1.border = Border(top=hair, left=thin, right=thin)

    h20_1 = ws1['H20']
    h20_1.value = offer_4_well_inspection_check
    # h20_1.font = Font(bold=True)
    h20_1.alignment = Alignment(horizontal='center')
    h20_1.border = Border(top=hair, left=thin, right=thin)

    e21_1 = ws1['E21']
    e21_1.value = offer_1_well_inspection_days
    # e21_1.font = Font(bold=True)
    e21_1.alignment = Alignment(horizontal='center')
    e21_1.border = Border(bottom=hair, left=thin, right=thin)

    f21_1 = ws1['F21']
    f21_1.value = offer_2_well_inspection_days
    # f21_1.font = Font(bold=True)
    f21_1.alignment = Alignment(horizontal='center')
    f21_1.border = Border(bottom=hair, left=thin, right=thin)

    g21_1 = ws1['G21']
    g21_1.value = offer_3_well_inspection_days
    # g21_1.font = Font(bold=True)
    g21_1.alignment = Alignment(horizontal='center')
    g21_1.border = Border( bottom=hair, left=thin, right=thin)

    h21_1 = ws1['H21']
    h21_1.value = offer_4_well_inspection_days
    # h21_1.font = Font(bold=True)
    h21_1.alignment = Alignment(horizontal='center')
    h21_1.border = Border(bottom=hair, left=thin, right=thin)

    ws1.merge_cells('C22:C23')
    top_left_finance_cont_1 = ws1['C22']
    top_left_finance_cont_1.value = 'Finance Contingency'
    # top_left_finance_cont_1.font = Font(bold=True)
    top_left_finance_cont_1.alignment = Alignment(horizontal='left', vertical='center')
    top_left_finance_cont_1.border = Border(top=hair, bottom=hair, left=thin)
    c22_1 = ws1['C22']
    c22_1.border = Border(top=hair, left=thin)
    d22_1 = ws1['D22']
    d22_1.border = Border(top=hair)
    c23_1 = ws1['C23']
    c23_1.border = Border(bottom=hair, left=thin)
    d23_1 = ws1['D23']
    d23_1.border = Border(bottom=hair)

    e22_1 = ws1['E22']
    e22_1.value = offer_1_finance_contingency_check
    # e22_1.font = Font(bold=True)
    e22_1.alignment = Alignment(horizontal='center')
    e22_1.border = Border(top=hair, left=thin, right=thin)

    f22_1 = ws1['F22']
    f22_1.value = offer_2_finance_contingency_check
    # f22_1.font = Font(bold=True)
    f22_1.alignment = Alignment(horizontal='center')
    f22_1.border = Border(top=hair, left=thin, right=thin)

    g22_1 = ws1['G22']
    g22_1.value = offer_3_finance_contingency_check
    # g22_1.font = Font(bold=True)
    g22_1.alignment = Alignment(horizontal='center')
    g22_1.border = Border(top=hair, left=thin, right=thin)

    h22_1 = ws1['H22']
    h22_1.value = offer_4_finance_contingency_check
    # h22_1.font = Font(bold=True)
    h22_1.alignment = Alignment(horizontal='center')
    h22_1.border = Border(top=hair, left=thin, right=thin)

    e23_1 = ws1['E23']
    e23_1.value = offer_1_finance_contingency_days
    # e23_1.font = Font(bold=True)
    e23_1.alignment = Alignment(horizontal='center')
    e23_1.border = Border(bottom=hair, left=thin, right=thin)

    f23_1 = ws1['F23']
    f23_1.value = offer_2_finance_contingency_days
    # f23_1.font = Font(bold=True)
    f23_1.alignment = Alignment(horizontal='center')
    f23_1.border = Border(bottom=hair, left=thin, right=thin)

    g23_1 = ws1['G23']
    g23_1.value = offer_3_finance_contingency_days
    # g23_1.font = Font(bold=True)
    g23_1.alignment = Alignment(horizontal='center')
    g23_1.border = Border(bottom=hair, left=thin, right=thin)

    h23_1 = ws1['H23']
    h23_1.value = offer_4_finance_contingency_days
    # h23_1.font = Font(bold=True)
    h23_1.alignment = Alignment(horizontal='center')
    h23_1.border = Border(bottom=hair, left=thin, right=thin)

    ws1.merge_cells('C24:C25')
    top_left_appraisal_cont_1 = ws1['C24']
    top_left_appraisal_cont_1.value = 'Appraisal Contingency'
    # top_left_appraisal_cont_1.font = Font(bold=True)
    top_left_appraisal_cont_1.alignment = Alignment(horizontal='left', vertical='center')
    top_left_appraisal_cont_1.border = Border(top=hair, bottom=hair, left=thin)
    c24_1 = ws1['C24']
    c24_1.border = Border(top=hair, left=thin)
    d24_1 = ws1['D24']
    d24_1.border = Border(top=hair)
    c25_1 = ws1['C25']
    c25_1.border = Border(bottom=hair, left=thin)
    d25_1 = ws1['D25']
    d25_1.border = Border(bottom=hair)

    e24_1 = ws1['E24']
    e24_1.value = offer_1_appraisal_contingency_check
    # e24_1.font = Font(bold=True)
    e24_1.alignment = Alignment(horizontal='center')
    e24_1.border = Border(top=hair, left=thin, right=thin)

    f24_1 = ws1['F24']
    f24_1.value = offer_2_appraisal_contingency_check
    # f24_1.font = Font(bold=True)
    f24_1.alignment = Alignment(horizontal='center')
    f24_1.border = Border(top=hair, left=thin, right=thin)

    g24_1 = ws1['G24']
    g24_1.value = offer_3_appraisal_contingency_check
    # g24_1.font = Font(bold=True)
    g24_1.alignment = Alignment(horizontal='center')
    g24_1.border = Border(top=hair, left=thin, right=thin)

    h24_1 = ws1['H24']
    h24_1.value = offer_4_appraisal_contingency_check
    # h24_1.font = Font(bold=True)
    h24_1.alignment = Alignment(horizontal='center')
    h24_1.border = Border(top=hair, left=thin, right=thin)

    e25_1 = ws1['E25']
    e25_1.value = offer_1_appraisal_contingency_days
    # e25_1.font = Font(bold=True)
    e25_1.alignment = Alignment(horizontal='center')
    e25_1.border = Border(bottom=hair, left=thin, right=thin)

    f25_1 = ws1['F25']
    f25_1.value = offer_2_appraisal_contingency_days
    # f25_1.font = Font(bold=True)
    f25_1.alignment = Alignment(horizontal='center')
    f25_1.border = Border(bottom=hair, left=thin, right=thin)

    g25_1 = ws1['G25']
    g25_1.value = offer_3_appraisal_contingency_days
    # g25_1.font = Font(bold=True)
    g25_1.alignment = Alignment(horizontal='center')
    g25_1.border = Border(bottom=hair, left=thin, right=thin)

    h25_1 = ws1['H25']
    h25_1.value = offer_4_appraisal_contingency_days
    # h25_1.font = Font(bold=True)
    h25_1.alignment = Alignment(horizontal='center')
    h25_1.border = Border(bottom=hair, left=thin, right=thin)

    ws1.merge_cells('C26:C27')
    top_left_home_sale_cont_1 = ws1['C26']
    top_left_home_sale_cont_1.value = 'Home Sale Contingency'
    # top_left_home_sale_cont_1.font = Font(bold=True)
    top_left_home_sale_cont_1.alignment = Alignment(horizontal='left', vertical='center')
    top_left_home_sale_cont_1.border = Border(top=hair, bottom=hair, left=thin)
    c26_1 = ws1['C26']
    c26_1.border = Border(top=hair, left=thin)
    d26_1 = ws1['D26']
    d26_1.border = Border(top=hair)
    c27_1 = ws1['C27']
    c27_1.border = Border(bottom=hair, left=thin)
    d27_1 = ws1['D27']
    d27_1.border = Border(bottom=hair)

    e26_1 = ws1['E26']
    e26_1.value = offer_1_home_sale_contingency_check
    # e26_1.font = Font(bold=True)
    e26_1.alignment = Alignment(horizontal='center')
    e26_1.border = Border(top=hair, left=thin, right=thin)

    f26_1 = ws1['F26']
    f26_1.value = offer_2_home_sale_contingency_check
    # f26_1.font = Font(bold=True)
    f26_1.alignment = Alignment(horizontal='center')
    f26_1.border = Border(top=hair, left=thin, right=thin)

    g26_1 = ws1['G26']
    g26_1.value = offer_3_home_sale_contingency_check
    # g26_1.font = Font(bold=True)
    g26_1.alignment = Alignment(horizontal='center')
    g26_1.border = Border(top=hair, left=thin, right=thin)

    h26_1 = ws1['H26']
    h26_1.value = offer_4_home_sale_contingency_check
    # h26_1.font = Font(bold=True)
    h26_1.alignment = Alignment(horizontal='center')
    h26_1.border = Border(top=hair, left=thin, right=thin)

    e27_1 = ws1['E27']
    e27_1.value = offer_1_home_sale_contingency_days
    # e27_1.font = Font(bold=True)
    e27_1.alignment = Alignment(horizontal='center')
    e27_1.border = Border(bottom=hair, left=thin, right=thin)

    f27_1 = ws1['F27']
    f27_1.value = offer_2_home_sale_contingency_days
    # f27_1.font = Font(bold=True)
    f27_1.alignment = Alignment(horizontal='center')
    f27_1.border = Border(bottom=hair, left=thin, right=thin)

    g27_1 = ws1['G27']
    g27_1.value = offer_3_home_sale_contingency_days
    # g27_1.font = Font(bold=True)
    g27_1.alignment = Alignment(horizontal='center')
    g27_1.border = Border(bottom=hair, left=thin, right=thin)

    h27_1 = ws1['H27']
    h27_1.value = offer_4_home_sale_contingency_days
    # h27_1.font = Font(bold=True)
    h27_1.alignment = Alignment(horizontal='center')
    h27_1.border = Border(bottom=hair, left=thin, right=thin)

    c28_1 = ws1['C28']
    c28_1.value = 'Pre Occupancy Start Date'
    # c28_1.font = Font(bold=True)
    c28_1.alignment = Alignment(horizontal='left', vertical='center')
    c28_1.border = Border(top=hair, bottom=hair, left=thin)

    d28_1 = ws1['D28']
    d28_1.border = Border(top=hair, bottom=hair)

    e28_1 = ws1['E28']
    e28_1.value = offer_1_pre_occupancy_date
    # e28_1.font = Font(bold=True)
    e28_1.alignment = Alignment(horizontal='right')
    e28_1.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    f28_1 = ws1['F28']
    f28_1.value = offer_2_pre_occupancy_date
    # f28_1.font = Font(bold=True)
    f28_1.alignment = Alignment(horizontal='right')
    f28_1.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    g28_1 = ws1['G28']
    g28_1.value = offer_3_pre_occupancy_date
    # g28_1.font = Font(bold=True)
    g28_1.alignment = Alignment(horizontal='right')
    g28_1.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    h28_1 = ws1['H28']
    h28_1.value = offer_4_pre_occupancy_date
    # h28_1.font = Font(bold=True)
    h28_1.alignment = Alignment(horizontal='right')
    h28_1.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    c29_1 = ws1['C29']
    c29_1.value = 'Post Occupancy Thru Date'
    # c29_1.font = Font(bold=True)
    c29_1.alignment = Alignment(horizontal='left', vertical='center')
    c29_1.border = Border(top=hair, bottom=thin, left=thin)

    d29_1 = ws1['D29']
    d29_1.border = Border(top=hair, bottom=thin)

    e29_1 = ws1['E29']
    e29_1.value = offer_1_post_occupancy_date
    # e29_1.font = Font(bold=True)
    e29_1.alignment = Alignment(horizontal='right')
    e29_1.border = Border(top=hair, bottom=thin, left=thin, right=thin)

    f29_1 = ws1['F29']
    f29_1.value = offer_2_post_occupancy_date
    # f29_1.font = Font(bold=True)
    f29_1.alignment = Alignment(horizontal='right')
    f29_1.border = Border(top=hair, bottom=thin, left=thin, right=thin)

    g29_1 = ws1['G29']
    g29_1.value = offer_3_post_occupancy_date
    # g29_1.font = Font(bold=True)
    g29_1.alignment = Alignment(horizontal='right')
    g29_1.border = Border(top=hair, bottom=thin, left=thin, right=thin)

    h29_1 = ws1['H29']
    h29_1.value = offer_4_post_occupancy_date
    # h29_1.font = Font(bold=True)
    h29_1.alignment = Alignment(horizontal='right')
    h29_1.border = Border(top=hair, bottom=thin, left=thin, right=thin)

    # Housing Related Cost Table
    c31_1 = ws1['C31']
    c31_1.value = 'HOUSING-RELATED COSTS'
    c31_1.font = Font(bold=True)
    c31_1.border = Border(bottom=thin)

    d31_1 = ws1['D31']
    d31_1.value = 'Calculation Description'
    d31_1.font = Font(bold=True)
    d31_1.border = Border(bottom=thin)

    c32_1 = ws1['C32']
    c32_1.value = 'Estimated Payoff - 1st Trust'
    c32_1.border = Border(top=thin, bottom=hair, left=thin)

    d32_1 = ws1['D32']
    d32_1.value = 'Principal Balance of Loan'
    d32_1.border = Border(top=thin, bottom=hair)

    e32_1 = ws1['E32']
    e32_1.value = first_trust
    e32_1.number_format = acct_fmt
    e32_1.border = Border(top=thin, bottom=hair, left=thin, right=thin)

    f32_1 = ws1['F32']
    f32_1.value = first_trust
    f32_1.number_format = acct_fmt
    f32_1.border = Border(top=thin, bottom=hair, left=thin, right=thin)

    g32_1 = ws1['G32']
    g32_1.value = first_trust
    g32_1.number_format = acct_fmt
    g32_1.border = Border(top=thin, bottom=hair, left=thin, right=thin)

    h32_1 = ws1['H32']
    h32_1.value = first_trust
    h32_1.number_format = acct_fmt
    h32_1.border = Border(top=thin, bottom=hair, left=thin, right=thin)

    c33_1 = ws1['C33']
    c33_1.value = 'Estimated Payoff - 2nd Trust'
    c33_1.border = Border(top=hair, bottom=hair, left=thin)

    d33_1 = ws1['D33']
    d33_1.value = 'Principal Balance of Loan'
    d33_1.border = Border(top=hair, bottom=hair)

    e33_1 = ws1['E33']
    e33_1.value = second_trust
    e33_1.number_format = acct_fmt
    e33_1.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    f33_1 = ws1['F33']
    f33_1.value = second_trust
    f33_1.number_format = acct_fmt
    f33_1.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    g33_1 = ws1['G33']
    g33_1.value = second_trust
    g33_1.number_format = acct_fmt
    g33_1.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    h33_1 = ws1['H33']
    h33_1.value = second_trust
    h33_1.number_format = acct_fmt
    h33_1.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    c34_1 = ws1['C34']
    c34_1.value = 'Purchaser Closing Cost / Contract'
    c34_1.border = Border(top=hair, bottom=hair, left=thin)

    d34_1 = ws1['D34']
    d34_1.value = 'Negotiated Into Contract'
    d34_1.border = Border(top=hair, bottom=hair)

    e34_1 = ws1['E34']
    e34_1.value = offer_1_closing_cost_subsidy_amt
    e34_1.number_format = acct_fmt
    e34_1.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    f34_1 = ws1['F34']
    f34_1.value = offer_2_closing_cost_subsidy_amt
    f34_1.number_format = acct_fmt
    f34_1.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    g34_1 = ws1['G34']
    g34_1.value = offer_3_closing_cost_subsidy_amt
    g34_1.number_format = acct_fmt
    g34_1.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    h34_1 = ws1['H34']
    h34_1.value = offer_4_closing_cost_subsidy_amt
    h34_1.number_format = acct_fmt
    h34_1.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    c35_1 = ws1['C35']
    c35_1.value = 'Prorated Taxes / Assessments'
    c35_1.border = Border(top=hair, bottom=hair, left=thin)

    d35_1 = ws1['D35']
    d35_1.value = '1 Year of Taxes Divided by 12 Multiplied by 3'
    d35_1.border = Border(top=hair, bottom=hair)

    e35_1 = ws1['E35']
    e35_1.value = prorated_taxes
    e35_1.number_format = acct_fmt
    e35_1.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    f35_1 = ws1['F35']
    f35_1.value = prorated_taxes
    f35_1.number_format = acct_fmt
    f35_1.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    g35_1 = ws1['G35']
    g35_1.value = prorated_taxes
    g35_1.number_format = acct_fmt
    g35_1.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    h35_1 = ws1['H35']
    h35_1.value = prorated_taxes
    h35_1.number_format = acct_fmt
    h35_1.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    c36_1 = ws1['C36']
    c36_1.value = 'Prorated HOA / Condo Dues'
    c36_1.border = Border(top=hair, bottom=thin, left=thin)

    d36_1 = ws1['D36']
    d36_1.value = '1 Year of Dues Divided by 12 Multiplied by 3'
    d36_1.border = Border(top=hair, bottom=thin)

    e36_1 = ws1['E36']
    e36_1.value = prorated_hoa_condo_fees
    e36_1.number_format = acct_fmt
    e36_1.border = Border(top=hair, bottom=thin, left=thin, right=thin)

    f36_1 = ws1['F36']
    f36_1.value = prorated_hoa_condo_fees
    f36_1.number_format = acct_fmt
    f36_1.border = Border(top=hair, bottom=thin, left=thin, right=thin)

    g36_1 = ws1['G36']
    g36_1.value = prorated_hoa_condo_fees
    g36_1.number_format = acct_fmt
    g36_1.border = Border(top=hair, bottom=thin, left=thin, right=thin)

    h36_1 = ws1['H36']
    h36_1.value = prorated_hoa_condo_fees
    h36_1.number_format = acct_fmt
    h36_1.border = Border(top=hair, bottom=thin, left=thin, right=thin)

    e37_1 = ws1['E37']
    e37_1.value = '=SUM(E32:E36)'
    e37_1.font = Font(bold=True)
    e37_1.number_format = acct_fmt
    e37_1.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    f37_1 = ws1['F37']
    f37_1.value = '=SUM(F32:F36)'
    f37_1.font = Font(bold=True)
    f37_1.number_format = acct_fmt
    f37_1.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    g37_1 = ws1['G37']
    g37_1.value = '=SUM(G32:G36)'
    g37_1.font = Font(bold=True)
    g37_1.number_format = acct_fmt
    g37_1.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    h37_1 = ws1['H37']
    h37_1.value = '=SUM(H32:H36)'
    h37_1.font = Font(bold=True)
    h37_1.number_format = acct_fmt
    h37_1.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    # Broker and Financing Costs Table
    c38_1 = ws1['C38']
    c38_1.value = 'BROKERAGE & FINANCING COSTS'
    c38_1.font = Font(bold=True)
    c38_1.border = Border(bottom=thin)

    c39_1 = ws1['C39']
    c39_1.value = 'Listing Company Compensation'
    c39_1.border = Border(top=thin, bottom=hair, left=thin)

    d39_1 = ws1['D39']
    d39_1.value = '% from Listing Agreement * Offer Amount ($)'
    d39_1.border = Border(top=thin, bottom=hair)

    e39_1 = ws1['E39']
    e39_1.value = listing_company_pct * offer_1_amt
    e39_1.number_format = acct_fmt
    e39_1.border = Border(top=thin, bottom=hair, left=thin, right=thin)

    f39_1 = ws1['F39']
    f39_1.value = listing_company_pct * offer_2_amt
    f39_1.number_format = acct_fmt
    f39_1.border = Border(top=thin, bottom=hair, left=thin, right=thin)

    g39_1 = ws1['G39']
    g39_1.value = listing_company_pct * offer_3_amt
    g39_1.number_format = acct_fmt
    g39_1.border = Border(top=thin, bottom=hair, left=thin, right=thin)

    h39_1 = ws1['H39']
    h39_1.value = listing_company_pct * offer_4_amt
    h39_1.number_format = acct_fmt
    h39_1.border = Border(top=thin, bottom=hair, left=thin, right=thin)

    c40_1 = ws1['C40']
    c40_1.value = 'Selling Company Compensation'
    c40_1.border = Border(top=hair, bottom=hair, left=thin)

    d40_1 = ws1['D40']
    d40_1.value = '% from Listing Agreement * Offer Amount ($)'
    d40_1.border = Border(top=hair, bottom=hair)

    e40_1 = ws1['E40']
    e40_1.value = selling_company_pct * offer_1_amt
    e40_1.number_format = acct_fmt
    e40_1.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    f40_1 = ws1['F40']
    f40_1.value = selling_company_pct * offer_2_amt
    f40_1.number_format = acct_fmt
    f40_1.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    g40_1 = ws1['G40']
    g40_1.value = selling_company_pct * offer_3_amt
    g40_1.number_format = acct_fmt
    g40_1.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    h40_1 = ws1['H40']
    h40_1.value = selling_company_pct * offer_4_amt
    h40_1.number_format = acct_fmt
    h40_1.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    c41_1 = ws1['C41']
    c41_1.value = 'Processing Fee'
    c41_1.border = Border(top=hair, bottom=thin, left=thin)

    d41_1 = ws1['D41']
    d41_1.border = Border(top=hair, bottom=thin)

    e41_1 = ws1['E41']
    e41_1.value = processing_fee
    e41_1.number_format = acct_fmt
    e41_1.border = Border(top=hair, bottom=thin, left=thin, right=thin)

    f41_1 = ws1['F41']
    f41_1.value = processing_fee
    f41_1.number_format = acct_fmt
    f41_1.border = Border(top=hair, bottom=thin, left=thin, right=thin)

    g41_1 = ws1['G41']
    g41_1.value = processing_fee
    g41_1.number_format = acct_fmt
    g41_1.border = Border(top=hair, bottom=thin, left=thin, right=thin)

    h41_1 = ws1['H41']
    h41_1.value = processing_fee
    h41_1.number_format = acct_fmt
    h41_1.border = Border(top=hair, bottom=thin, left=thin, right=thin)

    e42_1 = ws1['E42']
    e42_1.value = '=SUM(E39:E41)'
    e42_1.font = Font(bold=True)
    e42_1.number_format = acct_fmt
    e42_1.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    f42_1 = ws1['F42']
    f42_1.value = '=SUM(F39:F41)'
    f42_1.font = Font(bold=True)
    f42_1.number_format = acct_fmt
    f42_1.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    g42_1 = ws1['G42']
    g42_1.value = '=SUM(G39:G41)'
    g42_1.font = Font(bold=True)
    g42_1.number_format = acct_fmt
    g42_1.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    h42_1 = ws1['H42']
    h42_1.value = '=SUM(H39:H41)'
    h42_1.font = Font(bold=True)
    h42_1.number_format = acct_fmt
    h42_1.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    # Estimated Closing Costs Table
    c43_1 = ws1['C43']
    c43_1.value = 'ESTIMATED CLOSING COSTS'
    c43_1.font = Font(bold=True)
    c43_1.border = Border(bottom=thin)

    c44_1 = ws1['C44']
    c44_1.value = 'Settlement Fee'
    c44_1.border = Border(top=thin, bottom=hair, left=thin)

    d44_1 = ws1['D44']
    d44_1.value = 'Commonly Used Fee'
    d44_1.border = Border(top=thin, bottom=hair)

    e44_1 = ws1['E44']
    e44_1.value = settlement_fee
    e44_1.number_format = acct_fmt
    e44_1.border = Border(top=thin, bottom=hair, left=thin, right=thin)

    f44_1 = ws1['F44']
    f44_1.value = settlement_fee
    f44_1.number_format = acct_fmt
    f44_1.border = Border(top=thin, bottom=hair, left=thin, right=thin)

    g44_1 = ws1['G44']
    g44_1.value = settlement_fee
    g44_1.number_format = acct_fmt
    g44_1.border = Border(top=thin, bottom=hair, left=thin, right=thin)

    h44_1 = ws1['H44']
    h44_1.value = settlement_fee
    h44_1.number_format = acct_fmt
    h44_1.border = Border(top=thin, bottom=hair, left=thin, right=thin)

    c45_1 = ws1['C45']
    c45_1.value = 'Deed Preparation'
    c45_1.border = Border(top=hair, bottom=hair, left=thin)

    d45_1 = ws1['D45']
    d45_1.value = 'Commonly Used Fee'
    d45_1.border = Border(top=hair, bottom=hair)

    e45_1 = ws1['E45']
    e45_1.value = deed_preparation_fee
    e45_1.number_format = acct_fmt
    e45_1.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    f45_1 = ws1['F45']
    f45_1.value = deed_preparation_fee
    f45_1.number_format = acct_fmt
    f45_1.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    g45_1 = ws1['G45']
    g45_1.value = deed_preparation_fee
    g45_1.number_format = acct_fmt
    g45_1.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    h45_1 = ws1['H45']
    h45_1.value = deed_preparation_fee
    h45_1.number_format = acct_fmt
    h45_1.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    c46_1 = ws1['C46']
    c46_1.value = 'Release of Liens / Trusts'
    c46_1.border = Border(top=hair, bottom=thin, left=thin)

    d46_1 = ws1['D46']
    d46_1.value = 'Commonly Used Fee * Qty of Trusts or Liens'
    d46_1.border = Border(top=hair, bottom=thin)

    e46_1 = ws1['E46']
    e46_1.value = lien_trust_release_fee * lien_trust_release_qty
    e46_1.number_format = acct_fmt
    e46_1.border = Border(top=hair, bottom=thin, left=thin, right=thin)

    f46_1 = ws1['F46']
    f46_1.value = lien_trust_release_fee * lien_trust_release_qty
    f46_1.number_format = acct_fmt
    f46_1.border = Border(top=hair, bottom=thin, left=thin, right=thin)

    g46_1 = ws1['G46']
    g46_1.value = lien_trust_release_fee * lien_trust_release_qty
    g46_1.number_format = acct_fmt
    g46_1.border = Border(top=hair, bottom=thin, left=thin, right=thin)

    h46_1 = ws1['H46']
    h46_1.value = lien_trust_release_fee * lien_trust_release_qty
    h46_1.number_format = acct_fmt
    h46_1.border = Border(top=hair, bottom=thin, left=thin, right=thin)

    e47_1 = ws1['E47']
    e47_1.value = '=SUM(E44:E46)'
    e47_1.font = Font(bold=True)
    e47_1.number_format = acct_fmt
    e47_1.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    f47_1 = ws1['F47']
    f47_1.value = '=SUM(F44:F46)'
    f47_1.font = Font(bold=True)
    f47_1.number_format = acct_fmt
    f47_1.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    g47_1 = ws1['G47']
    g47_1.value = '=SUM(G44:G46)'
    g47_1.font = Font(bold=True)
    g47_1.number_format = acct_fmt
    g47_1.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    h47_1 = ws1['H47']
    h47_1.value = '=SUM(H44:H46)'
    h47_1.font = Font(bold=True)
    h47_1.number_format = acct_fmt
    h47_1.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    # Misc Costs Table
    c48_1 = ws1['C48']
    c48_1.value = 'MISCELLANEOUS COSTS'
    c48_1.font = Font(bold=True)
    c48_1.border = Border(bottom=thin)

    c49_1 = ws1['C49']
    c49_1.value = 'Recording Release(s)'
    c49_1.border = Border(top=thin, bottom=hair, left=thin)

    d49_1 = ws1['D49']
    d49_1.value = 'Commonly Used Fee * Qty of Trusts Recorded'
    d49_1.border = Border(top=thin, bottom=hair)

    e49_1 = ws1['E49']
    e49_1.value = recording_fee * recording_trusts_liens_qty
    e49_1.number_format = acct_fmt
    e49_1.border = Border(top=thin, bottom=hair, left=thin, right=thin)

    f49_1 = ws1['F49']
    f49_1.value = recording_fee * recording_trusts_liens_qty
    f49_1.number_format = acct_fmt
    f49_1.border = Border(top=thin, bottom=hair, left=thin, right=thin)

    g49_1 = ws1['G49']
    g49_1.value = recording_fee * recording_trusts_liens_qty
    g49_1.number_format = acct_fmt
    g49_1.border = Border(top=thin, bottom=hair, left=thin, right=thin)

    h49_1 = ws1['H49']
    h49_1.value = recording_fee * recording_trusts_liens_qty
    h49_1.number_format = acct_fmt
    h49_1.border = Border(top=thin, bottom=hair, left=thin, right=thin)

    c50_1 = ws1['C50']
    c50_1.value = 'Grantor\'s Tax'
    c50_1.border = Border(top=hair, bottom=hair, left=thin)

    d50_1 = ws1['D50']
    d50_1.value = '% of Offer Amount ($)'
    d50_1.border = Border(top=hair, bottom=hair)

    e50_1 = ws1['E50']
    e50_1.value = grantors_tax_pct * offer_1_amt
    e50_1.number_format = acct_fmt
    e50_1.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    f50_1 = ws1['F50']
    f50_1.value = grantors_tax_pct * offer_2_amt
    f50_1.number_format = acct_fmt
    f50_1.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    g50_1 = ws1['G50']
    g50_1.value = grantors_tax_pct * offer_3_amt
    g50_1.number_format = acct_fmt
    g50_1.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    h50_1 = ws1['H50']
    h50_1.value = grantors_tax_pct * offer_4_amt
    h50_1.number_format = acct_fmt
    h50_1.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    c51_1 = ws1['C51']
    c51_1.value = 'Congestion Relief Tax'
    c51_1.border = Border(top=hair, bottom=hair, left=thin)

    d51_1 = ws1['D51']
    d51_1.value = '% of Offer Amount ($)'
    d51_1.border = Border(top=hair, bottom=hair)

    e51_1 = ws1['E51']
    e51_1.value = congestion_tax_pct * offer_1_amt
    e51_1.number_format = acct_fmt
    e51_1.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    f51_1 = ws1['F51']
    f51_1.value = congestion_tax_pct * offer_2_amt
    f51_1.number_format = acct_fmt
    f51_1.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    g51_1 = ws1['G51']
    g51_1.value = congestion_tax_pct * offer_3_amt
    g51_1.number_format = acct_fmt
    g51_1.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    h51_1 = ws1['H51']
    h51_1.value = congestion_tax_pct * offer_4_amt
    h51_1.number_format = acct_fmt
    h51_1.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    c52_1 = ws1['C52']
    c52_1.value = 'Pest Inspection'
    c52_1.border = Border(top=hair, bottom=hair, left=thin)

    d52_1 = ws1['D52']
    d52_1.value = 'Commonly Used Fee'
    d52_1.border = Border(top=hair, bottom=hair)

    e52_1 = ws1['E52']
    e52_1.value = pest_inspection_fee
    e52_1.number_format = acct_fmt
    e52_1.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    f52_1 = ws1['F52']
    f52_1.value = pest_inspection_fee
    f52_1.number_format = acct_fmt
    f52_1.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    g52_1 = ws1['G52']
    g52_1.value = pest_inspection_fee
    g52_1.number_format = acct_fmt
    g52_1.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    h52_1 = ws1['H52']
    h52_1.value = pest_inspection_fee
    h52_1.number_format = acct_fmt
    h52_1.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    c53_1 = ws1['C53']
    c53_1.value = 'POA / Condo Disclosures'
    c53_1.border = Border(top=hair, bottom=hair, left=thin)

    d53_1 = ws1['D53']
    d53_1.value = 'Commonly Used Fee'
    d53_1.border = Border(top=hair, bottom=hair)

    e53_1 = ws1['E53']
    e53_1.value = poa_condo_disclosure_fee
    e53_1.number_format = acct_fmt
    e53_1.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    f53_1 = ws1['F53']
    f53_1.value = poa_condo_disclosure_fee
    f53_1.number_format = acct_fmt
    f53_1.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    g53_1 = ws1['G53']
    g53_1.value = poa_condo_disclosure_fee
    g53_1.number_format = acct_fmt
    g53_1.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    h53_1 = ws1['H53']
    h53_1.value = poa_condo_disclosure_fee
    h53_1.number_format = acct_fmt
    h53_1.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    c54_1 = ws1['C54']
    c54_1.value = 'Pre Occupancy Credit to Seller'
    c54_1.border = Border(top=hair, bottom=hair, left=thin)

    d54_1 = ws1['D54']
    d54_1.value = 'Negotiated Into Contract'
    d54_1.border = Border(top=hair, bottom=hair)

    e54_1 = ws1['E54']
    e54_1.value = offer_1_pre_occupancy_credit_amt
    e54_1.number_format = acct_fmt
    e54_1.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    f54_1 = ws1['F54']
    f54_1.value = offer_2_pre_occupancy_credit_amt
    f54_1.number_format = acct_fmt
    f54_1.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    g54_1 = ws1['G54']
    g54_1.value = offer_3_pre_occupancy_credit_amt
    g54_1.number_format = acct_fmt
    g54_1.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    h54_1 = ws1['H54']
    h54_1.value = offer_4_pre_occupancy_credit_amt
    h54_1.number_format = acct_fmt
    h54_1.border = Border(top=hair, bottom=hair, left=thin, right=thin)

    c55_1 = ws1['C55']
    c55_1.value = 'Post Occupancy Cost to Seller'
    c55_1.border = Border(top=hair, bottom=thin, left=thin)

    d55_1 = ws1['D55']
    d55_1.value = 'Negotiated Into Contract'
    d55_1.border = Border(top=hair, bottom=thin)

    e55_1 = ws1['E55']
    e55_1.value = offer_1_post_occupancy_cost_amt
    e55_1.number_format = acct_fmt
    e55_1.border = Border(top=hair, bottom=thin, left=thin, right=thin)

    f55_1 = ws1['F55']
    f55_1.value = offer_2_post_occupancy_cost_amt
    f55_1.number_format = acct_fmt
    f55_1.border = Border(top=hair, bottom=thin, left=thin, right=thin)

    g55_1 = ws1['G55']
    g55_1.value = offer_3_post_occupancy_cost_amt
    g55_1.number_format = acct_fmt
    g55_1.border = Border(top=hair, bottom=thin, left=thin, right=thin)

    h55_1 = ws1['H55']
    h55_1.value = offer_4_post_occupancy_cost_amt
    h55_1.number_format = acct_fmt
    h55_1.border = Border(top=hair, bottom=thin, left=thin, right=thin)

    e56_1 = ws1['E56']
    e56_1.value = '=SUM(E49:E53,E55)-E54'
    e56_1.font = Font(bold=True)
    e56_1.number_format = acct_fmt
    e56_1.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    f56_1 = ws1['F56']
    f56_1.value = '=SUM(F49:F53,F55)-F54'
    f56_1.font = Font(bold=True)
    f56_1.number_format = acct_fmt
    f56_1.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    g56_1 = ws1['G56']
    g56_1.value = '=SUM(G49:G53,G55)-G54'
    g56_1.font = Font(bold=True)
    g56_1.number_format = acct_fmt
    g56_1.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    h56_1 = ws1['H56']
    h56_1.value = '=SUM(H49:H53,H55)-H54'
    h56_1.font = Font(bold=True)
    h56_1.number_format = acct_fmt
    h56_1.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    # Subtotals section
    ws1.merge_cells('C58:D58')
    top_left_cell_four_1 = ws1['C58']
    top_left_cell_four_1.value = 'TOTAL ESTIMATED COST OF SETTLEMENT'
    top_left_cell_four_1.font = Font(bold=True)
    top_left_cell_four_1.alignment = Alignment(horizontal='right')

    e58_1 = ws1['E58']
    e58_1.value = '=SUM(E37,E42,E47,E56)'
    e58_1.font = Font(bold=True)
    e58_1.number_format = acct_fmt
    e58_1.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    f58_1 = ws1['F58']
    f58_1.value = '=SUM(F37,F42,F47,F56)'
    f58_1.font = Font(bold=True)
    f58_1.number_format = acct_fmt
    f58_1.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    g58_1 = ws1['G58']
    g58_1.value = '=SUM(G37,G42,G47,G56)'
    g58_1.font = Font(bold=True)
    g58_1.number_format = acct_fmt
    g58_1.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    h58_1 = ws1['H58']
    h58_1.value = '=SUM(H37,H42,H47,H56)'
    h58_1.font = Font(bold=True)
    h58_1.number_format = acct_fmt
    h58_1.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    ws1.merge_cells('C59:D59')
    top_left_cell_five_1 = ws1['C59']
    top_left_cell_five_1.value = 'Offer Amount ($)'
    top_left_cell_five_1.font = Font(bold=True)
    top_left_cell_five_1.alignment = Alignment(horizontal='right')

    e59_1 = ws1['E59']
    e59_1.value = offer_1_amt
    e59_1.font = Font(bold=True)
    e59_1.number_format = acct_fmt
    e59_1.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    f59_1 = ws1['F59']
    f59_1.value = offer_2_amt
    f59_1.font = Font(bold=True)
    f59_1.number_format = acct_fmt
    f59_1.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    g59_1 = ws1['G59']
    g59_1.value = offer_3_amt
    g59_1.font = Font(bold=True)
    g59_1.number_format = acct_fmt
    g59_1.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    h59_1 = ws1['H59']
    h59_1.value = offer_4_amt
    h59_1.font = Font(bold=True)
    h59_1.number_format = acct_fmt
    h59_1.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    ws1.merge_cells('C60:D60')
    top_left_cell_six_1 = ws1['C60']
    top_left_cell_six_1.value = 'LESS: Total Estimated Cost of Settlement'
    top_left_cell_six_1.font = Font(bold=True)
    top_left_cell_six_1.alignment = Alignment(horizontal='right')

    e60_1 = ws1['E60']
    e60_1.value = '=-SUM(E37,E42,E47,E56)'
    e60_1.font = Font(bold=True)
    e60_1.number_format = acct_fmt
    e60_1.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    f60_1 = ws1['F60']
    f60_1.value = '=-SUM(F37,F42,F47,F56)'
    f60_1.font = Font(bold=True)
    f60_1.number_format = acct_fmt
    f60_1.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    g60_1 = ws1['G60']
    g60_1.value = '=-SUM(G37,G42,G47,G56)'
    g60_1.font = Font(bold=True)
    g60_1.number_format = acct_fmt
    g60_1.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    h60_1 = ws1['H60']
    h60_1.value = '=-SUM(H37,H42,H47,H56)'
    h60_1.font = Font(bold=True)
    h60_1.number_format = acct_fmt
    h60_1.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    ws1.merge_cells('C61:D61')
    top_left_cell_seven_1 = ws1['C61']
    top_left_cell_seven_1.value = 'ESTIMATED TOTAL NET PROCEEDS'
    top_left_cell_seven_1.font = Font(bold=True)
    top_left_cell_seven_1.alignment = Alignment(horizontal='right')

    e61_1 = ws1['E61']
    e61_1.value = '=SUM(E59:E60)'
    e61_1.font = Font(bold=True)
    e61_1.number_format = acct_fmt
    e61_1.fill = PatternFill(start_color=yellow_fill, end_color=yellow_fill, fill_type='solid')
    e61_1.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    f61_1 = ws1['F61']
    f61_1.value = '=SUM(F59:F60)'
    f61_1.font = Font(bold=True)
    f61_1.number_format = acct_fmt
    f61_1.fill = PatternFill(start_color=yellow_fill, end_color=yellow_fill, fill_type='solid')
    f61_1.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    g61_1 = ws1['G61']
    g61_1.value = '=SUM(G59:G60)'
    g61_1.font = Font(bold=True)
    g61_1.number_format = acct_fmt
    g61_1.fill = PatternFill(start_color=yellow_fill, end_color=yellow_fill, fill_type='solid')
    g61_1.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    h61_1 = ws1['H61']
    h61_1.value = '=SUM(H59:H60)'
    h61_1.font = Font(bold=True)
    h61_1.number_format = acct_fmt
    h61_1.fill = PatternFill(start_color=yellow_fill, end_color=yellow_fill, fill_type='solid')
    h61_1.border = Border(top=thin, bottom=thin, left=thin, right=thin)

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
    ws1.merge_cells('C70:H74')
    top_left_cell_eight = ws1['C70']
    top_left_cell_eight.value = '''These estimates are not guaranteed and may not include escrows. Escrow balances are reimbursed by the existing lender. Taxes, rents & association dues are pro-rated at settlement. Under Virginia Law, the seller's proceeds may not be available for up to 2 business days following recording of the deed. Seller acknowledges receipt of this statement.'''
    top_left_cell_eight.font = Font(italic=True)
    top_left_cell_eight.alignment = Alignment(horizontal='left', vertical='top', wrapText=True)

    # hide calculation description column
    ws1.column_dimensions['D'].hidden = True
    ws2.column_dimensions['D'].hidden = True
    ws3.column_dimensions['D'].hidden = True


    # hide unused columns based on number of offers
    if offer_qty == 1:
        ws1.column_dimensions['F'].hidden = True
        ws1.column_dimensions['G'].hidden = True
        ws1.column_dimensions['H'].hidden = True
    elif offer_qty == 2:
        ws1.column_dimensions['G'].hidden = True
        ws1.column_dimensions['H'].hidden = True
    elif offer_qty == 3:
        ws1.column_dimensions['H'].hidden = True
    elif offer_qty >= 3:
        ws1.column_dimensions['E'].hidden = False
        ws1.column_dimensions['F'].hidden = False
        ws1.column_dimensions['G'].hidden = False

    # wb.save(filename=dest_filename)

    with NamedTemporaryFile() as tmp:
        wb.save(tmp.name)
        data = BytesIO(tmp.read())

    return data
