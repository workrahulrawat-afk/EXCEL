import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

# Create a new workbook with write_only=False to ensure proper formula handling
wb = Workbook()
wb.iso_dates = True

# Remove default sheet
if 'Sheet' in wb.sheetnames:
    wb.remove(wb['Sheet'])

# ============================================================================
# HELPER FUNCTIONS FOR STYLING
# ============================================================================
def style_header(ws, row, start_col, end_col, bg_color="366092", font_color="FFFFFF"):
    """Style header row with background and font color"""
    for col in range(start_col, end_col + 1):
        cell = ws.cell(row=row, column=col)
        cell.font = Font(bold=True, color=font_color, size=11)
        cell.fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )

def style_metric_banner(ws, row, col, label, bg_color="4472C4"):
    """Style metric banner cells"""
    label_cell = ws.cell(row=row, column=col)
    label_cell.value = label
    label_cell.font = Font(bold=True, color="FFFFFF", size=10)
    label_cell.fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")
    label_cell.alignment = Alignment(horizontal="center", vertical="center")
    
    value_cell = ws.cell(row=row+1, column=col)
    value_cell.font = Font(bold=True, size=12)
    value_cell.alignment = Alignment(horizontal="center", vertical="center")
    value_cell.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

# ============================================================================
# SHEET 1: TRANSACTIONS (Master Log)
# ============================================================================
ws_trans = wb.create_sheet("Transactions")

# Set column widths
ws_trans.column_dimensions['A'].width = 12
ws_trans.column_dimensions['B'].width = 20
ws_trans.column_dimensions['C'].width = 15
ws_trans.column_dimensions['D'].width = 12
ws_trans.column_dimensions['E'].width = 15
ws_trans.column_dimensions['F'].width = 15
ws_trans.column_dimensions['G'].width = 15

# Headers
headers = ['Date', 'Stock Name', 'Transaction Type', 'Quantity', 'Price per Share', 'Brokerage/STT', 'Total Amount']
for col_num, header in enumerate(headers, 1):
    ws_trans.cell(row=1, column=col_num, value=header)

style_header(ws_trans, 1, 1, 7)

# Add data validation for Transaction Type (Dropdown: Buy/Sell)
dv = DataValidation(type="list", formula1='"Buy,Sell"', allow_blank=False)
dv.error = 'Please select Buy or Sell'
dv.errorTitle = 'Invalid Entry'
ws_trans.add_data_validation(dv)
dv.add('C2:C1000')

# Add formula for Total Amount in column G (starting from row 2)
for row in range(2, 1001):
    ws_trans[f'G{row}'] = f'=IF(AND(D{row}<>"",E{row}<>""),D{row}*E{row}+IF(ISBLANK(F{row}),0,F{row}),"")'

# ============================================================================
# SHEET 2: FIFO_Buy_Table (The Engine)
# ============================================================================
ws_fifo = wb.create_sheet("FIFO_Buy_Table")

# Set column widths
ws_fifo.column_dimensions['A'].width = 20
ws_fifo.column_dimensions['B'].width = 12
ws_fifo.column_dimensions['C'].width = 12
ws_fifo.column_dimensions['D'].width = 15
ws_fifo.column_dimensions['E'].width = 15
ws_fifo.column_dimensions['F'].width = 18
ws_fifo.column_dimensions['G'].width = 18
ws_fifo.column_dimensions['H'].width = 18
ws_fifo.column_dimensions['I'].width = 15
ws_fifo.column_dimensions['J'].width = 15

# Headers
fifo_headers = ['Stock Name', 'Buy Date', 'Buy Qty', 'Buy Price', 'Total Value', 
                'Cumulative Buy Qty', 'Total Sold Qty', 'Sold from this Lot', 'Remaining Qty', 'Remaining Value']
for col_num, header in enumerate(fifo_headers, 1):
    ws_fifo.cell(row=1, column=col_num, value=header)

style_header(ws_fifo, 1, 1, 10, bg_color="70AD47")

# FIFO Engine - Row by row formulas that work with dynamic data
# These formulas will auto-fill based on the Transactions sheet

# Starting from row 2, add formulas for up to 200 rows (reduced for stability)
for row in range(2, 202):
    # Stock Name - get from filtered buy transactions
    ws_fifo[f'A{row}'] = f'=IFERROR(INDEX(SORT(FILTER(Transactions!B$2:B$1000,(Transactions!C$2:C$1000="Buy")*(Transactions!B$2:B$1000<>"")),1,1),{row-1}),"")'
    
    # Buy Date
    ws_fifo[f'B{row}'] = f'=IF(A{row}="","",INDEX(SORT(FILTER(Transactions!A$2:G$1000,(Transactions!C$2:C$1000="Buy")*(Transactions!B$2:B$1000<>"")),2,1),{row-1},1))'
    
    # Buy Qty
    ws_fifo[f'C{row}'] = f'=IF(A{row}="","",INDEX(SORT(FILTER(Transactions!A$2:G$1000,(Transactions!C$2:C$1000="Buy")*(Transactions!B$2:B$1000<>"")),2,1),{row-1},4))'
    
    # Buy Price
    ws_fifo[f'D{row}'] = f'=IF(A{row}="","",INDEX(SORT(FILTER(Transactions!A$2:G$1000,(Transactions!C$2:C$1000="Buy")*(Transactions!B$2:B$1000<>"")),2,1),{row-1},5))'
    
    # Total Value
    ws_fifo[f'E{row}'] = f'=IF(C{row}="","",C{row}*D{row})'
    
    # Cumulative Buy Qty (for this stock up to this row)
    ws_fifo[f'F{row}'] = f'=IF(A{row}="","",SUMIFS(C:C,A:A,A{row},B:B,"<="&B{row}))'
    
    # Total Sold Qty (for this stock)
    ws_fifo[f'G{row}'] = f'=IF(A{row}="","",SUMIFS(Transactions!D:D,Transactions!B:B,A{row},Transactions!C:C,"Sell"))'
    
    # Sold from this Lot (FIFO logic) - Fixed to avoid circular reference
    if row == 2:
        ws_fifo[f'H{row}'] = f'=IF(A{row}="","",MAX(0,MIN(C{row},G{row})))'
    else:
        ws_fifo[f'H{row}'] = f'=IF(A{row}="","",IF(G{row}<=0,0,MAX(0,MIN(C{row},G{row}-SUMIFS(H$2:H${row-1},A$2:A${row-1},A{row})))))'
    
    # Remaining Qty
    ws_fifo[f'I{row}'] = f'=IF(C{row}="","",MAX(0,C{row}-H{row}))'
    
    # Remaining Value
    ws_fifo[f'J{row}'] = f'=IF(I{row}="","",I{row}*D{row})'

# Add helper text for understanding
ws_fifo['L1'] = 'FIFO Engine - Auto-populated from Buy transactions'
ws_fifo['L1'].font = Font(italic=True, color="666666")

# Instructions cell
ws_fifo['L2'] = 'This sheet automatically tracks which buy lots have been sold using FIFO method.'
ws_fifo['L2'].font = Font(italic=True, size=9, color="666666")
ws_fifo['L2'].alignment = Alignment(wrap_text=True)

# Additional explanation
ws_fifo['L4'] = 'How FIFO works here:'
ws_fifo['L4'].font = Font(bold=True, size=10)
ws_fifo['L5'] = '1. All Buy transactions are sorted by date'
ws_fifo['L5'].font = Font(size=9)
ws_fifo['L6'] = '2. Sell transactions consume from oldest lots first'
ws_fifo['L6'].font = Font(size=9)
ws_fifo['L7'] = '3. Remaining Qty shows unsold shares from each lot'
ws_fifo['L7'].font = Font(size=9)
ws_fifo['L8'] = '4. Remaining Value = Remaining Qty × Original Buy Price'
ws_fifo['L8'].font = Font(size=9)

# ============================================================================
# SHEET 3: Holdings (Current Portfolio)
# ============================================================================
ws_holdings = wb.create_sheet("Holdings")

# Set column widths
ws_holdings.column_dimensions['A'].width = 20
ws_holdings.column_dimensions['B'].width = 15
ws_holdings.column_dimensions['C'].width = 18
ws_holdings.column_dimensions['D'].width = 18
ws_holdings.column_dimensions['E'].width = 15
ws_holdings.column_dimensions['F'].width = 15
ws_holdings.column_dimensions['G'].width = 18

# Headers
holdings_headers = ['Stock Name', 'Net Holding Qty', 'Current Avg Price', 'Total Invested Value', 
                    'CMP', 'Current Value', 'Unrealized P&L']
for col_num, header in enumerate(holdings_headers, 1):
    ws_holdings.cell(row=1, column=col_num, value=header)

style_header(ws_holdings, 1, 1, 7, bg_color="FFC000")

# Holdings formulas - Get unique stock names with remaining quantity
ws_holdings['A2'] = '=IFERROR(UNIQUE(FILTER(Transactions!B$2:B$1000,(Transactions!C$2:C$1000="Buy")*(Transactions!B$2:B$1000<>""))),"")' 

# Net Holding Qty formula
ws_holdings['B2'] = '=IF(A2="","",SUMIFS(Transactions!D:D,Transactions!B:B,A2,Transactions!C:C,"Buy")-SUMIFS(Transactions!D:D,Transactions!B:B,A2,Transactions!C:C,"Sell"))'

# Current Average Price formula (from FIFO remaining lots)
ws_holdings['C2'] = '=IF(B2="","",IF(B2<=0,"",SUMIFS(FIFO_Buy_Table!J:J,FIFO_Buy_Table!A:A,A2)/SUMIFS(FIFO_Buy_Table!I:I,FIFO_Buy_Table!A:A,A2)))'

# Total Invested Value (based on current holdings at original prices)
ws_holdings['D2'] = '=IF(B2="","",IF(B2<=0,"",SUMIFS(FIFO_Buy_Table!J:J,FIFO_Buy_Table!A:A,A2)))'

# CMP - left blank for user to input or use Excel Data Types
ws_holdings['E2'] = ''

# Current Value
ws_holdings['F2'] = '=IF(OR(E2="",B2=""),"",B2*E2)'

# Unrealized P&L
ws_holdings['G2'] = '=IF(OR(F2="",D2=""),"",F2-D2)'

# Add note about CMP column
ws_holdings['I1'] = 'Note: CMP Column'
ws_holdings['I1'].font = Font(bold=True, color="FF0000")
ws_holdings['I2'] = 'Enter current market price manually or use Excel Stock Data Type'
ws_holdings['I2'].font = Font(italic=True, size=9, color="666666")
ws_holdings['I2'].alignment = Alignment(wrap_text=True)
ws_holdings.merge_cells('I2:K3')

# ============================================================================
# SHEET 4: Dashboard
# ============================================================================
ws_dash = wb.create_sheet("Dashboard", 0)  # Insert at position 0 (first sheet)

# Set column widths
for col in ['A', 'B', 'C', 'D', 'E', 'F']:
    ws_dash.column_dimensions[col].width = 22

# Title
ws_dash['A1'] = 'STOCK PORTFOLIO TRACKER - FIFO METHOD'
ws_dash['A1'].font = Font(bold=True, size=16, color="FFFFFF")
ws_dash['A1'].fill = PatternFill(start_color="203764", end_color="203764", fill_type="solid")
ws_dash['A1'].alignment = Alignment(horizontal="center", vertical="center")
ws_dash.merge_cells('A1:F1')
ws_dash.row_dimensions[1].height = 30

# Metric banners
metrics = [
    ('A', 'Total Invested Value'),
    ('B', 'Current Portfolio Value'),
    ('C', 'Total Realized Profit'),
    ('D', 'Total Unrealized Profit'),
    ('E', 'Overall P&L'),
    ('F', 'Return %')
]

row = 3
for col_letter, label in metrics:
    col_num = ord(col_letter) - ord('A') + 1
    style_metric_banner(ws_dash, row, col_num, label)

# Formulas for metrics
# Total Invested Value
ws_dash['A4'] = '=SUMIF(Holdings!B:B,">0",Holdings!D:D)'
ws_dash['A4'].number_format = '₹#,##0.00'

# Current Portfolio Value
ws_dash['B4'] = '=SUMIF(Holdings!B:B,">0",Holdings!F:F)'
ws_dash['B4'].number_format = '₹#,##0.00'

# Total Realized Profit (Sale Revenue - FIFO Cost of Sold Shares)
ws_dash['C4'] = '=SUMIFS(Transactions!G:G,Transactions!C:C,"Sell")-SUMIF(FIFO_Buy_Table!H:H,">0",FIFO_Buy_Table!E:E)*SUMIF(FIFO_Buy_Table!H:H,">0",FIFO_Buy_Table!H:H)/SUMIF(FIFO_Buy_Table!C:C,">0",FIFO_Buy_Table!C:C)'
ws_dash['C4'].number_format = '₹#,##0.00'

# Total Unrealized Profit
ws_dash['D4'] = '=SUM(Holdings!G:G)'
ws_dash['D4'].number_format = '₹#,##0.00'

# Overall P&L
ws_dash['E4'] = '=IFERROR(C4+D4,0)'
ws_dash['E4'].number_format = '₹#,##0.00'

# Return %
ws_dash['F4'] = '=IF(A4=0,0,E4/A4)'
ws_dash['F4'].number_format = '0.00%'

# Add conditional formatting colors for P&L
from openpyxl.styles import Color
ws_dash['C4'].font = Font(bold=True, size=12)
ws_dash['D4'].font = Font(bold=True, size=12)
ws_dash['E4'].font = Font(bold=True, size=12)

# Summary table below
ws_dash['A7'] = 'CURRENT HOLDINGS SUMMARY'
ws_dash['A7'].font = Font(bold=True, size=12, color="FFFFFF")
ws_dash['A7'].fill = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid")
ws_dash.merge_cells('A7:F7')

# Reference to Holdings sheet
ws_dash['A9'] = '=Holdings!A1'
ws_dash['B9'] = '=Holdings!B1'
ws_dash['C9'] = '=Holdings!C1'
ws_dash['D9'] = '=Holdings!D1'
ws_dash['E9'] = '=Holdings!E1'
ws_dash['F9'] = '=Holdings!F1'
style_header(ws_dash, 9, 1, 6, bg_color="70AD47")

# Link to holdings data
for row in range(10, 30):
    ws_dash[f'A{row}'] = f'=IFERROR(Holdings!A{row-8},"")'
    ws_dash[f'B{row}'] = f'=IFERROR(Holdings!B{row-8},"")'
    ws_dash[f'C{row}'] = f'=IFERROR(Holdings!C{row-8},"")'
    ws_dash[f'D{row}'] = f'=IFERROR(Holdings!D{row-8},"")'
    ws_dash[f'E{row}'] = f'=IFERROR(Holdings!E{row-8},"")'
    ws_dash[f'F{row}'] = f'=IFERROR(Holdings!F{row-8},"")'
    
    # Format numbers
    ws_dash[f'B{row}'].number_format = '#,##0'
    ws_dash[f'C{row}'].number_format = '₹#,##0.00'
    ws_dash[f'D{row}'].number_format = '₹#,##0.00'
    ws_dash[f'E{row}'].number_format = '₹#,##0.00'
    ws_dash[f'F{row}'].number_format = '₹#,##0.00'

# ============================================================================
# SHEET 5: Stock_Template (Individual Stock Analysis)
# ============================================================================
ws_template = wb.create_sheet("Stock_Template")

# Set column widths
ws_template.column_dimensions['A'].width = 15
ws_template.column_dimensions['B'].width = 15
ws_template.column_dimensions['C'].width = 15
ws_template.column_dimensions['D'].width = 18
ws_template.column_dimensions['E'].width = 15

# Stock Name Placeholder
ws_template['A1'] = 'Stock Name:'
ws_template['A1'].font = Font(bold=True, size=12)
ws_template['B1'] = 'ENTER_STOCK_NAME_HERE'
ws_template['B1'].font = Font(bold=True, size=12, color="FF0000")
ws_template['B1'].fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

# Top Banner Metrics
banner_row = 3
metrics_template = [
    ('A', 'Realized P&L'),
    ('B', 'Unrealized P&L'),
    ('C', 'Current Holding Qty'),
    ('D', 'Current Invested Value'),
    ('E', 'Average Buy Price')
]

for col_letter, label in metrics_template:
    col_num = ord(col_letter) - ord('A') + 1
    style_metric_banner(ws_template, banner_row, col_num, label, bg_color="5B9BD5")

# Banner formulas referencing B1 (stock name placeholder)
# Realized P&L - simplified formula
ws_template['A4'] = '=IF(B1="ENTER_STOCK_NAME_HERE",0,SUMIFS(Transactions!G:G,Transactions!B:B,B1,Transactions!C:C,"Sell")-SUMIFS(FIFO_Buy_Table!E:E,FIFO_Buy_Table!A:A,B1,FIFO_Buy_Table!H:H,">0"))'
ws_template['A4'].number_format = '₹#,##0.00'

# Unrealized P&L
ws_template['B4'] = '=IF(B1="ENTER_STOCK_NAME_HERE",0,IFERROR(INDEX(Holdings!G:G,MATCH(B1,Holdings!A:A,0)),0))'
ws_template['B4'].number_format = '₹#,##0.00'

# Current Holding Qty
ws_template['C4'] = '=IF(B1="ENTER_STOCK_NAME_HERE",0,SUMIFS(Transactions!D:D,Transactions!B:B,B1,Transactions!C:C,"Buy")-SUMIFS(Transactions!D:D,Transactions!B:B,B1,Transactions!C:C,"Sell"))'
ws_template['C4'].number_format = '#,##0'

# Current Invested Value
ws_template['D4'] = '=IF(B1="ENTER_STOCK_NAME_HERE",0,IFERROR(INDEX(Holdings!D:D,MATCH(B1,Holdings!A:A,0)),0))'
ws_template['D4'].number_format = '₹#,##0.00'

# Average Buy Price
ws_template['E4'] = '=IF(OR(B1="ENTER_STOCK_NAME_HERE",C4=0),0,D4/C4)'
ws_template['E4'].number_format = '₹#,##0.00'

# Transaction History Table
ws_template['A7'] = 'COMPLETE TRANSACTION HISTORY'
ws_template['A7'].font = Font(bold=True, size=11, color="FFFFFF")
ws_template['A7'].fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
ws_template.merge_cells('A7:G7')

# History Headers
history_headers = ['Date', 'Stock Name', 'Type', 'Quantity', 'Price', 'Brokerage/STT', 'Total Amount']
for col_num, header in enumerate(history_headers, 1):
    ws_template.cell(row=8, column=col_num, value=header)

style_header(ws_template, 8, 1, 7, bg_color="4472C4")

# FILTER formula to show all transactions for the stock in B1
ws_template['A9'] = '=IFERROR(FILTER(Transactions!A$2:G$1000,Transactions!B$2:B$1000=B1),"No transactions found")'

# Instructions
ws_template['A6'] = 'Instructions: Enter stock name in B1, then duplicate this sheet for each stock you want to track individually.'
ws_template['A6'].font = Font(italic=True, size=9, color="666666")
ws_template.merge_cells('A6:G6')

# ============================================================================
# SAVE THE WORKBOOK
# ============================================================================
filename = '/projects/sandbox/EXCEL/Stock_Portfolio_Tracker_FIFO.xlsx'
wb.save(filename)
print(f"Excel file created successfully: {filename}")
