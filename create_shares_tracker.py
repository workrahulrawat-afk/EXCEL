#!/usr/bin/env python3
"""
Shares Buy/Sell Tracking Dashboard for Excel M365
Creates a comprehensive workbook with Transactions, Portfolio, Dashboard, and Settings sheets
"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, numbers
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule
from datetime import datetime, timedelta
import random

def create_shares_tracker():
    """Create the complete shares tracking workbook"""
    
    # Create workbook
    wb = Workbook()
    wb.remove(wb.active)  # Remove default sheet
    
    # Create sheets
    ws_transactions = wb.create_sheet("Transactions", 0)
    ws_portfolio = wb.create_sheet("Portfolio", 1)
    ws_dashboard = wb.create_sheet("Dashboard", 2)
    ws_settings = wb.create_sheet("Settings", 3)
    
    print("Creating Transactions sheet...")
    create_transactions_sheet(ws_transactions)
    
    print("Creating Portfolio sheet...")
    create_portfolio_sheet(ws_portfolio)
    
    print("Creating Dashboard sheet...")
    create_dashboard_sheet(ws_dashboard)
    
    print("Creating Settings sheet...")
    create_settings_sheet(ws_settings)
    
    print("Adding sample data...")
    add_sample_data(ws_transactions, ws_portfolio, ws_settings)
    
    # Save workbook
    filename = "Shares_Tracking_Dashboard.xlsx"
    wb.save(filename)
    print(f"\n✅ Excel file created successfully: {filename}")
    return filename

def create_transactions_sheet(ws):
    """Create the Transactions sheet with all formulas"""
    
    # Headers
    headers = [
        "Date", "Ticker Symbol", "Stock Name", "Type", 
        "Quantity", "Price per Share", "Total Amount", 
        "Brokerage Fee", "Net Amount"
    ]
    
    # Header styling
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_num)
        cell.value = header
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
    
    # Set column widths
    ws.column_dimensions['A'].width = 12
    ws.column_dimensions['B'].width = 12
    ws.column_dimensions['C'].width = 25
    ws.column_dimensions['D'].width = 8
    ws.column_dimensions['E'].width = 10
    ws.column_dimensions['F'].width = 15
    ws.column_dimensions['G'].width = 15
    ws.column_dimensions['H'].width = 14
    ws.column_dimensions['I'].width = 15
    
    # Add formulas starting from row 2
    for row in range(2, 102):  # Add formulas for 100 rows
        # G: Total Amount = Quantity * Price per Share
        ws[f'G{row}'] = f'=IF(E{row}="","",E{row}*F{row})'
        
        # I: Net Amount = IF(Type=BUY, Total+Fee, Total-Fee)
        ws[f'I{row}'] = f'=IF(D{row}="","",IF(D{row}="BUY",G{row}+H{row},G{row}-H{row}))'
    
    # Data Validation for Type column (BUY/SELL)
    dv = DataValidation(type="list", formula1='"BUY,SELL"', allow_blank=True)
    dv.error = 'Please select BUY or SELL'
    dv.errorTitle = 'Invalid Entry'
    ws.add_data_validation(dv)
    dv.add(f'D2:D1000')
    
    # Format columns
    for row in range(2, 102):
        ws[f'A{row}'].number_format = 'MM/DD/YYYY'
        ws[f'F{row}'].number_format = '$#,##0.00'
        ws[f'G{row}'].number_format = '$#,##0.00'
        ws[f'H{row}'].number_format = '$#,##0.00'
        ws[f'I{row}'].number_format = '$#,##0.00'
    
    # Freeze panes
    ws.freeze_panes = 'A2'

def create_portfolio_sheet(ws):
    """Create the Portfolio sheet with holdings summary"""
    
    # Headers
    headers = [
        "Ticker", "Stock Name", "Total Bought", "Total Sold", 
        "Current Holdings", "Avg Buy Price", "Total Invested", 
        "Current Market Price", "Current Value", "Profit/Loss", "P/L %"
    ]
    
    # Header styling
    header_fill = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_num)
        cell.value = header
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
    
    # Set column widths
    widths = [10, 25, 12, 12, 15, 14, 15, 18, 15, 15, 10]
    for i, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = width
    
    # Add formulas (starting from row 2, leaving space for manual ticker entry)
    for row in range(2, 52):  # Support up to 50 stocks
        # B: Stock Name - lookup from Transactions
        ws[f'B{row}'] = f'=IF(A{row}="","",IFERROR(INDEX(Transactions!C:C,MATCH(A{row},Transactions!B:B,0)),""))'
        
        # C: Total Bought
        ws[f'C{row}'] = f'=IF(A{row}="","",SUMIFS(Transactions!E:E,Transactions!B:B,A{row},Transactions!D:D,"BUY"))'
        
        # D: Total Sold
        ws[f'D{row}'] = f'=IF(A{row}="","",SUMIFS(Transactions!E:E,Transactions!B:B,A{row},Transactions!D:D,"SELL"))'
        
        # E: Current Holdings
        ws[f'E{row}'] = f'=IF(A{row}="","",C{row}-D{row})'
        
        # F: Avg Buy Price
        ws[f'F{row}'] = f'=IF(C{row}=0,"",SUMIFS(Transactions!I:I,Transactions!B:B,A{row},Transactions!D:D,"BUY")/C{row})'
        
        # G: Total Invested
        ws[f'G{row}'] = f'=IF(E{row}="","",E{row}*F{row})'
        
        # I: Current Value
        ws[f'I{row}'] = f'=IF(E{row}="","",IF(H{row}="","",E{row}*H{row}))'
        
        # J: Profit/Loss
        ws[f'J{row}'] = f'=IF(I{row}="","",I{row}-G{row})'
        
        # K: P/L %
        ws[f'K{row}'] = f'=IF(G{row}=0,"",IF(J{row}="","",(J{row}/G{row})*100))'
    
    # Format columns
    for row in range(2, 52):
        for col in ['F', 'G', 'H', 'I', 'J']:
            ws[f'{col}{row}'].number_format = '$#,##0.00'
        ws[f'K{row}'].number_format = '0.00"%"'
    
    # Conditional formatting for P/L
    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    red_font = Font(color="9C0006")
    green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    green_font = Font(color="006100")
    
    ws.conditional_formatting.add('J2:J51',
        CellIsRule(operator='lessThan', formula=['0'], fill=red_fill, font=red_font))
    ws.conditional_formatting.add('J2:J51',
        CellIsRule(operator='greaterThan', formula=['0'], fill=green_fill, font=green_font))
    
    ws.conditional_formatting.add('K2:K51',
        CellIsRule(operator='lessThan', formula=['0'], fill=red_fill, font=red_font))
    ws.conditional_formatting.add('K2:K51',
        CellIsRule(operator='greaterThan', formula=['0'], fill=green_fill, font=green_font))
    
    # Freeze panes
    ws.freeze_panes = 'A2'

def create_dashboard_sheet(ws):
    """Create the Dashboard sheet with key metrics"""
    
    # Title
    ws['A1'] = 'SHARES TRADING DASHBOARD'
    ws['A1'].font = Font(bold=True, size=18, color="1F4E78")
    ws.merge_cells('A1:F1')
    ws['A1'].alignment = Alignment(horizontal="center", vertical="center")
    
    # Key Metrics Section
    ws['A3'] = 'KEY METRICS'
    ws['A3'].font = Font(bold=True, size=14, color="1F4E78")
    ws.merge_cells('A3:C3')
    
    # Metric labels and formulas
    metrics = [
        ("Total Investment", "=SUM(Portfolio!G:G)", "B5"),
        ("Current Portfolio Value", "=SUM(Portfolio!I:I)", "B6"),
        ("Total Profit/Loss", "=B6-B5", "B7"),
        ("Total P/L %", "=IF(B5=0,0,(B7/B5)*100)", "B8"),
        ("Number of Active Stocks", "=COUNTIF(Portfolio!E:E,\">0\")", "B9"),
        ("Total Transactions", "=COUNTA(Transactions!A:A)-1", "B10"),
    ]
    
    label_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    label_font = Font(bold=True, size=11)
    value_fill = PatternFill(start_color="E7E6E6", end_color="E7E6E6", fill_type="solid")
    
    for idx, (label, formula, cell) in enumerate(metrics, 5):
        # Label
        ws[f'A{idx}'] = label
        ws[f'A{idx}'].fill = label_fill
        ws[f'A{idx}'].font = label_font
        ws[f'A{idx}'].alignment = Alignment(horizontal="left", vertical="center")
        
        # Value
        ws[cell] = formula
        ws[cell].fill = value_fill
        ws[cell].font = Font(size=11, bold=True)
        ws[cell].alignment = Alignment(horizontal="right", vertical="center")
        
        # Format based on metric type
        if idx in [5, 6, 7]:  # Currency
            ws[cell].number_format = '$#,##0.00'
        elif idx == 8:  # Percentage
            ws[cell].number_format = '0.00"%"'
        else:  # Number
            ws[cell].number_format = '0'
    
    # Conditional formatting for P/L
    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    
    ws.conditional_formatting.add('B7:B8',
        CellIsRule(operator='lessThan', formula=['0'], fill=red_fill))
    ws.conditional_formatting.add('B7:B8',
        CellIsRule(operator='greaterThan', formula=['0'], fill=green_fill))
    
    # Top Performers Section
    ws['D3'] = 'TOP PERFORMERS'
    ws['D3'].font = Font(bold=True, size=14, color="1F4E78")
    ws.merge_cells('D3:G3')
    
    top_headers = ["Ticker", "Stock Name", "Holdings", "P/L $"]
    for idx, header in enumerate(top_headers, 4):
        cell = ws.cell(row=4, column=idx)
        cell.value = header
        cell.fill = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid")
        cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center")
    
    # Note about top performers (manual entry needed)
    ws['D5'] = "Note: Manually copy top"
    ws['D6'] = "stocks from Portfolio sheet"
    ws.merge_cells('D5:G5')
    ws.merge_cells('D6:G6')
    ws['D5'].alignment = Alignment(horizontal="center", vertical="center")
    ws['D6'].alignment = Alignment(horizontal="center", vertical="center")
    ws['D5'].font = Font(italic=True, size=9)
    ws['D6'].font = Font(italic=True, size=9)
    
    # Recent Transactions Section
    ws['A12'] = 'RECENT TRANSACTIONS (Last 10)'
    ws['A12'].font = Font(bold=True, size=14, color="1F4E78")
    ws.merge_cells('A12:G12')
    
    recent_headers = ["Date", "Ticker", "Stock Name", "Type", "Quantity", "Price", "Net Amount"]
    for idx, header in enumerate(recent_headers, 1):
        cell = ws.cell(row=13, column=idx)
        cell.value = header
        cell.fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center")
    
    # Formulas to pull last 10 transactions (simple approach - latest entries)
    for row in range(14, 24):
        trans_row = row - 12  # Maps to Transactions rows 2-11
        ws[f'A{row}'] = f'=IF(Transactions!A{trans_row}="","",Transactions!A{trans_row})'
        ws[f'B{row}'] = f'=IF(Transactions!B{trans_row}="","",Transactions!B{trans_row})'
        ws[f'C{row}'] = f'=IF(Transactions!C{trans_row}="","",Transactions!C{trans_row})'
        ws[f'D{row}'] = f'=IF(Transactions!D{trans_row}="","",Transactions!D{trans_row})'
        ws[f'E{row}'] = f'=IF(Transactions!E{trans_row}="","",Transactions!E{trans_row})'
        ws[f'F{row}'] = f'=IF(Transactions!F{trans_row}="","",Transactions!F{trans_row})'
        ws[f'G{row}'] = f'=IF(Transactions!I{trans_row}="","",Transactions!I{trans_row})'
        
        # Format
        ws[f'A{row}'].number_format = 'MM/DD/YYYY'
        ws[f'F{row}'].number_format = '$#,##0.00'
        ws[f'G{row}'].number_format = '$#,##0.00'
    
    # Set column widths
    widths = [12, 10, 25, 8, 10, 12, 15]
    for i, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = width

def create_settings_sheet(ws):
    """Create the Settings sheet with reference data"""
    
    # Title
    ws['A1'] = 'SETTINGS & REFERENCE DATA'
    ws['A1'].font = Font(bold=True, size=14, color="1F4E78")
    ws.merge_cells('A1:D1')
    
    # Stock Reference Table
    ws['A3'] = 'Stock Reference Table'
    ws['A3'].font = Font(bold=True, size=12)
    ws.merge_cells('A3:D3')
    
    headers = ["Ticker Symbol", "Full Stock Name", "Sector", "Notes"]
    header_fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
    
    for idx, header in enumerate(headers, 1):
        cell = ws.cell(row=4, column=idx)
        cell.value = header
        cell.fill = header_fill
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
    
    # Set column widths
    ws.column_dimensions['A'].width = 15
    ws.column_dimensions['B'].width = 30
    ws.column_dimensions['C'].width = 20
    ws.column_dimensions['D'].width = 30
    
    # Instructions
    ws['F3'] = 'INSTRUCTIONS'
    ws['F3'].font = Font(bold=True, size=12, color="C00000")
    
    instructions = [
        "1. Enter all buy/sell transactions in the 'Transactions' sheet",
        "2. In the 'Portfolio' sheet, manually enter ticker symbols in column A",
        "3. All calculations will update automatically",
        "4. Update 'Current Market Price' in Portfolio sheet regularly",
        "5. View summary metrics in the 'Dashboard' sheet",
        "",
        "Tips:",
        "- Use consistent ticker symbols across all sheets",
        "- Brokerage fees can be 0 if not applicable",
        "- Dashboard shows real-time portfolio status",
        "- Export data regularly for backup",
    ]
    
    for idx, instruction in enumerate(instructions, 4):
        ws[f'F{idx}'] = instruction
        ws[f'F{idx}'].alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        if ":" in instruction:
            ws[f'F{idx}'].font = Font(bold=True)
    
    ws.column_dimensions['F'].width = 50

def add_sample_data(ws_transactions, ws_portfolio, ws_settings):
    """Add sample data to demonstrate the workbook"""
    
    # Sample stocks
    stocks = [
        ("AAPL", "Apple Inc.", "Technology"),
        ("MSFT", "Microsoft Corporation", "Technology"),
        ("GOOGL", "Alphabet Inc.", "Technology"),
        ("TSLA", "Tesla Inc.", "Automotive"),
        ("AMZN", "Amazon.com Inc.", "E-Commerce"),
        ("NVDA", "NVIDIA Corporation", "Technology"),
        ("META", "Meta Platforms Inc.", "Technology"),
    ]
    
    # Add to Settings sheet
    for idx, (ticker, name, sector) in enumerate(stocks, 5):
        ws_settings[f'A{idx}'] = ticker
        ws_settings[f'B{idx}'] = name
        ws_settings[f'C{idx}'] = sector
    
    # Sample transactions (10 transactions)
    sample_transactions = [
        ("2024-01-15", "AAPL", "Apple Inc.", "BUY", 10, 185.50, 10),
        ("2024-02-10", "MSFT", "Microsoft Corporation", "BUY", 5, 405.20, 8),
        ("2024-02-20", "GOOGL", "Alphabet Inc.", "BUY", 8, 142.30, 10),
        ("2024-03-05", "TSLA", "Tesla Inc.", "BUY", 15, 178.90, 12),
        ("2024-03-15", "AAPL", "Apple Inc.", "BUY", 5, 172.40, 8),
        ("2024-04-01", "NVDA", "NVIDIA Corporation", "BUY", 12, 880.50, 15),
        ("2024-04-20", "TSLA", "Tesla Inc.", "SELL", 5, 190.50, 10),
        ("2024-05-10", "META", "Meta Platforms Inc.", "BUY", 6, 485.30, 10),
        ("2024-05-15", "AMZN", "Amazon.com Inc.", "BUY", 7, 182.70, 10),
        ("2024-05-25", "AAPL", "Apple Inc.", "SELL", 3, 192.30, 8),
    ]
    
    # Add to Transactions sheet
    for idx, trans in enumerate(sample_transactions, 2):
        date_str, ticker, name, trans_type, qty, price, fee = trans
        ws_transactions[f'A{idx}'] = datetime.strptime(date_str, "%Y-%m-%d")
        ws_transactions[f'B{idx}'] = ticker
        ws_transactions[f'C{idx}'] = name
        ws_transactions[f'D{idx}'] = trans_type
        ws_transactions[f'E{idx}'] = qty
        ws_transactions[f'F{idx}'] = price
        ws_transactions[f'H{idx}'] = fee
    
    # Add tickers to Portfolio sheet (unique from transactions)
    unique_tickers = ["AAPL", "MSFT", "GOOGL", "TSLA", "NVDA", "META", "AMZN"]
    
    # Sample current market prices (slightly different from last transaction price)
    current_prices = {
        "AAPL": 195.50,
        "MSFT": 420.30,
        "GOOGL": 148.90,
        "TSLA": 185.20,
        "NVDA": 920.00,
        "META": 495.60,
        "AMZN": 188.90,
    }
    
    for idx, ticker in enumerate(unique_tickers, 2):
        ws_portfolio[f'A{idx}'] = ticker
        ws_portfolio[f'H{idx}'] = current_prices.get(ticker, 0)

if __name__ == "__main__":
    print("=" * 60)
    print("SHARES BUY/SELL TRACKING DASHBOARD CREATOR")
    print("=" * 60)
    print()
    
    try:
        filename = create_shares_tracker()
        print()
        print("=" * 60)
        print("✅ SUCCESS!")
        print("=" * 60)
        print(f"\nYour Excel file is ready: {filename}")
        print("\nThe workbook includes:")
        print("  📊 Transactions - Log all buy/sell trades")
        print("  💼 Portfolio - Current holdings summary")
        print("  📈 Dashboard - Key metrics and overview")
        print("  ⚙️  Settings - Reference data and instructions")
        print("\nSample data has been added for demonstration!")
        print()
    except Exception as e:
        print(f"❌ Error creating workbook: {e}")
        import traceback
        traceback.print_exc()
