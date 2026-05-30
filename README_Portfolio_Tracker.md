# Stock Portfolio Tracker - FIFO Method

## 📊 Overview

This Excel workbook is a **fully automated Stock Portfolio Tracker** that uses the **FIFO (First-In, First-Out)** method for calculating profits and managing holdings. 

**Key Feature:** 100% Formula-Based - No VBA, No Macros, No Scripts - Only native Excel 365 formulas!

---

## 📋 Workbook Structure

The workbook contains **5 sheets**:

### 1. **Dashboard** (Summary View)
The main overview sheet showing:
- **Total Invested Value** - Sum of money invested in current holdings
- **Current Portfolio Value** - Current market value of all holdings (requires CMP input)
- **Total Realized Profit** - Profit/Loss from completed (sold) transactions (FIFO-based)
- **Total Unrealized Profit** - Paper profit/loss on current holdings
- **Overall P&L** - Combined realized + unrealized profit
- **Return %** - Overall return percentage
- **Current Holdings Summary** - Live table showing all active holdings

### 2. **Transactions** (Master Log)
The heart of the system - where you enter ALL buy and sell transactions.

**Columns:**
- **Date** - Transaction date (DD/MM/YYYY or any date format)
- **Stock Name** - Name of the stock (use consistent naming!)
- **Transaction Type** - Dropdown: **Buy** or **Sell**
- **Quantity** - Number of shares
- **Price per Share** - Price at which you bought/sold
- **Brokerage/STT** - Additional charges (optional)
- **Total Amount** - Auto-calculated (Qty × Price + Charges)

**How to Use:**
1. Enter each transaction as a new row
2. Select "Buy" or "Sell" from the dropdown in column C
3. Total Amount calculates automatically
4. Keep transactions chronological for best results

### 3. **FIFO_Buy_Table** (The Engine)
The automated FIFO calculation engine - **DO NOT manually edit this sheet!**

**What it does:**
- Automatically extracts all Buy transactions sorted by date
- Tracks cumulative quantities
- Calculates how many shares from each lot have been sold (FIFO)
- Shows remaining quantity and value for each buy lot

**Columns:**
- **Stock Name** - Auto-populated from Transactions
- **Buy Date** - Date of purchase
- **Buy Qty** - Quantity purchased in this lot
- **Buy Price** - Price per share for this lot
- **Total Value** - Total investment in this lot
- **Cumulative Buy Qty** - Running total of purchases
- **Total Sold Qty** - Total shares sold for this stock
- **Sold from this Lot** - Shares from THIS lot that have been sold (FIFO)
- **Remaining Qty** - Unsold shares from this lot
- **Remaining Value** - Value of remaining shares at original buy price

**FIFO Logic:**
When you sell shares, they are consumed from the OLDEST lots first, exactly replicating real-world FIFO accounting.

### 4. **Holdings** (Current Portfolio)
Shows your current active holdings with live calculations.

**Columns:**
- **Stock Name** - Unique stocks with remaining quantity
- **Net Holding Qty** - Current total shares (Buys - Sells)
- **Current Avg Price** - Weighted average price of remaining lots
- **Total Invested Value** - Money invested in current holdings
- **CMP** - Current Market Price (YOU NEED TO ENTER THIS!)
- **Current Value** - Net Holding Qty × CMP
- **Unrealized P&L** - Current Value - Total Invested Value

**Note:** The CMP column is left blank for you to:
- Enter manually, OR
- Use Excel's Stock Data Type feature (right-click → Data Type → Stocks)

### 5. **Stock_Template** (Individual Stock Analysis)
A template sheet for detailed analysis of a single stock.

**How to Use:**
1. Enter the stock name in cell **B1** (replace "ENTER_STOCK_NAME_HERE")
2. The sheet automatically updates with:
   - **Realized P&L** - Profit from sold shares (FIFO-based)
   - **Unrealized P&L** - Paper profit on current holdings
   - **Current Holding Qty** - Shares you currently own
   - **Current Invested Value** - Money invested in current holdings
   - **Average Buy Price** - Weighted average of your buy prices
3. Below the banner, see **complete transaction history** for that stock

**To Track Multiple Stocks:**
- Right-click on the "Stock_Template" tab
- Select "Move or Copy"
- Check "Create a copy"
- Rename the new sheet (e.g., "RELIANCE", "TCS", etc.)
- Enter the stock name in B1
- Repeat for each stock you want to track

---

## 🚀 Getting Started

### Step 1: Open the File
Open `Stock_Portfolio_Tracker_FIFO.xlsx` in **Excel 365** (or Excel 2019+ with Office 365 subscription).

### Step 2: Enter Your Transactions
Go to the **Transactions** sheet and start entering your buy/sell history:

**Example:**
| Date | Stock Name | Transaction Type | Quantity | Price per Share | Brokerage/STT | Total Amount |
|------|-----------|------------------|----------|----------------|---------------|--------------|
| 01/01/2024 | RELIANCE | Buy | 10 | 2500 | 50 | 25050 |
| 15/01/2024 | TCS | Buy | 5 | 3500 | 30 | 17530 |
| 01/02/2024 | RELIANCE | Sell | 5 | 2700 | 40 | 13460 |
| 15/02/2024 | RELIANCE | Buy | 8 | 2600 | 45 | 20845 |

### Step 3: Enter Current Market Prices
Go to the **Holdings** sheet and enter the Current Market Price (CMP) for each stock in column E.

### Step 4: Check Your Dashboard
Go to the **Dashboard** sheet to see your complete portfolio summary!

### Step 5: Create Individual Stock Sheets (Optional)
Duplicate the **Stock_Template** sheet for detailed analysis of individual stocks.

---

## 💡 Key Features

### ✅ Fully Automated
- All calculations happen automatically via Excel formulas
- No manual calculations needed
- Updates in real-time as you add transactions

### ✅ True FIFO Implementation
- Sells consume from oldest lots first
- Accurate cost basis calculation
- Proper realized vs unrealized P&L tracking

### ✅ No Code Inside Excel
- Pure Excel 365 formulas (FILTER, UNIQUE, SUMIFS, INDEX, MATCH, etc.)
- No VBA macros
- No external dependencies
- Works on any Excel 365 installation

### ✅ Scalable
- Supports up to 500 buy transactions
- Unlimited sell transactions
- Multiple stocks in one portfolio

### ✅ Tax-Ready
- FIFO method is accepted by most tax authorities
- Clear lot-wise tracking
- Realized profit calculation for tax reporting

---

## 📊 Understanding FIFO

**FIFO = First-In, First-Out**

When you sell shares, they are matched against your oldest purchases first.

**Example:**
```
Buy 1: 10 shares @ ₹100 on Jan 1
Buy 2: 15 shares @ ₹120 on Jan 15
Sell: 12 shares @ ₹150 on Feb 1

FIFO Calculation:
- 10 shares from Buy 1 @ ₹100 = Cost ₹1,000
- 2 shares from Buy 2 @ ₹120 = Cost ₹240
- Total Cost = ₹1,240
- Sale Revenue = 12 × ₹150 = ₹1,800
- Profit = ₹1,800 - ₹1,240 = ₹560

Remaining Holdings:
- 13 shares from Buy 2 @ ₹120
- Average Price = ₹120
```

---

## 🔧 Troubleshooting

### Issue: Formulas show #SPILL! error
**Solution:** Make sure cells below/beside formula cells are empty. Excel 365 dynamic arrays need space to "spill" results.

### Issue: #NAME? error appears
**Solution:** You need Excel 365 or Excel 2019+ with Office 365. Older versions don't support FILTER, UNIQUE, SORT functions.

### Issue: Realized P&L seems incorrect
**Solution:** 
1. Check that all transactions are entered correctly
2. Verify stock names are spelled consistently
3. Ensure dates are in proper date format
4. Check the FIFO_Buy_Table to see lot allocations

### Issue: Holdings not showing a stock
**Solution:** 
1. Make sure you have Buy transactions for that stock
2. Verify Net Holding Qty > 0
3. Check if sells = buys (fully sold position won't appear)

### Issue: Dashboard metrics are zero
**Solution:**
1. Enter transactions in the Transactions sheet
2. Enter CMP values in the Holdings sheet
3. Make sure formulas haven't been accidentally deleted

---

## 📝 Best Practices

1. **Consistent Naming** - Always use the same name for each stock (e.g., "RELIANCE" not "Reliance" or "RIL")

2. **Enter All Transactions** - Include every buy and sell for accurate FIFO tracking

3. **Include Charges** - Add brokerage and STT in the Brokerage/STT column for accurate cost basis

4. **Chronological Order** - While not required, entering transactions in date order makes review easier

5. **Regular Updates** - Update CMP values regularly to see accurate unrealized P&L

6. **Backup Your File** - Save copies regularly, especially before major edits

7. **Don't Edit FIFO_Buy_Table** - This sheet calculates automatically. Manual edits will break formulas!

8. **Use Date Format** - Ensure dates are in proper Excel date format, not text

---

## 📈 Advanced Usage

### Using Excel Stock Data Types
1. Select the CMP column in Holdings sheet
2. Type a stock ticker symbol (e.g., RELIANCE.NS for NSE stocks)
3. Right-click → Data Types → Stocks
4. Click the icon that appears and select "Price" to auto-fill current price

### Creating Reports
1. Filter Transactions sheet by date range for period-specific reports
2. Use PivotTables on Transactions sheet for custom analysis
3. Copy Dashboard metrics to track performance over time

### Multiple Portfolios
1. Duplicate the entire workbook for different portfolios (Family, Trading, Long-term, etc.)
2. Or add a "Portfolio" column to Transactions and filter accordingly

---

## ⚠️ Limitations

1. **Excel 365 Required** - Dynamic array formulas (FILTER, UNIQUE, etc.) need Excel 365 or Excel 2019+
2. **500 Buy Transaction Limit** - FIFO_Buy_Table is pre-configured for 500 buy transactions (can be extended)
3. **Manual CMP Entry** - You need to manually enter or refresh current market prices
4. **Single Currency** - Designed for single currency (₹), modify formatting for other currencies
5. **No Dividends Tracking** - This tracker focuses on capital gains, not dividend income
6. **No Corporate Actions** - Splits, bonuses, mergers need manual adjustment

---

## 🎯 Formula Reference

### Key Formulas Used

**FILTER** - Extract rows matching criteria
```excel
=FILTER(array, include, [if_empty])
```

**UNIQUE** - Get unique values
```excel
=UNIQUE(array, [by_col], [exactly_once])
```

**SORT** - Sort data
```excel
=SORT(array, [sort_index], [sort_order])
```

**SUMIFS** - Conditional sum
```excel
=SUMIFS(sum_range, criteria_range1, criteria1, ...)
```

**INDEX/MATCH** - Lookup value
```excel
=INDEX(return_range, MATCH(lookup_value, lookup_range, 0))
```

---

## 📞 Support

For issues or questions:
1. Check the Troubleshooting section above
2. Verify you're using Excel 365 or Excel 2019+
3. Ensure all formulas are intact (haven't been accidentally deleted)
4. Review the FIFO_Buy_Table to understand lot allocations

---

## 📄 License

This Excel tracker is provided as-is for personal portfolio management. Feel free to modify and customize for your needs!

---

## 🔄 Version History

**Version 1.0** (May 2026)
- Initial release
- 5 sheets: Dashboard, Transactions, FIFO_Buy_Table, Holdings, Stock_Template
- Full FIFO implementation
- Excel 365 formula-based (no VBA)
- Support for 500+ buy transactions
- Realized and Unrealized P&L tracking

---

## ✨ Enjoy Tracking Your Portfolio! 📈

**Remember:** Always backup your file before making major changes!
