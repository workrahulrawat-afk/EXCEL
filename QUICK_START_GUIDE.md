# 🚀 Quick Start Guide - Stock Portfolio Tracker

## ✅ FIXED: No More Corruption Popup!

The file has been fixed to eliminate the corruption warning. You can now open it safely in Excel 365.

---

## 📥 Opening the File

1. **Download** `Stock_Portfolio_Tracker_FIFO.xlsx` from GitHub
2. **Right-click** the file → **Properties** → **Unblock** (if you see this option) → Click **OK**
3. **Open** the file in **Excel 365**
4. Click **Enable Editing** if prompted (yellow banner at top)
5. ✅ **Ready to use!** No corruption warnings!

---

## 🎯 5-Minute Setup

### Step 1: Add Your First Transaction (30 seconds)

1. Go to **Transactions** sheet
2. Click on Row 2
3. Enter your first buy:
   - **Date**: `01/01/2024`
   - **Stock Name**: `RELIANCE`
   - **Transaction Type**: Select `Buy` from dropdown
   - **Quantity**: `10`
   - **Price per Share**: `2500`
   - **Brokerage/STT**: `50` (optional)
   - **Total Amount**: *Auto-calculates to 25,050*

### Step 2: Add More Transactions (2 minutes)

Continue adding all your transactions:

```
Row 3: 15/01/2024 | TCS | Buy | 5 | 3500 | 30
Row 4: 01/02/2024 | RELIANCE | Sell | 5 | 2700 | 40
Row 5: 15/02/2024 | INFY | Buy | 20 | 1500 | 60
```

### Step 3: Enter Current Prices (1 minute)

1. Go to **Holdings** sheet
2. In Column E (CMP), enter current market prices:
   - RELIANCE: `2800`
   - TCS: `3600`
   - INFY: `1550`

### Step 4: Check Your Dashboard (30 seconds)

1. Go to **Dashboard** sheet
2. 🎉 See your complete portfolio summary!
   - Total Invested
   - Current Value
   - Realized Profit
   - Unrealized Profit
   - Overall P&L
   - Return %

---

## 🔍 Understanding FIFO

**Example Walkthrough:**

### Your Transactions:
1. **Buy 10 RELIANCE @ ₹2,500** on Jan 1 = ₹25,000
2. **Buy 8 RELIANCE @ ₹2,600** on Jan 15 = ₹20,800
3. **Sell 12 RELIANCE @ ₹2,800** on Feb 1 = ₹33,600

### FIFO Calculation (Automatic):
When you sell 12 shares:
- First 10 shares come from **Buy 1** @ ₹2,500 = Cost ₹25,000
- Next 2 shares come from **Buy 2** @ ₹2,600 = Cost ₹5,200
- **Total Cost** = ₹30,200
- **Sale Revenue** = ₹33,600
- **Realized Profit** = ₹3,400 ✅

### Remaining Holdings:
- 6 shares from **Buy 2** @ ₹2,600
- **Average Price** = ₹2,600
- **Invested Value** = ₹15,600

**Check the FIFO_Buy_Table sheet to see this breakdown!**

---

## 📊 Sheet-by-Sheet Guide

### 1️⃣ Dashboard (Your Control Center)
- **Don't edit** - Just view
- Auto-updates when you add transactions
- Shows overall performance

### 2️⃣ Transactions (Your Input Sheet)
- **This is where you work!**
- Enter every buy and sell
- Use dropdown for Transaction Type
- Total Amount calculates automatically

### 3️⃣ FIFO_Buy_Table (The Engine)
- **Don't edit** - Fully automated
- Shows which lots are sold (FIFO)
- See "Sold from this Lot" column
- Check "Remaining Qty" for unsold shares

### 4️⃣ Holdings (Current Portfolio)
- **View your active positions**
- Only stocks with Remaining Qty > 0
- **Enter CMP** in Column E (only manual input here)
- Unrealized P&L calculates automatically

### 5️⃣ Stock_Template (Individual Analysis)
- **Change B1** to stock name (e.g., RELIANCE)
- See complete transaction history for that stock
- View realized + unrealized P&L
- **Duplicate this sheet** for each stock you want to track

---

## 🎨 Pro Tips

### ✅ DO:
- **Use consistent stock names** (always "RELIANCE", not "Reliance" or "RIL")
- **Enter transactions chronologically** (easier to track)
- **Include brokerage** for accurate cost basis
- **Update CMP regularly** for current P&L
- **Backup your file** weekly

### ❌ DON'T:
- **Don't edit FIFO_Buy_Table** (auto-calculated)
- **Don't delete formula cells** (if you do, close without saving and reopen)
- **Don't use different names** for same stock (breaks tracking)
- **Don't edit Dashboard** (it's a view-only summary)

---

## 🧪 Test With Sample Data

Want to see how it works? Try this sample portfolio:

### Sample Transactions:
```
01/01/2024 | RELIANCE | Buy | 10 | 2500 | 50
05/01/2024 | TCS | Buy | 5 | 3500 | 30
10/01/2024 | INFY | Buy | 20 | 1500 | 40
15/01/2024 | RELIANCE | Buy | 8 | 2600 | 45
20/01/2024 | TCS | Sell | 2 | 3600 | 25
01/02/2024 | RELIANCE | Sell | 12 | 2750 | 50
05/02/2024 | INFY | Buy | 10 | 1520 | 30
```

### Sample Current Prices (in Holdings):
```
RELIANCE: 2800
TCS: 3650
INFY: 1580
```

**Expected Results:**
- Total Invested: ~₹40,000
- Realized Profit: Should show profit from TCS and RELIANCE sales
- Holdings: 6 RELIANCE, 3 TCS, 30 INFY

---

## 🆘 Troubleshooting

### Problem: Popup says "We found a problem with some content"
**Solution:** ✅ **FIXED!** Download the latest version from GitHub.

### Problem: #SPILL! error appears
**Solution:** 
- Clear cells below/beside the formula
- Excel needs empty cells for dynamic arrays to "spill" results

### Problem: #NAME? error appears
**Solution:** 
- You need Excel 365 or Excel 2019+ with Office 365
- FILTER, UNIQUE, SORT functions not available in older Excel

### Problem: Holdings not showing a stock
**Solution:**
- Check if Net Holding Qty > 0
- If fully sold, stock won't appear (correct behavior)
- Check stock name spelling is consistent

### Problem: Realized P&L seems wrong
**Solution:**
- Go to FIFO_Buy_Table sheet
- Check "Sold from this Lot" column
- Verify transactions are entered correctly
- Make sure stock names match exactly

### Problem: Dashboard shows #DIV/0!
**Solution:**
- Normal when no transactions entered yet
- Will disappear once you add transactions

---

## 📱 Excel Mobile / Web

**Excel Online (Web):**
- ✅ Full functionality
- ✅ All formulas work
- ✅ Dynamic arrays supported

**Excel Mobile (iOS/Android):**
- ⚠️ Limited - View only recommended
- ❌ Some dynamic formulas may not work
- 💡 Best to edit on desktop Excel 365

---

## 🎓 Learn More

- **Full Documentation:** See `README_Portfolio_Tracker.md`
- **FIFO Method:** Check "Understanding FIFO" section in README
- **Formula Reference:** All formulas explained in README
- **Advanced Features:** Template duplication, Excel Stock Data Types

---

## 📊 Example Workflow

### Morning Routine (5 minutes daily):
1. Open the tracker
2. Enter yesterday's transactions in **Transactions** sheet
3. Update CMP in **Holdings** sheet (or use Excel Stock Data Type)
4. Check **Dashboard** for portfolio performance
5. Save the file

### Monthly Review (15 minutes):
1. Review **Dashboard** metrics
2. Check individual stock sheets for detailed performance
3. Verify all transactions are entered
4. Export or screenshot Dashboard for records
5. Backup the file

---

## ✨ You're All Set!

**Your portfolio tracker is ready to use!**

Start with a few transactions, see the magic happen, then add your complete history.

**Questions?** Check the full README_Portfolio_Tracker.md for detailed explanations.

**Happy Tracking! 📈**
