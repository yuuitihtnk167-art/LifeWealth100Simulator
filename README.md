# LifeWealth 100 Simulator

LifeWealth 100 Simulator is a simple web app to estimate the asset balance at age 100
based on current assets, expected return, and monthly cash flow.

## Features
- Monthly simulation from today to age 100
- Imports asset totals from pasted table data
- Separate parameters for working years and retirement
- Investment contribution inputs and end ages
- Annual CSV export (same columns as asset trend data)

## How to use
1. Open `index.html` in your browser.
2. Set your birth date, current assets, and expected annual return.
3. Paste asset trend data (optional) and click **貼り付けデータで資産更新**.
4. Enter income/expense parameters for working and retirement periods.
5. Check the asset balance at age 100 and export the annual CSV if needed.

## Input formats
- Asset list and asset trend data are expected as pasted table text.
- The first row should contain headers.
- Tabs or commas are supported as delimiters.

## Notes
- Contributions move cash into investments; they do not increase total assets by themselves.
- Warnings appear if cash balance becomes negative.

