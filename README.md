# Problem
A barcode scanner fills a spreadsheet but does not maintain an inventory of the items scanned.

# No problem!
Just use this Excel macro!

## How it works
You have two spreadsheets, source and destination. The source spreadsheet collects items scanned with the barcode scanner. The destination spreadsheet contains your inventory. The problem is, you want the new scans to update the inventory.

### Prepare the solution
Add the code in the `macros.bas` file to your Excel macros.
Edit the variable assignments in the top-level code block to adapt for your spreadsheets.

### Run the solution
Press the `Ctrl`+`G` keyboard shortcut (or run the `ProcessNewScans` macro via the Excel UI). You will be prompted before any changes are made. If you agree, the quantity field in your destination spreadsheet will be updated based on the new items in the source spreadsheet.
