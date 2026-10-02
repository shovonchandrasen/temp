NIC Report Generator - Web UI
==============================

QUICK START
-----------
1. Double-click: START_Web_Report_Processor.bat in the main folder.
2. Your browser opens at http://127.0.0.1:5000/
3. Input files:
   - If the three required report files are placed in the main folder alongside
     START_Web_Report_Processor.bat:
       * Inventory Activity Report <Mon YY>.xlsx
       * Closing Stock Report <Mon YY>.xlsx
       * Store Wise Stock Report-from-YYYY-MM-DD-to-YYYY-MM-DD.csv
     the application automatically loads them and displays their filenames
     instead of the upload box.
   - If the files are not in the main folder, the upload option is shown until
     all three files are selected/uploaded.
4. Click "Proceed":
   - If "Review mappings" is NOT ticked (default), the application automatically
     applies previously saved mappings from mapping_memory.json, generates the
     processed Excel reports, and goes directly to the "Download files" page.
   - If "Review mappings" IS ticked, the application goes to the "Review mappings"
     page first, and from there to the "Download files" page after clicking
     "Generate reports".
5. Under "Unmatched CSV Products" you will see only CSV products that do not
   match both Excel sheets and have a non-zero value in Opening Stock, Goods
   Receipt, Goods Return, Inventory Deducted, or Closing Stock.
6. Each mapping control is a searchable list: type a search term, then choose
   a product from the displayed list. Free text cannot be submitted as a
   mapping. Selecting a product in one Excel column automatically selects the
   same product in the other column when it is available there.
7. Tick "Lock row" after checking a mapping. A locked row is protected and its
   selected product cannot be used by another row in the same Excel column.
   Click "Edit mapping" if you need to change it; tick again to lock it.
8. "View locked mappings" opens a separate review window/panel. Only locked
   manual mappings are included in the run; exact-name matches are always
   processed automatically.
9. If a previous run has saved mappings for the current CSV names, a
   "Use saved mappings" checkbox appears. Select it to populate and lock all
   valid known mappings automatically.
10. Click "Generate reports". The two processed Excel files and a ZIP download
    are shown on the final page.

MAPPING MEMORY
--------------
Confirmed mappings are saved locally in:
    webapp\mapping_memory.json

This file stays on the same computer and is never uploaded anywhere. If you
change a previous mapping by pressing "Edit mapping", choosing a different
product, locking the row again, and generating reports, the saved entry is
updated for the next run.

FILES
-----
START_Web_Report_Processor.bat   - double-click in main folder to launch
webapp/app.py                    - Flask server and local mapping memory
webapp/processor.py              - validation, matching, mapping and output logic
webapp/templates/index.html      - professional animated web interface
webapp/mapping_memory.json       - saved product mappings
webapp/logic_update_for_excel.py - original script, kept untouched

The uploads and outputs folders are created automatically. Keep the command
window open while using the application; press Ctrl+C or close the window to
stop the server.
