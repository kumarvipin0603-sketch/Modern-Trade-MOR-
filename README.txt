PO FULFILMENT CONTROL TOWER V14 — SAFE SALE REGISTER REPAIR

WHY REPAIR EXISTS
Repair/Rebuild is NOT part of the normal daily workflow.
It is only needed when Sale Register data was previously loaded by an older
dashboard version that mapped fields such as Product/Item No., Line Amount,
GST or Gross Amount incorrectly.

V14 CHANGES
- Repair button appears only when the dashboard detects old/incomplete Sale Register mapping.
- If the Sale Register is healthy, the app says "No Repair/Rebuild is required."
- Fixed SQLite "too many SQL variables".
- Sale Register staging now uses safe executemany batching instead of one giant multi-value SQL statement.
- Item Ledger staging uses the same safe batching.
- Repair builds the staging table first and only replaces live Sale Register data after staging succeeds.
- Normal daily incremental Sale Register uploads remain unchanged.

RUN
1. Stop Streamlit: Ctrl+C
2. Replace app.py with V14.
3. Keep control_tower_data unchanged.
4. Run:
   py -m streamlit run app.py

If V14 detects old bad Sale Register data, run "Repair Old Sale Register Data" once.
After that, repair should no longer be required.


V15 SALES ORDER MAPPING FIX
- Confirmed ERP Sales Order No. = Column A.
- Confirmed Customer PO No. = Column F.
- The same Sales Order file can now be reprocessed even if its checksum already exists.
- This is required to correct files previously processed with the old Column-D rule.
- Upload History now shows the current count of unique PO -> ERP Sales Order mappings.

HOW TO CORRECT EXISTING SALES ORDER DATA
1. Start V15.
2. Go to Upload Centre -> Sales Orders.
3. Select the same latest Sales Orders Excel.
4. Click Process Sales Orders.
5. V15 will NOT reject it merely because it was already uploaded.
6. Confirm the success message says Column A and Column F.


V16 — DASHBOARDS ALWAYS REFLECT AVAILABLE DATA

MAIN RECONCILIATION
- No PO search = show all available rows from uploaded Sale Register + SO mapping + Blocked + FG stock + PO + GRN.
- Sale Register rows are used directly as the base if Customer PO is not yet uploaded.
- Visible source row counters added.

FACTORY STOCK REQUIREMENT
- Shows all ERP items / FG stock even if PO pending qty is not yet available.
- Pending / shortage remain blank until enough PO quantity information exists.
- Does not show an empty page merely because Customer PO is missing.

SALES & RETURN 360
- Reads directly from consolidated Sale Register.
- If Sale Register has rows, the dashboard reflects them.
- Legacy Sale Register field names are backfilled at dataframe level.

AUDIT
- Added Database Rows summary for each source to verify what is actually loaded.

RUN
1. Stop Streamlit: Ctrl+C
2. Replace app.py with V16
3. KEEP control_tower_data
4. py -m streamlit run app.py


V17 — SALE REGISTER RECOVERY

WHY SCREEN CAN SHOW:
- Upload History: Sale Register Processed 37,288
- Main Dashboard: Sale Register 0 rows

Upload History is an audit log. It records that a file was processed at that time.
The live dashboard reads the sale_register SQLite table.
An older failed Repair/Rebuild could empty the live table while leaving Upload History unchanged.

V17 DETECTS THIS CONDITION AUTOMATICALLY.
If:
- live Sale Register = 0 rows, AND
- Upload History shows a previous Sale Register with rows,
V17 shows a recovery warning.

If the original stored workbook still exists under control_tower_data/uploads,
click:
Recover Sale Register From Stored Upload

No re-upload is needed.

Recovery is safe:
1. New staging table is built first.
2. Live table is replaced only after staging succeeds.
3. Uses safe 1000-row batches.


V18 — SALES ORDER COLUMN F LOCK

CONFIRMED RULE
- Column A = ERP Sales Order No.
- Column F = Customer PO No.

V18 does not infer these two fields from header names.
It always reads physical Excel Column A and Column F.

IMPORTANT
The same Sales Order file can be uploaded and processed again even if it
was already uploaded earlier. This allows old incorrect Column-D mappings
to be overwritten by the corrected Column-F mapping.

The Sales Orders tab also displays a preview:
Column A = <header>
Column F = <header>

After processing, the success message explicitly confirms Column A and Column F.


V19 — SALES ORDER MAPPING ONLY

CONFIRMED BUSINESS RULE
- Sales Order Excel Column A = ERP Sales Order No.
- Sales Order Excel Column D = Customer PO reference.

MAIN RECONCILIATION
- Match Main Reconciliation Po Number against Sales Order Excel Column D.
- Populate ONLY the Sales Order No. field from Column A.
- Do NOT take Ledger, Item, Qty, Price, Date, Stock, Blocked Qty or any other field from the Sales Order Excel.
- Those fields continue to come from their own source files.

REPROCESSING
The same Sales Order file can be processed again so old incorrect mappings can be overwritten.


V20 — PERFORMANCE BUILD

DASHBOARD SPEED
- Main Reconciliation is now vectorized: source tables are aggregated once and merged.
- Removed per-invoice repeated database filtering.
- Main dashboard is paginated (100/250/500/1000 rows per page).
- Full download still contains every filtered row.
- Sales & Return 360 renders a selectable preview while download retains all rows.
- Short-lived Streamlit caches make navigation between dashboards much faster.
- Cache is cleared automatically after uploads/updates.

UPLOAD SPEED
- Uses python-calamine for XLSX reading when available; openpyxl remains fallback.
- Sale Register and Item Ledger retain safe batch SQLite loading.
- SQLite memory cache and mmap are enabled.
- Daily incremental uploads remain the recommended workflow after the initial FY load.

BUSINESS RULE RETAINED
Sales Order Excel is mapping-only:
- Column A = ERP Sales Order No.
- Column D = Customer PO reference.
Only Main Reconciliation Sales Order No. is updated from this source.


V21 — FACTORY REQUIREMENT BRANCH FILTER

Factory Stock Requirement now includes a Branch filter:
- All Branches
- Each individual branch available in Item Ledger

ALL BRANCHES
- Pending billing qty across all branches
- FG stock summed across all branches

INDIVIDUAL BRANCH
- Pending billing rows restricted to selected Branch Code
- FG stock restricted to selected branch and FG locations only
- Blocked qty and factory shortage recalculated for that branch
- Download file includes the selected branch name


V22 MAIN RECONCILIATION SUMMARY
Only these top metrics are shown:
- Sale Qty
- Sale Value
- Return Qty
- Return Value
- GRN Qty
- Short GRN
- Tracking IDs

Removed old source/operational counters from the Main Reconciliation top area.

PO Search:
- Blank / Show All = all uploaded reconciliation data.
- Search a PO = summary and detail only for that PO.


V23 MAIN RECONCILIATION SUMMARY
Top Reconciliation Summary now shows only:
- Sale Qty
- Sale Value
- Return Qty
- Return Value

Removed from Reconciliation Summary:
- GRN Qty
- Short GRN
- Tracking IDs

GRN and tracking information remain available in the detailed reconciliation rows where applicable.
PO search behavior remains:
- Blank / Show All = all uploaded reconciliation data.
- Search a PO = summary and detail for that PO only.


V24 AUTO-HEIGHT RECONCILIATION TABLE
- Reconciliation Detail height now matches the actual number of visible rows.
- 1 row of data = compact 1-row table.
- 4 rows of data = compact 4-row table.
- No large blank grid area below small PO search results.
- For larger datasets, the table grows up to a 600px maximum and then scrolls.


V25 — MULTI-PO SEARCH + LEDGER FILTER

MAIN RECONCILIATION PO SEARCH
- One PO: PO001
- Multiple PO: PO001, PO002, PO003
- Also accepts semicolon or new-line separated PO numbers.
- Duplicate PO numbers in the search are ignored.
- Blank search / Show All = all uploaded reconciliation data.

MAIN RECONCILIATION LEDGER FILTER
- Added Ledger dropdown.
- Ledger options are based on the current PO search scope.
- All Ledgers = no ledger restriction.
- Choosing a Ledger filters both the Reconciliation Summary and Detail.

COMBINED EXAMPLE
Search:
PO001, PO002, PO003
Then choose one Ledger.
The Summary, Reconciliation Detail, GRN working view and blocked detail all reflect that selected scope.


V26 — ONLY PENDING QTY BY LEDGER

MAIN RECONCILIATION
- Added "Only Pending Qty" click button.
- Added "Show All Rows" reset button.
- Pending filter is applied after PO search and Ledger selection.
- Pending means Pending Billing Qty > 0.
- If a Ledger is selected, only pending rows for that Ledger are shown.
- If All Ledgers is selected, pending rows across all Ledgers are shown.
- Pending-only view also shows:
  * Pending PO References
  * Total Pending Qty

Examples:
1. Select Ledger = RetailEZ
2. Click Only Pending Qty
3. Dashboard shows only RetailEZ rows where Pending Billing Qty > 0.

The same works with one PO, multiple POs, or all POs.


V27 — LEDGER NAME FROM SALE REGISTER COLUMN BM

CONFIRMED RULE
- Ledger Name must come from ERP Sale Register Excel Column BM.
- Excel BM = 65th column = pandas index 64.
- Main Reconciliation Ledger filter uses this field.

EXISTING DATA
- Added "Refresh Ledger Names from Column BM" in:
  Upload Centre -> Sale Register
- This allows already-loaded Sale Register rows to update their Ledger Name
  from the stored original workbook without re-uploading the full file.

NEW / DAILY SALE REGISTER UPLOADS
- Ledger Name is automatically read from Column BM.


V28 — LEDGER FILTER CORRECTION

ISSUE FIXED
Names such as Abhishek, Arpita Singh, Aseem Bansal and Asim were appearing
because the previous build forced a physical dataframe column position.

CORRECT RULE
- Read ONLY the Sale Register column whose header is exactly: Ledger Name
- Do not use Customer Name, Salesperson, User, Employee or any fallback field.
- If exact Ledger Name header is missing, Sale Register processing stops with an error instead of using a wrong field.

EXISTING DATA
Go to:
Upload Centre -> Sale Register -> Correct Ledger Names from Sale Register

Run this once after installing V28. It refreshes existing database ledger_name
values from the exact Ledger Name field in the stored Sale Register.


V29 — LEDGER FILTER FROM PHYSICAL COLUMN BM

LOCKED RULE
- Ledger Name is taken ONLY from physical Excel Column BM in the ERP Sale Register.
- Excel BM = 65th physical column.
- The app captures BM before pandas concatenates/reorders sheet columns.
- Column I, header-based Ledger Name, salesperson, user or customer-name fields are ignored.

EXISTING DATABASE
After starting V29:
Upload Centre -> Sale Register -> Refresh Ledger Names from Physical BM

Run once to correct the already-loaded ledger_name values.

NEW SALE REGISTER UPLOADS
Every new/incremental Sale Register automatically stores Ledger Name from physical BM.


V30 — LEDGER BM FULL HISTORY FIX

Why wrong names remained:
Earlier correction used only the latest Sale Register upload. Since the Sale Register
is incremental, older consolidated rows kept their previous wrong ledger values.

V30:
- Reads all stored ERP Sale Register uploads.
- Captures physical Excel Column BM only.
- Re-verifies consolidated rows using source_key.
- Ledger dropdown contains only BM-verified values.
- Unverified historical values are excluded from the Ledger dropdown.

Run once:
Upload Centre -> Sale Register -> Correct Ledger Filter from ALL Sale Registers


V31 — STRICT DUPLICATE CONTROL FOR SALE / RETURN

RULE
Any duplicate Sale, Invoice, CN or SR transaction must NOT be inserted again.

Stable duplicate identity:
PO No + Document No + ERP Item + Signed Qty + Line Amount + Gross Amount + Document Type

V31:
- Adds business_key independent of parser/version metadata.
- Creates a unique SQLite index on business_key.
- INSERT OR IGNORE prevents duplicate rows in every future upload.
- Automatically cleans historical duplicate transactions at startup.
- Sale & Return 360 and Main Reconciliation also de-duplicate in memory as a safety net.
- Added Upload Centre -> Sale Register -> Clean Existing Sale / Return Duplicates.

This specifically prevents a repeated return document from being counted twice while its sale is counted once.


V32 — SHIP-TO LOCATION CODE MASTER

NEW UPLOAD
Upload Centre -> Ship-to Location Master

Required Excel columns:
- Ledger Name
- Pin Code
- Ship to Location Code

Optional:
- Ship to Location Name

MATCHING RULE
Ship to Location Code = Ledger Name + 6-digit PIN Code

PIN can be picked automatically from:
1. Customer PO Ship to Location / address
2. Sale Register Post Code
3. Sale Register Ship-to Address 1 / Address 2

MAIN RECONCILIATION
New column:
- Ship to Location Code

MASTER MAINTENANCE
- Same Ledger + PIN updates existing code
- New Ledger + PIN adds a new row
- No duplicate Ledger + PIN records
- Download current master / blank template as Excel


V33 — CUSTOMER PO EXPIRY / DELIVERY DATE + SHIP-TO GST

Main Reconciliation now includes:
- PO Expiry/DELIVERY DATE
- Ship to GST no as per PO

SOURCE
These two values come only from the uploaded Customer PO when available.
If the customer PO does not contain a value, the Main Reconciliation cell stays blank.

EXCEL PO ALIASES
PO Expiry/DELIVERY DATE supports common headings including:
- PO Expiry/DELIVERY DATE
- PO Expiry Date
- PO Delivery Date
- Delivery Date
- Expiry Date
- PO Expiry / Delivery Date

Ship-to GST supports common headings including:
- Ship to GST no as per PO
- Ship to GST No as per PO
- Ship to GST No.
- Ship To GSTIN
- Ship to GSTIN
- GSTIN Ship To

EXISTING DATABASE
V33 adds the two new PO columns automatically. Existing data is preserved.

NOTE ON PDF
The current Python tower stores Customer PO PDFs but does not yet have a universal PDF line parser.
These fields are fully active for Excel PO uploads; customer-specific PDF extraction can populate the
same database fields when those PDF parsers are configured.


V34 — FINANCIAL YEAR FILTER

GLOBAL FILTER
- All
- 2025
- 2026

FY DEFINITION
- 2025 = 01-Apr-2025 to 31-Mar-2026
- 2026 = 01-Apr-2026 to 31-Mar-2027

APPLIES TO
1. Main Reconciliation
   - Invoice Date used first
   - PO Date used as fallback for unbilled PO rows
   - Works with single PO, multiple PO, Ledger and Pending Qty filters

2. Factory Stock Requirement
   - Pending order requirement is filtered by FY
   - Current FG stock remains latest current stock (not historical FY stock)

3. Sales & Return 360
   - Sales and returns are filtered by Invoice Date FY

All = complete uploaded history.


V35 — SALE / RETURN CONTROL TOTAL FIX

Reviewed against ERP Sale Register business logic.

CORRECT CALCULATION
Sale Qty   = unique genuine Invoice/Sale rows -> Quantity
Sale Value = unique genuine Invoice/Sale rows -> Gross Amount
Return Qty = unique CN/SR/Sales Return rows -> ABS(Quantity)
Return Value = unique CN/SR/Sales Return rows -> ABS(Gross Amount)

IMPORTANT FIXES
- Main Reconciliation summary now calculates directly from the consolidated Sale Register,
  not from enriched reconciliation rows.
- Return rows can no longer be counted as Sales.
- SR-prefixed documents and Sales Return/Return document types are classified as Returns.
- Existing business_key duplicate protection remains active.
- FY filter is applied on Invoice Date before totals.
- FY2025 full-control benchmark added:
  Sale Qty = 280,023
  Sale Value = 1,022,327,994


V36 — SINGLE SALE REGISTER SOURCE

LOCKED ARCHITECTURE
Main Reconciliation Summary and Sales & Return 360° now call the SAME function:
filtered_sale_register_source()

NO reconciliation join is allowed in Sale/Return KPI calculation.

ORDER OF CALCULATION
1. Consolidated ERP Sale Register
2. Business-key duplicate removal
3. Financial Year filter on Invoice Date
4. Optional PO / Ledger / Branch / SKU filters
5. Classify genuine Sale vs CN/SR/Return
6. Calculate:
   Sale Qty     = SUM sale Quantity
   Sale Value   = SUM sale Gross Amount
   Return Qty   = SUM ABS(return Quantity)
   Return Value = SUM ABS(return Gross Amount)

FY2025 CONTROL
When FY=2025 and no PO/Ledger/Branch/SKU filter is applied:
- Sale Qty must equal 280,023
- Sale Value must equal 1,022,327,994

Both Main Reconciliation Summary and Sales & Return 360° now display a
control-match / control-mismatch message using the same calculation.

AUDIT
Sales & Return 360° now provides:
Download Sale Register Control Extract
This downloads the exact unique Sale Register rows used for the current totals.


V37 — EXACT ERP INVOICE / CREDIT MEMO LOGIC

ROOT CAUSE FOUND
The previous build incorrectly treated any row with Return Order No. as a Return.
In the actual ERP Sale Register, an Invoice row can contain the linked SR number
in Return Order No. That field is a cross-reference, not the transaction type.

LOCKED SOURCE COLUMNS
Excel Column P  = Quantity
Excel Column AA = Gross Amount
Excel Column BH = Document Type

LOCKED CALCULATION
Document Type = Invoice:
  Sale Qty   = SUM Quantity
  Sale Value = SUM Gross Amount

Document Type = Credit Memo:
  Return Qty   = SUM ABS(Quantity)
  Return Value = SUM ABS(Gross Amount)

Return Order No. is NEVER used to decide Sale vs Return.

CONTROL TOTALS — ALL FINANCIAL YEARS
Sale Qty:     420,745
Sale Value:   1,594,938,672
Return Qty:   55,276
Return Value: 192,850,233

CONTROL TOTALS — FY2025
Sale Qty:   280,023
Sale Value: 1,022,327,994

Both Main Reconciliation Summary and Sales & Return 360 use the same calculation.

EXISTING DATABASE
Upload Centre -> Sale Register -> Repair Sale / Return Classification
Run once after installing V37 to clean legacy CN helper fields.


V38 — ROW-PRESERVING SALE REGISTER REBUILD

ROOT CAUSE OF LOW FY2026 TOTALS
V31-V37 used a business-level duplicate key:
PO + Invoice + SKU + Qty + Value + Document Type.

That key could treat genuine repeated ERP Sale Register rows as duplicates.
Once those rows were removed from the SQLite database, dashboard totals became
much lower than the original Sale Register.

V38 DUPLICATE RULE
- Genuine identical-looking ERP line occurrences are preserved using occurrence numbers.
- Re-uploading the same source rows generates the same occurrence keys and is ignored.
- No dashboard-level drop_duplicates is applied to Sale / Return totals.

ONE-TIME REPAIR
Upload Centre -> Sale Register -> Rebuild Full Sale Register (V38)

The rebuild:
1. Clears only the consolidated sale_register table.
2. Re-reads all stored ERP Sale Register source files.
3. Rebuilds row-preserving occurrence keys.
4. Ignores overlapping/re-uploaded source occurrences.
5. Leaves PO / SO / Stock / Blocked / GRN tables untouched.

FY2026 CONTROL
Sale Qty:     140,722
Sale Value:   572,610,678
Return Qty:   13,043
Return Value: 45,410,167

FY2025 CONTROL
Sale Qty:   280,023
Sale Value: 1,022,327,994

ALL CONTROL
Sale Qty:     420,745
Sale Value:   1,594,938,672
Return Qty:   55,276
Return Value: 192,850,233


V39 — SALE REGISTER REBUILD UNIQUE-CONSTRAINT FIX

V38 ERROR
Full Sale Register rebuild failed:
UNIQUE constraint failed: sale_register.business_key

CAUSE
V38 temporarily dropped the business_key unique index, loaded every stored
Sale Register file including overlapping incremental rows, and attempted to
create the unique index only after loading was complete. The database then
contained repeated business_key values, so index creation failed.

V39 FIX
- sale_register is cleared for the rebuild
- occurrence-aware UNIQUE business_key index is created BEFORE any file loads
- each stored Sale Register is imported in upload order
- genuine identical lines inside one ERP file remain as occurrence 1,2,3...
- if the same rows appear again in a later/incremental upload, INSERT OR IGNORE
  prevents them from being loaded again
- no UNIQUE-index creation failure at the end

Run once:
Upload Centre -> Sale Register -> Rebuild Full Sale Register (V39)


V40 — SOURCE_KEY OCCURRENCE FIX

ROOT CAUSE FOUND AFTER V39
The original sale_register table has:
    source_key TEXT UNIQUE

V38/V39 changed business_key to an occurrence-aware key, but source_key still
used the old transaction-level hash. Therefore genuine repeated ERP rows with
the same transaction signature were STILL rejected by SQLite on source_key.

V40 FIX
- source_key = occurrence-aware business_key
- both SQLite unique constraints now preserve genuine repeated source lines
- same rows uploaded again generate the same occurrence-aware keys and are ignored
- full rebuild button now reports FY2025, FY2026 and ALL control totals immediately

EXPECTED CONTROLS
FY2025:
  Sale Qty 280,023
  Sale Value 1,022,327,994

FY2026:
  Sale Qty 140,722
  Sale Value 572,610,678
  Return Qty 13,043
  Return Value 45,410,167

ALL:
  Sale Qty 420,745
  Sale Value 1,594,938,672
  Return Qty 55,276
  Return Value 192,850,233

ONE-TIME ACTION
Upload Centre -> Sale Register -> Rebuild Full Sale Register (V40)


V41 — FINAL CONSOLIDATION + PERFORMANCE BUILD

SALE REGISTER SOURCE LOCK
- Quantity = physical Excel Column P
- Gross Amount = physical Excel Column AA
- Document Type = physical Excel Column BH
- Ledger Name = physical Excel Column BM
- Invoice = Sale
- Credit Memo = Return

DUPLICATE RULE
- The duplicate identity is now the exact complete Excel source row.
- Exact duplicate rows in the same file or a later upload are ignored.
- Different ERP rows are preserved even when PO / Invoice / SKU / Qty / Value happen to match.
- Old business-level and occurrence-level duplicate logic is no longer used for new/rebuilt data.

ONE-CLICK REPAIR
Upload Centre -> Sale Register -> One-Click Optimize & Rebuild Sale Register
- Reads every stored Sale Register source file.
- Builds a complete staging dataset first.
- Removes only exact duplicate rows.
- Replaces live Sale Register only after staging succeeds.
- Rebuilds performance indexes.
- Clears Streamlit caches.
- Shows FY2025 / FY2026 / ALL control totals immediately.

PERFORMANCE
- Heavy dashboard data remains cached until an upload/update invalidates it.
- Removed 8-second cache expiry that forced frequent recalculation.
- Added compound SQLite indexes for FY, PO, Ledger, Branch, SKU, GRN, blocked shipment and FG-stock queries.
- Main Reconciliation and Sales & Return KPI cards use direct SQLite aggregation.
- No second dashboard-level Sale Register deduplication.
- Successful uploads invalidate cache and refresh automatically.

EXPECTED CONTROL TOTALS
FY2025:
  Sale Qty 280,023
  Sale Value 1,022,327,994

FY2026:
  Sale Qty 140,722
  Sale Value 572,610,678
  Return Qty 13,043
  Return Value 45,410,167

ALL:
  Sale Qty 420,745
  Sale Value 1,594,938,672
  Return Qty 55,276
  Return Value 192,850,233


V42 — CUSTOMER PO PDF PARSER / WALMART INDIA

WALMART PO PDF
The Customer PO upload now parses Walmart India PDF purchase orders directly.

Captured from PDF:
- PO No.
- PO Date
- PO Cancel Date -> PO Expiry/DELIVERY DATE
- Ledger / Customer
- Customer Article -> PO Item
- ERP Item from Customer SKU & Price Master
- Article Description
- PO Qty
- PO Unit Price / Cost
- PO Value = line Total Amount incl. taxes
- Ship-to Location
- Ship-to GST no as per PO

EXAMPLE VERIFIED
Walmart PO 6700070365:
- PO Date: 07-Aug-2026
- PO Cancel / Expiry: 06-Sep-2026
- Ship To: Cash N Carry - Amristar, AMRITSAR-143109
- Ship-to GST: 03AADCB2110L1ZA
- 6 PO article lines

EXISTING STORED PDF
Upload Centre -> Customer PO -> Reprocess Stored Customer PO PDFs

This allows PDFs uploaded by earlier builds (which only stored the PDF) to populate
po_lines without uploading the file again.

MASTER MAPPING
If Walmart ledger spelling differs from SKU master spelling, V42 first checks exact
ledger, then blank-ledger master, then a unique Customer Item -> ERP Item mapping.


V43 — PDF DEPENDENCY FIX

PROBLEM
ModuleNotFoundError: No module named 'pdfplumber'

FIX
- pdfplumber is now an optional dependency at app startup.
- Missing pdfplumber no longer crashes the complete dashboard.
- Customer PO tab shows an "Install PDF Parser" button.
- One-click installer uses the SAME Python interpreter running Streamlit.
- All Excel uploads and dashboards continue working even before PDF parser installation.

RECOMMENDED ONE-TIME WINDOWS SETUP
Option 1:
Double-click INSTALL_REQUIREMENTS.bat

Option 2:
Command Prompt in the tower folder:
py -m pip install -r requirements.txt

Then:
py -m streamlit run app.py

If the dashboard is already open without pdfplumber:
Upload Centre -> Customer PO -> Install PDF Parser


V44 — FLIPKART EXCEL FIXED-CELL PO PARSER

LOCKED FLIPKART MAPPING
PO No:
  B2

PO Expiry / Delivery Date:
  Q2

Ship-to Location:
  N5

Ship-to GST No:
  U5

LINE ITEMS FROM ROW 11 DOWNWARD
Customer Item Code:
  Column C (C11 downward)

PO Qty:
  Column D (D11 downward)

PO Value:
  Column W (W11 downward)

ERP ITEM
Customer Item is mapped through Customer SKU & Price Master.

LEDGER
Normalized to Flipkart India Pvt. Ltd. when Ship-to text identifies Flipkart.

IMPORTANT
PO Date fixed cell has NOT been guessed because it was not confirmed in the
Flipkart layout instruction. Existing Sale Register PO/invoice dates remain available
in reconciliation, while Customer PO PO Date stays blank until its exact cell is confirmed.

EXISTING FLIPKART EXCEL FILES
Upload Centre -> Customer PO -> Reprocess Stored Flipkart PO Excels


V45 — CUSTOMER ITEM -> ERP ITEM FOR ALL CUSTOMER POS

LOCKED RULE
Every customer PO will contain Customer Item Code.
The dashboard must always derive ERP Item Code from Customer SKU & Price Master.

FLOW
Customer PO Customer Item Code
    -> Customer SKU & Price Master
    -> ERP Item Code
    -> Main Reconciliation / Stock / Blocked / Invoice / GRN matching

APPLIES TO
- Generic Excel Customer POs
- Flipkart Excel fixed-layout POs
- Walmart India PDF POs
- Future customer PO parsers using the central resolver

MASTER PRIORITY
1. Exact Ledger + Customer Item Code
2. Blank-ledger Customer Item mapping
3. Unique Customer Item -> ERP Item mapping across the master

EXISTING POs
Upload Centre -> SKU & Price Master -> Refresh ERP Item Codes in All POs

Also, every SKU Master upload now automatically refreshes ERP Item Code
against all existing PO lines.

IMPORTANT
Customer Item Code is retained in PO Item for customer reference.
ERP Item Code is the matching key used to reconcile against ERP Sale Register,
Item Ledger, Shipment Not Invoiced and GRN.


V46 — FLIPKART PO DATE LOCK

Flipkart fixed-cell mapping now includes:
- B2 = PO No.
- V2 = PO / Order Date
- Q2 = PO Expiry / Delivery Date
- N5 = Ship-to Location
- U5 = Ship-to GST No.
- C11 downward = Customer Item Code
- D11 downward = PO Qty
- W11 downward = PO Value

Existing stored Flipkart Excel POs can be corrected via:
Upload Centre -> Customer PO -> Reprocess Stored Flipkart PO Excels


V47 — PO DETAILS MAPPING MASTER

PURPOSE
Different customers send POs in different Excel/PDF formats.
The dashboard now has a configurable Mapping Master so new layouts do not
require Python code changes for every customer.

LOCATION
Upload Centre -> Customer PO -> PO Details Mapping Master

MAPPING PROFILE
Create one Profile Name per customer/layout version.

SUPPORTED EXCEL SOURCES
- CELL: fixed cell, e.g. B2 / V2 / N5
- COLUMN: line data down a column from Start Row, e.g. C from row 11
- CONSTANT

SUPPORTED PDF SOURCES
- REGEX: extract a header field from parsed PDF text
- TABLE_COLUMN: extract line data from a PDF table column
- CONSTANT
- Optional Extract Regex can pull a customer item/code from a table cell.

PROFILE DETECTION
Excel:
- Detector Cell + Detector Contains

PDF:
- Detector Contains anywhere in parsed PDF text

STANDARD TARGET FIELDS
- PO No
- PO Date
- PO Expiry/DELIVERY DATE
- Ship to Location
- Ship to GST no as per PO
- Ledger Name
- Customer Item Code
- Item Description
- PO Qty
- PO Unit Price
- PO Value

ERP ITEM RULE
Customer Item Code is extracted from every PO.
ERP Item Code is ALWAYS fetched from Customer SKU & Price Master.
Do not map ERP Item Code from a customer's PO layout.

PROCESSING PRIORITY
1. PO Details Mapping Master
2. Existing customer-specific fallback parser (Flipkart/Walmart)
3. Generic Excel header parser

This makes future customer/layout additions an admin mapping task rather than
a Python-code change.


V48 — GRN DETAILS MAPPING MASTER

PURPOSE
Every customer can send GRN in a different Excel/PDF format.
GRN extraction is now configurable through an admin-maintained mapping Excel.

LOCATION
Upload Centre -> GRN -> GRN Details Mapping Master

SUPPORTED TARGET FIELDS
- PO No
- Ledger Name
- Invoice No
- Invoice Date
- ERP Item Code
- Customer Item Code
- Item Description
- Invoice Qty
- Transporter
- Docket No
- GRN No
- GRN Date
- GRN Qty
- Delivery / Invoice Cancel Date
- Delivery Remarks
- Short Delivered
- MIR No
- Sumit Invoice Upload
- POD Remarks
- Status

SUPPORTED EXCEL SOURCES
- CELL
- COLUMN from Start Row
- CONSTANT

SUPPORTED PDF SOURCES
- REGEX
- TABLE_COLUMN
- CONSTANT

ERP SKU RULE
If ERP Item Code exists in the GRN, use it.
If only Customer Item Code exists, the dashboard can resolve ERP Item Code
through Customer SKU & Price Master.

CN RULE
CN No / CN Qty / CN Value are NOT sourced from GRN mapping.
They continue to come from ERP Sale Register Credit Memo / return rows.

PROCESSING PRIORITY
1. Active GRN Mapping Master profile
2. Existing generic GRN Excel header parser
3. Unmatched PDF is stored and shown as mapping-required

EXISTING FILES
Upload Centre -> GRN -> Reprocess Stored GRN Files


V49 — PDF PO MAPPING FIX

Resolved errors:
- Metro PDF: item table is on page 2, not page 1.
- Dawntech Amazon PDF: item table is table 3, not table 0.
- RetailEZ Amazon PDF: item table is table 3, not table 0.
- Myntra PDF: pdfplumber exposes only the item table header; V49 parses SKU rows
  from page text directly.

V49 behavior:
1. Extracts tables from ALL PDF pages.
2. Scores every table against the active Customer Item / Qty / Value mapping.
3. Automatically selects the table that yields valid PO rows.
4. Supports optional Page No / Table No columns in PO Mapping Master.
5. If Page/Table hints are stale, automatically falls back to table scoring.
6. Myntra uses a dedicated text-row fallback:
   Customer Item = SKU Code (GGG...)
   Qty = quantity after Style ID
   PO Value = Total plus Taxes from the SKU block.

After installing V49:
Upload Centre -> Customer PO -> Process PO Mapping Master
(upload the reviewed mapping master once)
then reprocess/upload the affected POs.


V50 — PO HEADER / SHIP-TO / ERP ITEM FIX

Corrected:
- Blinkit: PO Date, R.O. Expiry, Delivered-To address/GST; joins line-broken Item Codes.
- Scootsy: PO Date, Expiry, Shipping Address/GST.
- Zepto: Shipping Address/GST.
- CP Wholesale 1733173: customer Ship-to only from CPWI Pvt Ltd.-GNI warehouse block;
  creation date and expiry date mapped.
- Metro 5115322648: only the Delivery Address block is used as Ship-to.
- ERP mapping now canonicalizes customer item codes (spaces/newlines/.0 artifacts)
  before matching Customer SKU & Price Master.

Reprocess affected Customer POs after installing V50.


V51 — CUSTOMER PO VISIBILITY IN MAIN RECONCILIATION

ROOT CAUSE
Main Reconciliation used Sale Register invoice rows as its only row base and
joined Customer PO only on PO No + ERP Item. Therefore uploaded POs remained
blank when:
- ERP Item mapping was still missing,
- Customer Item did not yet map to ERP SKU,
- or the PO had not yet been invoiced.

V51 FIX
1. PO-level header fallback by PO Number:
   PO Date
   PO Expiry/Delivery Date
   Ship-to Location
   Ship-to GST
   now show on Sale Register rows even if ERP-item line matching is not ready.

2. Every uploaded Customer PO line is visible immediately in Main Reconciliation.
   Uninvoiced/unmatched rows are appended as PO-only reconciliation rows.

3. PO-only rows show:
   PO No, Date, Expiry, Customer Item, PO Qty, PO Value, Ship-to,
   Ship-to GST, ERP Item if mapped, Sales Order, FG Stock and blocked data.

4. Missing ERP Item is not hidden:
   Remarks = "ERP Item mapping pending"

5. Once Customer SKU & Price Master is updated/reprocessed, the same PO line
   can reconcile to Sale Register by PO + ERP Item.

6. Added one-click "Refresh Reconciliation" button.


V52 — LIVE PO SOURCE OVERLAY

ROOT CAUSE ADDRESSED
Even after V51, Main Reconciliation could still display blank PO columns because
the optimized reconciliation DataFrame was built/cached separately from the live
po_lines source.

V52 GUARANTEE
On every Main Reconciliation load:
1. Existing PO Customer Item -> ERP Item mappings are refreshed from SKU Master.
2. po_lines is read directly from SQLite (not Streamlit cache).
3. Current PO fields are overlaid onto reconciliation rows.
4. Matching priority:
   PO + ERP Item -> exact line
   single-line PO -> safe single-line fallback
   otherwise -> PO header fields only (no guessed Qty/Value)
5. Any PO line not represented by Sale Register is appended and displayed.
6. Search shows a source diagnostic:
   Customer PO source matched: N lines; M have ERP mapping.

Therefore if Customer PO source contains the PO, PO Date / Expiry / Ship-to /
GST must display immediately, and line Qty/Value displays whenever exact ERP
mapping or a safe single-line match is available.


V53 — FILL PO DETAILS ONLY WHEN PO IS UPLOADED

Locked behavior:
- If PO exists in po_lines:
  fill PO Date, Expiry/Delivery Date, Ship-to Location and Ship-to GST by PO No.
- Fill PO Item, PO Qty and PO Value by PO + ERP Item.
- If a PO has only one line, use that line safely.
- If a multi-line PO cannot yet be matched to an invoice SKU, append the
  actual uploaded PO lines so PO Item / Qty / Value remain visible.
- If PO does NOT exist in po_lines:
  keep PO-specific columns blank.

Also:
- PO matching now canonicalizes punctuation/dashes/apostrophes in PO numbers.
- ERP Item is resolved live from the current Customer SKU & Price Master on
  every Main Reconciliation load.


V54 — B2B ORDER STAGING + UPLOADED PO VIEW

B2B ORDER STAGING
New navigation option:
Control Tower -> B2B Order Staging

Shows every uploaded Customer PO line in the requested format:
ID | Customer No. | Ship | Customer PO No. | Customer PO Date |
Posting Date | Item No. | Quantity | Unit Price | Status |
Sales Order No. | So Created | PO Expiry Date

The staging is downloadable as Excel.

MAIN RECONCILIATION
Added Customer PO View:
- Uploaded POs First (default)
- Uploaded POs Only
- All Rows
- PO Not Uploaded

This prevents confusion in the all-row view: Sale Register POs that were never
uploaded can legitimately have blank PO Date/Item/Qty/Value/Ship-to fields.
Uploaded POs are now surfaced first and their live po_lines data is overlaid.

Search diagnostics now use canonical PO matching, not raw SQL exact text.


V55 — SMART PO LINE MATCH + B2B DOWNLOAD ERROR FIX

MAIN RECONCILIATION
The attached reconciliation export showed the root cause clearly:
- uploaded PO line exists with Customer Item / Qty / Value,
- invoice row exists with ERP Product,
- ERP Item on the PO line can still be blank,
- therefore the invoice row showed blank PO Item / Qty / Value and a separate
  PO-only row appeared below.

V55 matches those safely inside the same PO using model/description evidence.
Examples covered:
- Walmart Article 87211 -> SA9016MCK / SA 9016 Multi Cook Kettle
- Dawntech Amazon Selena BLDC 76 -> CH60CTSELENADCAC76
- multi-line customer POs where customer-item master mapping is incomplete

After a smart match:
- PO Item / PO Qty / PO Value fill on the actual invoice row
- the matched PO-only duplicate row is not appended
- genuinely unbilled PO lines remain visible separately

DATE FIX
DD.MM.YYYY / DD-MM-YYYY / DD/MM/YYYY is now day-first.
Example: 07.08.2026 = 07-Aug-2026.

B2B ORDER STAGING
Fixed runtime Excel-download error by importing openpyxl Font/Alignment.
Missing staging mappings remain visible as "Mapping Pending".


V56 — FAST / STABLE RUNTIME FIX

FIXED NAMEERRORS
- uploaded_po_keys is restored.
- build_b2b_order_staging is restored.
- b2b_order_staging_excel_bytes is restored.
These functions were accidentally removed in V55 when the reconciliation
overlay function was replaced.

PERFORMANCE FIX
V55 repeatedly scanned Customer SKU & Price Master once for every PO line.
V56:
- loads SKU master once;
- builds normalized lookup maps once;
- resolves all PO lines in memory;
- avoids database UPDATE work on every Streamlit rerun;
- writes refreshed ERP mappings only when Refresh Reconciliation / B2B refresh
  is explicitly clicked.

B2B
The B2B staging page and Excel download now use the restored helper functions.


V57 — ERP ITEM COLUMN NEXT TO PO ITEM

Main Reconciliation column order now starts:
PO Number | PO Date | PO Expiry/DELIVERY DATE | PO Item | ERP Item |
PO Qty | PO Value | Ship to Location | ...

ERP Item source:
Customer PO Customer Item Code
  -> Customer SKU & Price Master
  -> ERP Item Code

Behavior:
- Uploaded PO + mapped Customer Item: ERP Item is filled.
- Uploaded PO but mapping missing: ERP Item remains blank and row status shows
  ERP Item mapping pending.
- Existing invoice Product/Item No remains available later in the reconciliation
  table as the ERP billing/source product field.


V58 — SHIP CODE FROM LEDGER + PIN ACROSS B2B / DASHBOARDS

Rule:
Ship to Location Code = Ship-to Location Master match by
Ledger Name + 6-digit PIN extracted from uploaded PO Ship-to address.

Applied to:
- Main Reconciliation -> Ship to Location Code
- B2B Order Staging -> Ship
- Existing enrich_ship_to_location_codes() consumers
- Factory Requirement when that view contains Ship-to Location Code

Baseline mappings supplied by user are seeded/updated at startup, including:
BI Worldwide 600077 -> BI-CHENNAI
BLINK 121006 -> FBD-BLINK
BLINK 140417 -> PATIALA
BLINK 201306 -> NOIDA -N1
BLINK 302037 -> RJ-BLINK
BLINK 382213 -> AHMD A2
BLINK 403501 -> GOA
BLINK 410501 -> NIGHOJE
BLINK 410506 -> PUNE P2
BLINK 421306 -> MUM - M10
BLINK 441501 -> NR-BLINK
BLINK 500101 -> HYD - H3
BLINK 520007 -> VIJAYAWADA
BLINK 531173 -> AP-BLINK
BLINK 562106 -> BGLR - B3
BLINK 562114 -> DODDENAHAL
BLINK 600052 -> CHEN C5


V59 — REST BLOCKED IN FACTORY STOCK REQUIREMENT

Factory Stock Requirement column order:
Branch | ERP Item Code | Item Description | Overall Pending Qty | FG Stock |
Blocked Against PO | Rest Blocked | Net Free Stock |
Stock Shortage / Factory Requirement

REST BLOCKED SOURCE
Rest Blocked is the item-wise sum of Main Reconciliation "Rest Blocked Qty"
for the selected branch / financial-year view.

Existing calculations are intentionally unchanged:
Net Free Stock = FG Stock - Blocked Against PO
Factory Requirement = max(Overall Pending Qty - Net Free Stock, 0)

This build only adds the requested Rest Blocked visibility; it does not silently
change the existing stock-shortage formula.


V60 — FAST CACHE / VECTORIZED DASHBOARDS

PERFORMANCE CHANGES
- Cached live PO source.
- Cached SKU master lookup maps.
- Cached Ship-to master lookup maps.
- Cached B2B staging dataset and Excel export.
- Cached Factory Stock Requirement by Branch + FY.
- Cached Factory branch list.
- Vectorized Ship-to code mapping instead of DataFrame row-by-row apply.
- Vectorized Factory Requirement groupby/merge instead of per-SKU Python loops.

CACHE INVALIDATION
Existing invalidate_dashboard_cache() still clears st.cache_data after uploads,
repairs and data changes. Therefore uploads remain immediately visible while
normal navigation/filter changes reuse prepared datasets.

EXPECTED EFFECT
Main navigation, Factory branch/FY switching, B2B staging and repeated dashboard
views should avoid repeated SQLite scans and repeated Python row loops.


V61 — GLOBAL DUPLICATE UPLOAD GUARD

RULE
No identical file is processed twice within the same upload source/module.

USER NOTIFICATION
When the same file is selected again:
"The same details/file have already been uploaded. Nothing was uploaded again.
Please review the existing uploaded data before uploading."

APPLIED TO NORMAL UPLOAD FLOW
- ERP Sale Register
- Sales Orders
- Shipment Not Invoiced
- Item Ledger
- Customer PO Excel/PDF
- Existing row-level duplicate controls remain active for consolidated tables.

Sales Orders no longer bypass duplicate-file protection.
Customer PO PDFs no longer silently reprocess on normal duplicate upload.
Dedicated repair/reprocess buttons remain separate administrative actions.


V62 — SHIP CODE SOURCE PRIORITY

Locked priority:
1. Uploaded Customer PO available:
   Ship-to Location Code comes from Customer PO Ship-to address PIN +
   Ship-to Location Master.

2. If PO source is missing/unresolved but billing exists:
   use Sale Register Ledger Name + Post Code / Ship-to Address to resolve
   Ship-to Location Code from the same master.

3. Existing code remains only if neither source can resolve a current code.

Applied to:
- Main Reconciliation
- B2B Order Staging
- common Ship-to enrichment used by dashboards

This means billing rows do not remain blank for Ship code merely because the
Customer PO was not uploaded, while an uploaded PO destination always has
priority over Sale Register shipping data.
