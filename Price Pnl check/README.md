
# System Requirements
Windows OS (Outlook COM integration)

Python 3.8+

Microsoft Outlook installed with access to the relevant mailbox/folders

Internal network access for API endpoints using Kerberos authentication

# Daily PnL Review Automation Tool

This repository provides a robust Python automation script for performing **daily profit and loss (PnL) validation** by fetching data from emails, parsing Excel and CSV files, calling internal pricing and security APIs, and generating a consolidated report with comments. It **reads incoming emails**, **processes attachments**, and **sends a response email** with the final output — streamlining daily investment operations workflows.

---

## Key Features

-  **Reads Outlook Emails**: Automatically fetches emails from a specific Outlook folder (e.g., "PnL Review") and downloads attachments.
-  **Parses Attachments**: Supports parsing of:
  - Excel workbooks (.xlsx)
  - CSV reports
  - (Optional) PDFs (OCR commented out)
-  **Calls Internal APIs**:
  - Security metadata (for canonical mapping)
  - Pricing details with date-based and SPN filters
  - Pricing source data
-  **Performs PnL Validation**:
  - Flags high-impact rows based on business rules (bps threshold, type, subtype)
  - Validates pricing hierarchy (e.g., Bloomberg, WM/Reuters)
  - Adds contextual comments (e.g., “Marked as per agreed hierarchy”)
-  **Generates Output Files**:
  - Annotated Excel file with comments and highlights
  - Pricing source mapping workbook
-  **Sends Email Replies**:
  - Composes and sends an Outlook reply with processed files attached
  - Maintains the subject/thread for easy tracking


# Output Files
-Two files are created under the output/ folder:

-PnL Check [date].xlsx: Final report with PnL flags and pricing comments.

data_pricing source.xlsx: Combined data from pricing API queries.
