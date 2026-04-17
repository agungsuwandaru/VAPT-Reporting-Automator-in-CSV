# ⚡ VAPT Reporting Automator (Google Apps Script)

## 📖 Overview
The **VAPT Reporting Automator** is a suite of Google Apps Script tools designed to enforce standardization across Vulnerability Assessment and Penetration Testing (VAPT) reports. In a hybrid environment where VAPT is conducted by both Internal Security Teams and third-party MSSPs (Managed Security Service Providers), maintaining a uniform reporting format is critical. 

This script acts as a centralized controller. It reads "Master Standards" (Headers, Notes, Risk Matrices, and PoC contents) from a central spreadsheet and distributes them automatically to multiple target report spreadsheets via a custom UI menu.

## ✨ Key Features
* **Custom UI Menu:** Automatically creates a `⚡ VAPT TOOLS` menu upon opening the spreadsheet for easy execution.
* **Smart Merge Handling:** Intelligently injects Notes into merged cells without losing them in the background.
* **Auto Text Normalization:** Bypasses typos, extra spaces, and case-sensitivity (`"  MSSP\n "` matches `"MSSP"`).
* **Safe Execution:** Ignores dashes (`"-"`) in standard definitions to prevent script from accidentally overwriting empty data cells formatted with dashes.
* **Bulk Processing:** Can update multiple target spreadsheets in a single run based on a control list.

---

## 🏗️ System Logic & Workflow

### 1. The Control Panel (`Report List` Tab)
The script relies on a master control tab (usually GID 0) named `Report List`. The script scans this list to determine which target spreadsheets to process. It looks for specific columns:
* `Run` (Checkbox): Must be set to `TRUE` for the script to process the row.
* `Update` (String): Must perfectly match the intended function (e.g., `"Kolom - Dashboard BGN"`).
* `Report Link` (URL): The full Google Sheets URL of the target report. The script automatically extracts the unique Document ID from this URL.
* `Tab` (String): The specific sheet name inside the target document to update.

### 2. The Master Standards
The script pulls the "source of truth" from designated standard tabs (e.g., `Standar Header`, `Standar Content`). It dynamically searches for columns containing the words "Nama/Header" (Key) and "Note/Penjelasan" (Value) to build a dictionary. 

### 3. Execution & Injection
Once triggered, the script:
1. Validates the `Report List` conditions.
2. Opens the target spreadsheets in the background.
3. Scans the top rows (up to 10 rows to support multi-tier/merged headers).
4. Matches the target cell text with the dictionary.
5. Injects the Note or Content precisely where it belongs.

---

## 🛠️ Functions Breakdown

### 🎨 UI Menu Function
* **`onOpen()`**
    * **Purpose:** Triggers automatically when the Google Sheet is opened. It builds the custom `⚡ VAPT TOOLS` dropdown menu in the Google Sheets toolbar, linking each menu item to its respective function.

### 📝 Header & Notes Updaters (`Kolom`)
These functions focus on injecting Google Sheets **Notes** (the small black triangle in the corner of a cell that shows a pop-up explanation) into the headers of target reports.
* **`jalankanUpdateKolomDetailFindingVAPT()`**: Standardizes the column explanations for the *Detail Finding* tab (e.g., explaining what "Status", "Risk Level", or "CVSS" means).
* **`jalankanUpdateKolomPoCVAPT()`**: Standardizes the column explanations for the *Proof of Concept* tab.
* **`jalankanUpdateKolomDashboardBGN()`**: Standardizes the matrix explanations in the Dashboard tab. Features deep-scanning (up to 10 rows) to handle complex, heavily merged dashboard table headers.

### 🗃️ Content & Data Updaters (`Content`)
These functions focus on syncing actual cell **Values/Content** (text, dropdown lists, matrices) rather than just Notes.
* **`jalankanUpdateContentDetailFindingVAPT()`**: Synchronizes standard vulnerability descriptions, impacts, or remediation recommendations based on the finding's category.
* **`jalankanUpdateContentPoCVAPT()`**: Injects standard reproduction steps (PoC frameworks) for specific vulnerability types without overwriting the pentester's custom screenshots/evidence.
* **`jalankanUpdateContentHelperVAPT()`**: Synchronizes the hidden/helper tabs in target reports. This ensures all reports share the exact same Risk Matrices, Dropdown Data Validations, and SLA definitions.

---

## 🚀 How to Use

### Setup Installation
1. Open your Master Google Spreadsheet.
2. Go to **Extensions > Apps Script**.
3. Create two files:
   * `menu.gs` (Paste the `onOpen` function here).
   * `main.gs` (Paste all the `jalankanUpdate...` functions here).
4. Save the project and refresh your Google Spreadsheet. 

### Running the Tools
1. Wait for the `⚡ VAPT TOOLS` menu to appear in the top toolbar (next to Help).
2. Go to your `Report List` tab.
3. Check (`TRUE`) the box in the **Run** column for the documents you want to update.
4. Ensure the **Update** column matches the operation you want to perform (e.g., `"Kolom - Dashboard BGN"`).
5. Click **⚡ VAPT TOOLS** in the menu and select the corresponding action.
6. The script will run in the background and show a popup alert with the success count once finished.

*Note: On the very first run, Google will ask for Authorization to allow the script to read/edit your spreadsheets. Click "Review Permissions" -> Choose your account -> "Advanced" -> "Go to script (unsafe)" -> "Allow".*
