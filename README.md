# Store Performance Automation Project
This project automates daily performance reporting for a fictional retail clothing chain with 25 stores, providing concise and actionable insights for store managers and executives. By streamlining the reporting process, it enhances decision-making and ensures data-driven management. The automation reduces manual reporting time, improves operational efficiency, and provides precise performance tracking against predefined goals.

---

## Features

- **Daily KPI Reporting**: Generate daily reports comparing store performance against predefined annual and daily goals.
- **Executive Summaries**: Provide rankings of stores by revenue, highlighting the best and worst performers.
- **Email Automation**: Automatically send One Page reports to store managers and detailed revenue summaries to the board.
- **Historical Tracking**: Save timestamped backup files for YTD sales data, enabling trend analysis and auditing.

---

## Key Performance Indicators (KPIs)

### Tracked Indicators
1. **Revenue**:
   - Annual target: $1,650,000
   - Daily target: $1,000
2. **Product Diversity**:
   - Annual target: 120 products
   - Daily target: 4 products
3. **Average Ticket per Sale**:
   - Annual and daily target: $500

---

## Workflow

1. **Daily Automation**:
   - Calculate daily and YTD KPIs for each store.
   - Generate HTML-based One Page summaries for email distribution.
   - Save YTD sales data as timestamped backup files.
   
2. **Email Distribution**:
   - Store Managers: Receive daily reports via email with KPI tables and attached sales data.
   - Executive Team: Receive board reports summarizing revenue performance rankings.

3. **Historical Backup**:
   - Maintain daily backup files in `store_backup_files` directory for each store.

---

## Files Used

1. **emails.xlsx**:
   - Contains store IDs, manager names, and email addresses.

2. **sales.xlsx**:
   - Consolidated sales data with revenue details.

3. **stores.csv**:
   - Maps store IDs to store names.

4. **products.xlsx**:
   - Product details, including IDs, names, and unit prices.

---

## Output

1. **Manager Reports**:
   - Daily email with:
     - KPI tables for daily and YTD performance.
     - Attached sales file (`store_backup_files/{store_name}_YYYY_MM_DD_sales.xlsx`).

2. **Board of Directors Reports**:
   - Email with:
     - Rankings for daily and YTD revenue.
     - Attachments (`daily_ranking_YYYY_MM_DD.xlsx`, `ytd_ranking_YYYY_MM_DD.xlsx`).

---

## Getting Started

### Dependencies
1. **Python 3.13**
2. **Required Libraries**:
   - `pandas`: For data manipulation and analysis.
   - `python-dotenv`: To load environment variables securely.
   - `pywin32`: For automating Outlook email sending.
   - `psutil`: To check if Outlook is running and manage system processes.
   - `openpyxl`: For reading and writing Excel files in the `.xlsx` format.

Install the required libraries using pip:
```bash
pip install pandas python-dotenv pywin32 psutil openpyxl
```
3. **Additional Requirements**:
- **Outlook**: You must have the classic desktop version of Microsoft Outlook installed and set up, as the "New Outlook" does not fully support automation.
- **.env File**: Create a `.env` file in the root directory of your project. This file securely stores sensitive information like email addresses.

**Variables to include**:
- `EMAIL_FROM`: The sender's email address. This must match an account configured in Outlook.
- `EMAIL_TO`: A recipient email address to receive the emails sent by the script.

**Example .env File**:
```env
EMAIL_FROM=your_email@example.com
EMAIL_TO=your_test_email@example.com
```

---
### Running the Script
1. Place the required files (`emails.xlsx`, `sales.xlsx`, etc.) in a `data_sources` folder within the same directory as the main script file.
2. Ensure the `.env` file is in the same directory as the main script file.
3. If you want to **preview** the emails instead of sending them, modify the call to `send_email` in the code to include:
```python
send_email(email_from, email_to, subject, email_body, file_paths=[attachments_path], preview=True) 
```
4. To limit the script to handle only one email during testing, uncomment the `break` line in the loop:
```python
for row in emails_with_kpis_df.itertuples():
   send_email(EMAIL_FROM, email_to, subject, email_body, file_paths=[ytd_file_path])
   break # Uncomment this line for testing to send/preview only one email
```
5. Run the script:
```bash
python main.py
```
---
### Limitations
- Requires the classic desktop version of Microsoft Outlook. The "New Outlook" is not supported.
- Assumes input files (`emails.xlsx`, `sales.xlsx`, etc.) are correctly formatted.
- Designed for a fictional clothing store chain and may require customization for other use cases.


### License
This project is licensed under the MIT License. See the [LICENSE](./LICENSE) for details.

### Contact
If you have any questions or suggestions, feel free to reach out to **Enrico Petrucci** at [enrico.petrucci@outlook.com](mailto:enrico.petrucci@outlook.com).