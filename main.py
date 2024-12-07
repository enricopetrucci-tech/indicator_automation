import pandas as pd
from dotenv import load_dotenv
import os
from datetime import date, timedelta
import win32com.client as win32
import pythoncom

# Load environment variables from .env file
load_dotenv()
# Get the current working directory (where the .ipynb is running)
script_dir = os.getcwd()
DAILY_REVENUE_TARGET = 1_000
YEARLY_REVENUE_TARGET = 1_650_000

DAILY_DISTINCT_PRODUCTS_TARGET = 4
YEARLY_DISTINCT_PRODUCTS_TARGET = 120

DAILY_AVG_TICKET_TARGET = 500
YEARLY_AVG_TICKET_TARGET = 500

EMAIL_FROM = os.getenv('EMAIL_FROM')
COLUMN_MAPPING = {
    'Sales Code': 'sales_code',
    'Date': 'date',
    'Store ID': 'store_id',
    'Store Name': 'store_name',
    'Product ID': 'product_id',
    'Product Name': 'product_name',
    'Quantity': 'quantity',
    'Unit Price': 'unit_price',
    'Manager': 'manager',
    'Email': 'email',
    }
def send_email(email_from, email_to, subject, email_body, file_path=None):
    '''
    
    '''
    
    # Create mail item in Outlook
    outlook = win32.Dispatch('outlook.application')
    email = outlook.CreateItem(0) 

    # Account selection
    account = outlook.Session.Accounts[email_from] 
    email._oleobj_.Invoke(*(64209, 0, 8, 0, account))

    # Email personalization
    email.To = email_to
    email.Subject = subject
    email.HTMLBody = email_body
    # email.Attachments.Add(file_path)

    email.print()
    # email.Send()

def rename_columns(df, column_mapping, reverse=False):
    '''
    Renames DataFrame columns between normalized and print-friendly names.

    Args:
        df (pd.DataFrame): The DataFrame to rename.
        column_mapping (dict): Dictionary mapping print-friendly names to normalized names.
        reverse (bool): If True, renames from normalized to print-friendly. Default is False.

    Returns:
        pd.Dataframe: DataFrame with renamed columns.
    '''
    if reverse:
        # Reverse the mapping for normalized -> aesthetic
        column_mapping = {v: k for k, v in column_mapping.items()}
    return df.rename(columns=column_mapping)
# Data sources path
data_path = os.path.join(script_dir, 'data_sources')

# Data import
emails_df = pd.read_excel(os.path.join(data_path, 'emails.xlsx'))
products_df = pd.read_excel(os.path.join(data_path, 'products.xlsx'))
stores_df = pd.read_csv(os.path.join(data_path, 'stores.csv'))
sales_df = pd.read_excel(os.path.join(data_path, 'sales.xlsx'))
# Renaming process
emails_df = rename_columns(emails_df, COLUMN_MAPPING)
products_df = rename_columns(products_df, COLUMN_MAPPING)
stores_df = rename_columns(stores_df, COLUMN_MAPPING)
sales_df = rename_columns(sales_df, COLUMN_MAPPING)

print(emails_df)
print(products_df)
print(stores_df)
print(sales_df)

### Filtered Sales Dataframes
# Getting a daily filtered sales dataframe
today = date.today()
yesterday = today - timedelta(days=1)
daily_sales_df = sales_df[sales_df['date'] == pd.Timestamp(yesterday)]

# Getting a yearly filtered sales dataframe
current_year = today.year
yearly_sales_df = sales_df[(sales_df['date'].dt.year == current_year) & (sales_df['date'].dt.date <= yesterday)]

## Indicator 1: Revenue
### Daily Revenue
# Daily revenue calculation
daily_sales_df_merged = daily_sales_df.merge(products_df, on='product_id')[['sales_code', 'date', 'store_id', 'quantity', 'unit_price']]
daily_revenue_df = daily_sales_df_merged.groupby('store_id').apply(lambda df: (df['quantity'] * df['unit_price']).sum(), include_groups=False).reset_index(name='revenue')
print(daily_revenue_df)
### Yearly Revenue
# Yearly revenue calculation
yearly_sales_df_merged = yearly_sales_df.merge(products_df, on='product_id')[['sales_code', 'date', 'store_id', 'quantity', 'unit_price']]
yearly_revenue_df = yearly_sales_df_merged.groupby('store_id').apply(lambda df: (df['quantity'] * df['unit_price']).sum(), include_groups=False).reset_index(name='revenue')
print(yearly_revenue_df)
## Indicator 2: Product Diversity
### Daily Product Diversity
# Daily
daily_distinct_products_df = daily_sales_df.groupby('store_id')['product_id'].nunique().reset_index(name='distinct_products')
print(daily_distinct_products_df)
### Yearly Product Diversity
# Yearly
yearly_distinct_products_df = yearly_sales_df.groupby('store_id')['product_id'].nunique().reset_index(name='distinct_products')
print(yearly_distinct_products_df)
## Indicator 3: Average Ticket per Sale
### Daily Average Ticket per Sale
# Daily
daily_distinct_sales_df = daily_sales_df.groupby('store_id')['sales_code'].nunique().reset_index(name='distinct_sales')
daily_kpis_df = daily_revenue_df.merge(daily_distinct_sales_df, on='store_id')
daily_kpis_df['avg_ticket'] = (daily_kpis_df['revenue'] / daily_kpis_df['distinct_sales']).round(2)
print(daily_kpis_df)
### Yearly Average Ticket per Sale
# Yearly
yearly_distinct_sales_df = yearly_sales_df.groupby('store_id')['sales_code'].nunique().reset_index(name='sales')
yearly_kpis_df = yearly_revenue_df.merge(yearly_distinct_sales_df, on='store_id')
yearly_kpis_df['avg_ticket'] = (yearly_kpis_df['revenue'] / yearly_kpis_df['sales']).round(2)
print(yearly_kpis_df)
## Grouping all KPI's together in one dataframe
daily_kpis_df = daily_kpis_df.merge(daily_distinct_products_df, on='store_id')
daily_kpis_df = daily_kpis_df[['store_id', 'revenue', 'distinct_products', 'avg_ticket']]
print(daily_kpis_df)

yearly_kpis_df = yearly_kpis_df.merge(yearly_distinct_products_df, on='store_id')
yearly_kpis_df = yearly_kpis_df[['store_id', 'revenue', 'distinct_products', 'avg_ticket']]
print(yearly_kpis_df)

# Merging both Daily and Yearly
all_kpis_df = daily_kpis_df.merge(yearly_kpis_df, on='store_id', suffixes=('_daily', '_yearly'))
print(all_kpis_df)
## Email Sending
# Initialize COM threading
pythoncom.CoInitialize()

# Get personal email from .env file
email_to = os.getenv('EMAIL_TO')

# Change dataframe emails to personal email for testing
emails_df['email'] = emails_df['email'].str.replace(
    r"email\+(.*?)@address\.com",  # Regex to extract manager name
    lambda match: f"{email_to.split('@')[0]}+{match.group(1)}@{email_to.split('@')[1]}",
    regex=True
)
emails_with_kpis_df = emails_df.merge(stores_df, on='store_id', how='left').merge(all_kpis_df, on='store_id', how='left')
print(emails_with_kpis_df)

for index, row in emails_with_kpis_df.iterrows():
    
    if row['store_id'] == 'BOARD':
        continue
    
    store_id = row['store_id']
    email_to = row['email']
    manager = row['manager']
    store_name = row['store_name']

    manager_name = manager.split(' ')[0]
    subject = f'OnePage {format(yesterday, r'%Y/%m/%d')} - {store_name}'
    daily_kpis = [
        {"name": "Revenue", "value": row['revenue_daily'], "target": DAILY_REVENUE_TARGET, 'type': 'currency'},
        {"name": "Distinct Products", "value": row['distinct_products_daily'], "target": DAILY_DISTINCT_PRODUCTS_TARGET, 'type': 'integer'},
        {"name": "Avg Ticket", "value": row['avg_ticket_daily'], "target": DAILY_AVG_TICKET_TARGET, 'type': 'currency'},
    ]    
    yearly_kpis = [
        {"name": "Revenue", "value": row['revenue_yearly'], "target": YEARLY_REVENUE_TARGET, 'type': 'currency'},
        {"name": "Distinct Products", "value": row['distinct_products_yearly'], "target": YEARLY_DISTINCT_PRODUCTS_TARGET, 'type': 'integer'},
        {"name": "Avg Ticket", "value": row['avg_ticket_yearly'], "target": YEARLY_AVG_TICKET_TARGET, 'type': 'currency'},
    ]
    email_body = f'''
    <p>Good Morning, {manager_name}</p>
    <p>Yesterday's result ({format(yesterday, r'%m/%d')}) of {store_name} was:</p>

    <p>Daily Values:</p>
    <table border="1" cellpadding="5" cellspacing="0" style="border-collapse: collapse; text-align: center; font-size: 14pt">
        <thead>
            <tr>
                <th>Indicator</th>
                <th>Day Value</th>
                <th>Day Target</th>
                <th>Day Scenario</th>
            </tr>
        </thead>
        <tbody>
    '''
    for kpi in daily_kpis:
        if kpi['type'] == 'currency':
            value = f'${kpi['value']:,.2f}'
            target = f'${kpi['target']:,.2f}'
        elif kpi['type'] == 'integer':
            value = f'{int(kpi['value'])}'
            target = f'{kpi['target']}'
        
        color = "green" if kpi['value'] >= kpi['target'] else "red"
        symbol = '◙'

        email_body += f'''
            <tr>
                <td>{kpi['name']}</td>
                <td>{value}</td>
                <td>{target}</td>
                <td style="color: {color}; font-size: 18pt;">{symbol}</td>
            </tr>
        '''

    # Close daily table and add yearly table
    email_body += '''
        </tbody>
    </table>
    <p>Yearly values:</p>
    <table border="1" cellpadding="5" cellspacing="0" style="border-collapse: collapse; text-align: center; font-size: 14pt;">
        <thead>
            <tr>
                <th>Indicator</th>
                <th>Year Value</th>
                <th>Year Target</th>
                <th>Year Scenario</th>
            </tr>
        </thead>
        <tbody>
    '''

    # Add rows for yearly KPIs
    for kpi in yearly_kpis:
        if kpi['type'] == 'currency':
            value = f'${kpi['value']:,.2f}'
            target = f'${kpi['target']:,.2f}'
        elif kpi['type'] == 'integer':
            value = f'{int(kpi['value'])}'
            target = f'{kpi['target']}'

        color = "green" if kpi['value'] >= kpi['target'] else "red"
        symbol = '◙'
        email_body += f'''
            <tr>
                <td>{kpi['name']}</td>
                <td>{value}</td>
                <td>{target}</td>
                <td style="color: {color}; font-size: 18pt;">{symbol}</td>
            </tr>
        '''

    # Close yearly table
    email_body += '''
        </tbody>
    </table>
    <p>Please find attached the spreadsheet with all the data for further details.</p>
    <p>Should you have any questions, feel free to reach out.</p>
    <br>
    <p>Best Regards,</p>
    <p>Enrico Petrucci</p>
    '''

    store_name = store_name.casefold().replace(' ', '_')
    
    yearly_sales_df = yearly_sales_df[yearly_sales_df['store_id'] == store_id]
    yearly_sales_df = yearly_sales_df.merge(products_df, on='product_id').merge(stores_df, on='store_id')
    yearly_sales_df['date'] = yearly_sales_df['date'].dt.strftime(r'%Y/%m/%d')
    yearly_sales_df = yearly_sales_df[['sales_code', 'date', 'product_name', 'store_name', 'quantity', 'unit_price']]
    if not os.path.exists(f'store_backup_files/{store_name}'):
        os.mkdir(f'store_backup_files/{store_name}')
    sales_excel = rename_columns(yearly_sales_df, COLUMN_MAPPING, reverse=True).to_excel(f'store_backup_files/{store_name}/{store_name}_sales_{yesterday.year}.xlsx', index=False)
        
    send_email(EMAIL_FROM, email_to, subject, email_body)
    break
### Board of Directors - Email
# Day
ranking_daily_df = daily_revenue_df.merge(stores_df, on='store_id')
ranking_daily_df = ranking_daily_df.sort_values(by='revenue', ascending=False)[['store_id', 'store_name', 'revenue']]

best_daily_store = ranking_daily_df.iloc[0]['store_name']
best_daily_store_revenue = ranking_daily_df.iloc[0]['revenue']
worst_daily_store = ranking_daily_df.iloc[-1]['store_name']
worst_daily_store_revenue = ranking_daily_df.iloc[-1]['revenue']
# Year
ranking_yearly_df = yearly_revenue_df.merge(stores_df, on='store_id')
ranking_yearly_df = ranking_yearly_df.sort_values(by='revenue', ascending=False)[['store_id', 'store_name', 'revenue']]

best_yearly_store = ranking_yearly_df.iloc[0]['store_name']
best_yearly_store_revenue= ranking_yearly_df.iloc[0]['revenue']
worst_yearly_store = ranking_yearly_df.iloc[-1]['store_name']
worst_yearly_store_revenue= ranking_yearly_df.iloc[-1]['revenue']
manager = emails_df.loc[emails_df['store_id'] == 'BOARD', 'manager'].iloc[0]
manager = manager.casefold().split(' ')[0]
if not os.path.exists(f'store_backup_files/{manager}'):
    os.mkdir(f'store_backup_files/{manager}')
ranking_path = f'store_backup_files/{manager}'
export_ranking_daily = ranking_daily_df.to_excel(f'{ranking_path}/{manager}_ranking_{yesterday}.xlsx', index=False)
export_ranking_yearly = ranking_yearly_df.to_excel(f'{ranking_path}/{manager}_ranking_{yesterday.year}.xlsx', index=False)
# Email Sending

email_to = emails_df.loc[emails_df['store_id'] == 'BOARD', 'email'].iloc[0]
subject = 'Daily and Yearly Store Revenue Rankings'
email_body = f'''
    <p>Dear Board,</p>

    <p>We are pleased to share the revenue performance rankings for our stores:</p>

    <p><b>Daily Performance (Date: {format(yesterday, r'%m/%d')}):</b></p>
    <ul>
        <li><b>Best Store:</b> {best_daily_store} with a revenue of <b>${best_daily_store_revenue:,.2f}</b>.</li>
        <li><b>Worst Store:</b> {worst_daily_store} with a revenue of <b>${worst_daily_store_revenue:,.2f}</b>.</li>
    </ul>

    <p><b>Yearly Performance (Year: {yesterday.year}):</b></p>
    <ul>
        <li><b>Best Store:</b> {best_yearly_store} with a revenue of <b>${best_yearly_store_revenue:,.2f}</b>.</li>
        <li><b>Worst Store:</b> {worst_yearly_store} with a revenue of <b>${worst_yearly_store_revenue:,.2f}</b>.</li>
    </ul>

    <p>The detailed daily and yearly rankings are attached for your review.</p>

    <p>Should you have any questions, feel free to reach out.</p>

    <p>Best regards,</p>
    <p>Enrico Petrucci</p>
'''

send_email(EMAIL_FROM, email_to, subject, email_body)