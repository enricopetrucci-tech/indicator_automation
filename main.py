from datetime import date, timedelta
from pathlib import Path
import os
import time

from dotenv import load_dotenv
import win32com.client as win32
import pandas as pd

# Get the current parent path
SCRIPT_DIR = Path(__file__).parent

# Load environment variables from .env file
load_dotenv(dotenv_path=SCRIPT_DIR / '.env')

# GLOBAL VARIABLES

# Financial Targets
DAILY_TARGETS = {
    'revenue': 1_000,
    'distinct_products': 4,
    'avg_ticket': 500
}

YEARLY_TARGETS = {
    'revenue': 1_650_000,
    'distinct_products': 120,
    'avg_ticket': 500
}
        
EMAIL_FROM = os.getenv('EMAIL_FROM')
if not EMAIL_FROM:
    raise ValueError('EMAIL_FROM environment variable is not set.')

## Functions
def construct_kpi_list(row, targets, period):
    '''
    Construct a list of KPI dictionaries for a given period (daily or YTD).

    Args:
        row (pd.Series): A row of the DataFrame containing KPI values.
        targets (dict): A dictionary of target values for the KPIs (e.g., revenue, distinct_products, avg_ticket).
        period (str): The period for the KPIs ('daily', 'yearly').

    Returns:
        list: A list of dictionaries, each representing a KPI with its name, value, target and type.
    '''
    return [
        {'name': 'Revenue', 'value': getattr(row, f'revenue_{period}'), 'target': targets['revenue'], 'type': 'currency'},
        {'name': 'Distinct Products', 'value': getattr(row, f'distinct_products_{period}'), 'target': targets['distinct_products'], 'type': 'integer'},
        {'name': 'Avg Ticket', 'value': getattr(row, f'avg_ticket_{period}'), 'target': targets['avg_ticket'], 'type': 'currency'}
    ]

def format_kpi_table(kpis, title, is_yearly=False):
    '''
    Formats a KPI table as an HTML string.

    Args:
        kpis (list): A list of dictionaries representing KPIs.
        title (str): The title of the table (e.g., 'Daily Values').
        is_yearly (bool): Whether the KPIs are yearly. Default is False.
    
    Returns:
        str: An HTML string for the table.
    '''
    table = f'''
    <div style='text-align: center; margin-bottom: 10px; font-size: 16pt'>
        <b>{title}</b>
    '''
    table += '''
    <table border='1' cellpadding='10' cellspacing='0' style='border-collapse: collapse; text-align: center; font-size: 14pt; margin: 20px 0; background-color: #f9f9f9; width: 100%; font-family: Calibri, sans-serif;'>
        <thead>
            <tr>
                <th>Indicator</th>
                <th>{}</th>
                <th>Target</th>
                <th>Scenario</th>
            </tr>
        </thead>
        <tbody>
    '''.format('Year Value' if is_yearly else 'Day Value')

    for kpi in kpis:
        value = f'${kpi['value']:,.2f}' if kpi['type'] == 'currency' else f'{int(kpi['value'])}'
        target = f'${kpi['target']:,.2f}' if kpi['type'] == 'currency' else f'{int(kpi['target'])}'
        color = "green" if kpi['value'] >= kpi['target'] else "red"
        symbol = '◙'
        table += f'''
            <tr>
                <td>{kpi['name']}</td>
                <td>{value}</td>
                <td>{target}</td>
                <td style="color: {color}; font-size: 18pt;">{symbol}</td>
            </tr>
        '''
    table += '''
        </tbody>
    </table>
    '''
    return table

def get_ranking_info(df, column_name='revenue'):
    '''
    Sorts the DataFrame by the specified column and extracts the best and worst store rankings.

    Args:
        df (pd.DataFrame): DataFrame containing store performance data.
        column_name (str): Column to sort by for rankings. Default is 'revenue'.
    
    Returns:
        tuple: (sorted_df, best_store_name, best_store_value, worst_store_name, worst_store_value)
    '''
    sorted_df = df.sort_values(by=column_name, ascending=False)[['store_id', 'store_name', column_name]]
    best_store_name = sorted_df.iloc[0]['store_name']
    best_store_value = sorted_df.iloc[0][column_name]
    worst_store_name = sorted_df.iloc[-1]['store_name']
    worst_store_value = sorted_df.iloc[-1][column_name]
    return sorted_df, best_store_name, best_store_value, worst_store_name, worst_store_value

def send_email(email_from, email_to, subject, email_body, file_paths=None, preview=False):
    '''
    Sends an email using Outlook automation.

    Args:
        email_from (str): Sender email address.
        email_to (str): Recipient email address.
        subject (str): Subject of the email.
        email_body (str): HTML content of the email body.
        file_paths (list, optional): List of file paths to attach. Defaults to None.
        preview (bool): If True, displays the email instead of sending it. Default is False.
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

    # Add attachments if provided
    if file_paths:
        for file_path in file_paths:
            path = Path(file_path)
            if path.exists() and path.is_file():
                email.Attachments.Add(str(path.resolve())) # Convert Path object to string
            else:
                print(f'Warning: File not found - {path}')

    # Preview or send email
    if preview:
        email.Display() # Open email for preview
    else:
        email.Send() # Send the email
        print(f'Email sent to {email_to} with subject: {subject}')

def standardize_column_names(df):
    '''
    Standardizes column names by converting them to lowercase and replacing spaces with underscores.

    Args:
        df (pd.DataFrame): The DataFrame whose column names are to be standardized.
    
    Returns:
        pd.DataFrame: DataFrame with standardized column names.
    '''
    df.columns = df.columns.str.lower().str.replace(' ', '_')
    return df

def restore_column_names(df):
    '''
    Restores column names by capitalizing and replacing underscores with spaces.

    Args:
        df (pd.DataFrame): The DataFrame whose column names are to be restored.
    
    Returns:
        pd.DataFrame: DataFrame with restored column names.
    '''
    df.columns = df.columns.str.title().str.replace('_', ' ')
    return df

def main():
    # Data sources path
    data_sources_path = SCRIPT_DIR / 'data_sources'

    # Data import and renaming
    emails_df = standardize_column_names(pd.read_excel(data_sources_path / 'emails.xlsx'))
    products_df = standardize_column_names(pd.read_excel(data_sources_path / 'products.xlsx'))
    stores_df = standardize_column_names(pd.read_csv(data_sources_path / 'stores.csv'))
    sales_df = standardize_column_names(pd.read_excel(data_sources_path / 'sales.xlsx'))

    # Ensure 'date' column is in datetime format
    sales_df['date'] = pd.to_datetime(sales_df['date'])

    # Get today's and yesterday's dates
    today = date.today()
    yesterday_date = today - timedelta(days=1)

    # Filter daily sales
    daily_sales_df = sales_df[sales_df['date'] == pd.Timestamp(yesterday_date)]

    # Filter yearly sales
    current_year = yesterday_date.year
    yearly_sales_df = sales_df[(sales_df['date'].dt.year == current_year) & (sales_df['date'] <= pd.to_datetime(yesterday_date))]

    ## Indicator 1: Revenue

    # Daily
    daily_sales_df_merged = daily_sales_df.merge(products_df, on='product_id')[['sales_code', 'date', 'store_id', 'quantity', 'unit_price']]
    daily_sales_df_merged['revenue'] = daily_sales_df_merged['quantity'] * daily_sales_df_merged['unit_price']
    daily_revenue_df = daily_sales_df_merged.groupby('store_id')['revenue'].sum().reset_index(name='revenue')

    # Yearly
    yearly_sales_df_merged = yearly_sales_df.merge(products_df, on='product_id')[['sales_code', 'date', 'store_id', 'quantity', 'unit_price']]
    yearly_sales_df_merged['revenue'] = yearly_sales_df_merged['quantity'] * yearly_sales_df_merged['unit_price']
    yearly_revenue_df = yearly_sales_df_merged.groupby('store_id')['revenue'].sum().reset_index(name='revenue')

    ## Indicator 2: Product Diversity

    # Daily
    daily_distinct_products_df = daily_sales_df.groupby('store_id')['product_id'].nunique().reset_index(name='distinct_products')

    # Yearly
    yearly_distinct_products_df = yearly_sales_df.groupby('store_id')['product_id'].nunique().reset_index(name='distinct_products')

    ## Indicator 3: Average Ticket per Sale

    # Daily
    daily_distinct_sales_df = daily_sales_df.groupby('store_id')['sales_code'].nunique().reset_index(name='distinct_sales')
    daily_avg_ticket_df = daily_revenue_df.merge(daily_distinct_sales_df, on='store_id')
    daily_avg_ticket_df['avg_ticket'] = (daily_avg_ticket_df['revenue'] / daily_avg_ticket_df['distinct_sales'])

    # Yearly
    yearly_distinct_sales_df = yearly_sales_df.groupby('store_id')['sales_code'].nunique().reset_index(name='distinct_sales')
    yearly_avg_ticket_df = yearly_revenue_df.merge(yearly_distinct_sales_df, on='store_id')
    yearly_avg_ticket_df['avg_ticket'] = (yearly_avg_ticket_df['revenue'] / yearly_avg_ticket_df['distinct_sales'])

    ## Combine all KPIs into summary DataFrames

    # Daily KPIs
    daily_kpis_df = daily_revenue_df.merge(daily_distinct_products_df, on='store_id')
    daily_kpis_df = daily_kpis_df.merge(daily_avg_ticket_df[['store_id', 'avg_ticket']], on='store_id')

    # Yearly KPIs
    yearly_kpis_df = yearly_revenue_df.merge(yearly_distinct_products_df, on='store_id')
    yearly_kpis_df = yearly_kpis_df.merge(yearly_avg_ticket_df[['store_id', 'avg_ticket']], on='store_id')

    # Merging both Daily and Yearly
    all_kpis_df = daily_kpis_df.merge(yearly_kpis_df, on='store_id', suffixes=('_daily', '_yearly'))

    ## Email Sending

    # Get personal email from .env file
    email_to = os.getenv('EMAIL_TO')

    # Change dataframe emails to personal email for testing
    emails_df['email'] = emails_df['email'].str.replace(
        r"email\+(.*?)@address\.com",  # Regex to extract manager name
        lambda match: f"{email_to.split('@')[0]}+{match.group(1)}@{email_to.split('@')[1]}",
        regex=True
    )
    emails_with_kpis_df = emails_df.merge(stores_df, on='store_id', how='left').merge(all_kpis_df, on='store_id', how='left')

    for row in emails_with_kpis_df.itertuples():
        if row.store_id == 'BOARD':
            continue
        
        manager_name = row.manager.split(' ')[0]
        store_name = row.store_name
        store_id = row.store_id
        email_to = row.email

        daily_kpis = construct_kpi_list(row, DAILY_TARGETS, 'daily')
        yearly_kpis = construct_kpi_list(row, YEARLY_TARGETS, 'yearly')

        daily_table = format_kpi_table(daily_kpis, 'Daily Values')
        yearly_table = format_kpi_table(yearly_kpis, 'Yearly Values', is_yearly=True)

        subject = f'OnePage {yesterday_date.strftime(r'%Y/%m/%d')} - {store_name}'
        email_body = f'''
        <style>
            body, table, p, td {{
                font-family: Calibri, sans-serif;
            }}
        </style>  
        <p>Good Morning, {manager_name}</p>
        <p>Yesterday's result ({yesterday_date.strftime(r'%m/%d')}) of {store_name} was:</p>
        <table style='width: 100%; border-collapse: collapse;'>
            <tr>
                <td style='width: 50%; vertical-align: top; padding: 10px;'>
                    {daily_table}
                </td>
                <td style='width: 50%; vertical-align: top; padding: 10px;'>
                    {yearly_table}
                </td>
            </tr>
        </table>
        <p>Please find attached the spreadsheet with all the data for further details.</p>
        <p>Should you have any questions, feel free to reach out.</p>
        <br>
        <p>Best Regards,</p>
        <p>Enrico Petrucci</p>
        '''

        # Generate and Save the Year-to-Date Excel File
        store_name_safe = store_name.casefold().replace(' ', '_')
        backup_dir = SCRIPT_DIR / 'store_backup_files' / store_name_safe
        backup_dir.mkdir(parents=True, exist_ok=True)
        
        # Filter sales data for the store and enrich with additional details 
        yearly_sales_df_filtered = yearly_sales_df[yearly_sales_df['store_id'] == store_id]
        yearly_sales_df_filtered = yearly_sales_df_filtered.merge(products_df, on='product_id').merge(stores_df, on='store_id')
        yearly_sales_df_filtered['date'] = yearly_sales_df_filtered['date'].dt.strftime(r'%Y/%m/%d')

        # Select and rename columns for clarity
        yearly_sales_df_filtered = yearly_sales_df_filtered[['sales_code', 'date', 'product_name', 'store_name', 'quantity', 'unit_price']]
        yearly_sales_df_filtered = restore_column_names(yearly_sales_df_filtered)
        
        # Save the backup .xlsx file
        ytd_file_path = backup_dir / f'{store_name_safe}_sales.xlsx'
        yearly_sales_df_filtered.to_excel(ytd_file_path, index=False) # Fixed filename for overwriting
        
        # Send email with Attachment
        send_email(EMAIL_FROM, email_to, subject, email_body, file_paths=[ytd_file_path])
        # break
        time.sleep(2)

    # Board of Directors - Email

    ranking_daily_df, best_daily_store, best_daily_store_revenue, worst_daily_store, worst_daily_store_revenue = get_ranking_info(
        daily_revenue_df.merge(stores_df, on='store_id')
    )
    ranking_yearly_df, best_yearly_store, best_yearly_store_revenue, worst_yearly_store, worst_yearly_store_revenue = get_ranking_info(
        yearly_revenue_df.merge(stores_df, on='store_id')
    )

    # Ensure directory for rankings
    ranking_dir = SCRIPT_DIR / 'store_backup_files' / 'board_of_directors'
    ranking_dir.mkdir(parents=True, exist_ok=True)

    # Generate and save ranking files
    daily_export_path = ranking_dir / f'daily_ranking_{yesterday_date.strftime(r'%Y_%m_%d')}.xlsx'
    ranking_daily_df = restore_column_names(ranking_daily_df)
    ranking_daily_df.to_excel(daily_export_path, index=False)

    yearly_export_path = ranking_dir / f'yearly_ranking_{yesterday_date.year}.xlsx'
    ranking_yearly_df = restore_column_names(ranking_yearly_df)
    ranking_yearly_df.to_excel(yearly_export_path, index=False)

    # Email Sending
    email_to = emails_df.loc[emails_df['store_id'] == 'BOARD', 'email'].iloc[0]
    subject = f'Daily and YTD Store Revenue Rankings - {yesterday_date.strftime(r'%Y/%m/%d')}'
    email_body = f'''
        <style>
            body, table, p, td, li {{
                font-family: Calibri, sans-serif;
            }}
        </style>
        <p>Dear Board,</p>

        <p>We are pleased to share the revenue performance rankings for our stores:</p>

        <p><b>Daily Performance (Date: {yesterday_date.strftime(r'%m/%d')}):</b></p>
        <ul>
            <li><span style='color: green;'>Best Store:</span> <b>{best_daily_store}</b> with a revenue of <b>${best_daily_store_revenue:,.2f}</b>.</li>
            <li><span style='color: red;'>Worst Store:</span> <b>{worst_daily_store}</b> with a revenue of <b>${worst_daily_store_revenue:,.2f}</b>.</li>
        </ul>

        <p><b>YTD Performance (Year: {yesterday_date.year}):</b></p>
        <ul>
            <li><span style='color: green;'>Best Store:</span> <b>{best_yearly_store}</b> with a revenue of <b>${best_yearly_store_revenue:,.2f}</b>.</li>
            <li><span style='color: red;'>Worst Store:</span> <b>{worst_yearly_store}</b> with a revenue of <b>${worst_yearly_store_revenue:,.2f}</b>.</li>
        </ul>

        <p>The detailed daily and yearly rankings are attached for your review.</p>

        <p>Should you have any questions, feel free to reach out.</p>

        <p>Best regards,</p>
        <p>Enrico Petrucci</p>
    '''
    send_email(EMAIL_FROM, email_to, subject, email_body, file_paths=[daily_export_path, yearly_export_path])

if __name__ == '__main__':
    main()