import pandas as pd
import yfinance as yf
import numpy as np
import datetime
from dateutil.relativedelta import relativedelta
import warnings
import openpyxl

# Suppress warnings
warnings.filterwarnings('ignore')

def fetch_colombian_indices():
    """
    Fetches daily price close percentage changes for Colombian financial indices
    for the last 5 years and exports to Excel with proper percentage formatting.
    
    Indices: Colcap, Coltes, Colir, USD COP TRM, MSCI Colcap
    """
    print("Starting Colombian indices data fetch process...")
    
    # Set up date range (last 5 years)
    end_date = datetime.datetime.now()
    start_date = end_date - relativedelta(years=5)
    
    print(f"Time period: {start_date.strftime('%Y-%m-%d')} to {end_date.strftime('%Y-%m-%d')}")
    
    # Colombian market tickers
    tickers_dict = {
        'Colcap': ['ICOLCAP.CL', 'BVC.CL', 'ICOL'],
        'Coltes': ['GTCOL10Y.BO', 'GCL.CL', 'TLT', 'COLT.CL'],  # Added more bond proxies
        'Colir': ['CLH.CL', 'HCOLSEL.CL'],
        'USD COP TRM': ['COP=X', 'USDCOP=X'],
        'MSCI Colcap': ['MSCI.CL', 'EEM']
    }
    
    # Create empty result dataframe
    result_df = pd.DataFrame()
    
    # Fetch and process each index
    for index_name, ticker_list in tickers_dict.items():
        print(f"\nAttempting to fetch {index_name} data...")
        
        for ticker in ticker_list:
            try:
                print(f"Trying ticker: {ticker}")
                # Download data
                data = yf.download(ticker, start=start_date, end=end_date, progress=False)
                
                if data.empty or 'Close' not in data.columns:
                    print(f"No data found for ticker {ticker}")
                    continue
                
                # Process the data
                if not data.empty:
                    print(f"Successfully fetched data for {ticker} with {len(data)} rows")
                    
                    # Calculate daily percentage change
                    data_pct = data['Close'].pct_change() * 100
                    
                    # Remove timezone information from index if present
                    if data_pct.index.tz is not None:
                        data_pct.index = data_pct.index.tz_localize(None)
                    
                    # Add to result dataframe
                    if result_df.empty:
                        result_df = pd.DataFrame(index=data_pct.index)
                        result_df.index.name = 'Fecha'
                    
                    # Add the data to the result dataframe
                    result_df[index_name] = data_pct
                    
                    print(f"Added {index_name} data to the result dataframe")
                    print(f"Sample: {data_pct.head(3)}")
                    
                    # Break the loop for this index since we found valid data
                    break
                
            except Exception as e:
                print(f"Error with ticker {ticker}: {str(e)}")
    
    # If we still couldn't find Coltes data, try synthetic approach
    if 'Coltes' not in result_df.columns or result_df['Coltes'].isna().all():
        print("\nUsing TLT (US Treasury Bond ETF) as a proxy for Coltes...")
        try:
            coltes_proxy = yf.download('TLT', start=start_date, end=end_date, progress=False)
            if not coltes_proxy.empty:
                coltes_data = coltes_proxy['Close'].pct_change() * 100
                if coltes_data.index.tz is not None:
                    coltes_data.index = coltes_data.index.tz_localize(None)
                result_df['Coltes'] = coltes_data
                print("Added Coltes proxy data (using TLT as substitute)")
        except Exception as e:
            print(f"Error creating Coltes proxy: {e}")
    
    # If no tickers worked, create an empty dataframe with the desired date range
    if result_df.empty:
        print("\nNo data was successfully fetched. Creating empty dataframe with date range.")
        date_range = pd.date_range(start=start_date, end=end_date, freq='B')
        result_df = pd.DataFrame(index=date_range)
        result_df.index.name = 'Fecha'
    
    # Ensure all indices are in the dataframe
    for index_name in tickers_dict.keys():
        if index_name not in result_df.columns:
            result_df[index_name] = np.nan
    
    # Print data coverage summary
    print("\nData coverage summary:")
    for column in result_df.columns:
        missing_pct = result_df[column].isna().mean() * 100
        available_pct = 100 - missing_pct
        print(f"{column}: {available_pct:.2f}% available ({missing_pct:.2f}% missing)")
    
    # Save to Excel with percentage formatting
    output_file = 'colombian_indices_data.xlsx'
    
    # First save with Pandas
    result_df.to_excel(output_file)
    print(f"\nData exported to {output_file}")
    
    # Then apply percentage formatting
    try:
        # Load the workbook
        wb = openpyxl.load_workbook(output_file)
        ws = wb.active
        
        # Get the number of rows and columns
        max_row = ws.max_row
        max_col = ws.max_column
        
        # Apply percentage formatting to all data cells
        for col in range(2, max_col + 1):  # Skip the date column
            for row in range(2, max_row + 1):  # Skip the header row
                cell = ws.cell(row=row, column=col)
                if cell.value is not None:
                    # Direct number format without using NumberFormat class
                    cell.number_format = '0.00"%"'
        
        # Save the formatted workbook
        wb.save(output_file)
        print("Excel file formatted with percentage symbols")
    except Exception as e:
        print(f"Error applying Excel formatting: {e}")
    
    # If still missing Coltes data, provide specific instructions
    if 'Coltes' not in result_df.columns or result_df['Coltes'].isna().all():
        print("\n" + "="*50)
        print("INSTRUCTIONS FOR ACQUIRING COLTES DATA")
        print("="*50)
        print("\nTo get Coltes (Colombian Government Bond) data:")
        print("\n1. Official Banco de la República source:")
        print("   - Visit: https://www.banrep.gov.co/es/estadisticas/tes")
        print("   - Download historical TES data")
        
        # Provide code to manually add Coltes data
        print("\nAfter downloading Coltes data, use this code to add it to your Excel file:")
        print("""
# Python code to add Coltes data to the existing Excel file
import pandas as pd

# Load the excel file we created
df = pd.read_excel('colombian_indices_data.xlsx', index_col='Fecha')

# Import the Coltes data (adjust file path and format as needed)
coltes_data = pd.read_csv('coltes_data.csv')
coltes_data['Fecha'] = pd.to_datetime(coltes_data['Fecha'], format='%Y-%m-%d')
coltes_data = coltes_data.set_index('Fecha')

# Calculate percentage changes if needed
if 'Valor' in coltes_data.columns:
    coltes_data['Coltes'] = coltes_data['Valor'].pct_change() * 100

# Update our main dataframe
df['Coltes'] = coltes_data['Coltes']

# Save updated dataframe
df.to_excel('colombian_indices_data_complete.xlsx')

# Apply percentage formatting
import openpyxl

wb = openpyxl.load_workbook('colombian_indices_data_complete.xlsx')
ws = wb.active
for col in range(2, ws.max_column + 1):
    for row in range(2, ws.max_row + 1):
        cell = ws.cell(row=row, column=col)
        if cell.value is not None:
            cell.number_format = '0.00"%"'
wb.save('colombian_indices_data_complete.xlsx')
        """)
    
    return result_df

# Execute the function
if __name__ == "__main__":
    data = fetch_colombian_indices()
    print("\nScript execution complete.")