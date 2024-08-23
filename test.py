import pandas as pd
from datetime import date
from vnstock import *

def save_to_excel(stock_price_df, stock_fundamental_info_df, filename="one.xlsx"):
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        stock_price_df.to_excel(writer, sheet_name='Stock Price', index=False)
        stock_fundamental_info_df.to_excel(writer, sheet_name='Fundamental Info', index=False)
    print(f"Data saved to {filename}")

def get_stock_price(stock):
    try:
        stock_hist = stock_historical_data(stock, "2020-07-01", str(date.today()), "1D", "stock")
        if stock_hist.empty:
            print("No historical data available.")
            return stock_hist
        
        stock_hist["%_change"] = round(((stock_hist['close'] - stock_hist['close'].shift(1))
                                        / stock_hist['close'].shift(1)) * 100, 2)
        stock_hist['change'] = (stock_hist['close'] - stock_hist['close'].shift(1))
        stock_hist['time'] = stock_hist['time'].astype(str)
        stock_hist = stock_hist.rename(columns={'ticker': 'stock_name', 'time': 'date'})
        
        return stock_hist
    except Exception as e:
        print(f"Error in get_stock_price: {e}")
        return pd.DataFrame()  # Return empty DataFrame on error

def get_stock_fundamental_info(stock):
    try:
        fundamental_ratio = company_fundamental_ratio(stock)
        if fundamental_ratio.empty:
            print("No fundamental ratio data available.")
            return fundamental_ratio
        
        profile = company_profile(stock)
        overview = company_overview(stock)
        evaluation = stock_evaluation(stock)[["ticker", "PE", "industryPE", "PB", "industryPB"]][-1:]

        if profile.empty or overview.empty or evaluation.empty:
            print("One of the required data sets is empty.")
            return pd.DataFrame()  # Return empty DataFrame if any data is missing
        
        stock_info = pd.merge(fundamental_ratio[['ticker']], profile[['companyName', 'ticker']], on='ticker', how='inner')
        stock_info = pd.merge(stock_info, overview[['ticker', 'outstandingShare', 'exchange']], on='ticker', how='inner')
        stock_info = pd.merge(stock_info, evaluation, on='ticker', how='inner')

        stock_info = stock_info.rename(columns={'ticker': 'stock_name', 'companyName': 'company_name', 'outstandingShare': 'outstanding_share'})
        stock_info['outstanding_share'] = stock_info['outstanding_share'] * 1000000
        
        return stock_info
    except Exception as e:
        print(f"Error in get_stock_fundamental_info: {e}")
        return pd.DataFrame()  # Return empty DataFrame on error

# Gọi các hàm và lưu kết quả vào cùng một file Excel
stock_price_df = get_stock_price('HPG')
stock_fundamental_info_df = get_stock_fundamental_info('HPG')
save_to_excel(stock_price_df, stock_fundamental_info_df, filename="one.xlsx")
