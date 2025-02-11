import requests
import pandas as pd
from datetime import datetime
import time
import os
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

class CryptoDataFetcher:
    def __init__(self):
        self.api_key = os.getenv('COINGECKO_API_KEY')
        self.base_url = 'https://api.coingecko.com/api/v3'
        self.excel_file = 'crypto_data.xlsx'
        
    def fetch_top_50_crypto(self):
        """Fetch top 50 cryptocurrencies data from CoinGecko API"""
        endpoint = f"{self.base_url}/coins/markets"
        params = {
            'vs_currency': 'usd',
            'order': 'market_cap_desc',
            'per_page': 50,
            'page': 1,
            'sparkline': False
        }
        
        try:
            response = requests.get(endpoint, params=params)
            response.raise_for_status()
            return response.json()
        except requests.RequestException as e:
            print(f"Error fetching data: {e}")
            return None

    def process_crypto_data(self, data):
        """Process the raw cryptocurrency data"""
        if not data:
            return None
        
        processed_data = []
        for coin in data:
            processed_data.append({
                'Name': coin['name'],
                'Symbol': coin['symbol'].upper(),
                'Current Price (USD)': coin['current_price'],
                'Market Cap': coin['market_cap'],
                '24h Trading Volume': coin['total_volume'],
                'Price Change 24h (%)': coin['price_change_percentage_24h']
            })
        
        return pd.DataFrame(processed_data)

    def analyze_data(self, df):
        """Perform analysis on the cryptocurrency data"""
        analysis = {
            'Top 5 by Market Cap': df.nlargest(5, 'Market Cap')[['Name', 'Market Cap']],
            'Average Price': df['Current Price (USD)'].mean(),
            'Highest 24h Change': df.nlargest(1, 'Price Change 24h (%)')[['Name', 'Price Change 24h (%)']].iloc[0],
            'Lowest 24h Change': df.nsmallest(1, 'Price Change 24h (%)')[['Name', 'Price Change 24h (%)']].iloc[0]
        }
        return analysis

    def update_excel(self, df, analysis):
        """Update Excel file with latest data and analysis"""
        with pd.ExcelWriter(self.excel_file, engine='openpyxl', mode='w') as writer:
            # Write main data
            df.to_excel(writer, sheet_name='Live Crypto Data', index=False)
            
            # Write analysis
            analysis_df = pd.DataFrame()
            analysis_df['Top 5 by Market Cap'] = analysis['Top 5 by Market Cap']['Name'].values
            analysis_df['Market Cap'] = analysis['Top 5 by Market Cap']['Market Cap'].values
            analysis_df.to_excel(writer, sheet_name='Analysis', startrow=1, index=False)
            
            # Write summary statistics
            summary = pd.DataFrame({
                'Metric': ['Average Price (USD)', 'Highest 24h Change', 'Lowest 24h Change'],
                'Value': [
                    f"${analysis['Average Price']:.2f}",
                    f"{analysis['Highest 24h Change']['Name']}: {analysis['Highest 24h Change']['Price Change 24h (%)']:.2f}%",
                    f"{analysis['Lowest 24h Change']['Name']}: {analysis['Lowest 24h Change']['Price Change 24h (%)']:.2f}%"
                ]
            })
            summary.to_excel(writer, sheet_name='Analysis', startrow=8, index=False)

    def run_continuous_update(self, interval_minutes=5):
        """Run continuous updates at specified intervals"""
        while True:
            print(f"\nFetching data at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            
            # Fetch and process data
            raw_data = self.fetch_top_50_crypto()
            df = self.process_crypto_data(raw_data)
            
            if df is not None:
                # Perform analysis
                analysis = self.analyze_data(df)
                
                # Update Excel
                self.update_excel(df, analysis)
                print("Excel file updated successfully!")
                
                # Print some key statistics
                print("\nQuick Summary:")
                print(f"Top Cryptocurrency: {df.iloc[0]['Name']} at ${df.iloc[0]['Current Price (USD)']:.2f}")
                print(f"Average Price: ${analysis['Average Price']:.2f}")
            
            # Wait for the specified interval
            time.sleep(interval_minutes * 60)

if __name__ == "__main__":
    fetcher = CryptoDataFetcher()
    fetcher.run_continuous_update()