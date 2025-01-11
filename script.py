from binance.client import Client
import pandas as pd
import schedule
import time

BINANCE_API_KEY = 'Gs2TkXOCDmtpXCfGbAiWKOYAulL7EZq99UWJheQFxlZ2hAtPeEh4rgstiMDcAq7i'
BINANCE_SECRET = 'WG4ITosmDwegvH3wVYo1xamJFrVurjqbNr2BbHWOqg4R4qO3nQjJj6zGqCDWZIpe'

client = Client(BINANCE_API_KEY, BINANCE_SECRET)

def fetch_top_50_cryptos():
    tickers = client.get_ticker()  

    sorted_tickers = sorted(tickers, key=lambda x: float(x['quoteVolume']), reverse=True)
    sorted_tickers = [ticker for ticker in sorted_tickers if ticker['symbol'].endswith('USDT')][:50]

    data = [{
        "Name": ticker['symbol'].replace('USDT', ''),  
        "Symbol": ticker['symbol'],
        "Price (USD)": float(ticker['lastPrice']),
        "Market Cap": float(ticker['quoteVolume']) * float(ticker['lastPrice']),
        "24h Change (%)": float(ticker['priceChangePercent']),
        "Volume (24h)": float(ticker['quoteVolume'])
    } for ticker in sorted_tickers]

    return pd.DataFrame(data)

def analyze_data(data):
    top_5 = data.nlargest(5, 'Market Cap')[['Name', 'Market Cap']]

    avg_price = data['Price (USD)'].mean()


    highest_change = data.loc[data['24h Change (%)'].idxmax()]
    lowest_change = data.loc[data['24h Change (%)'].idxmin()]

    return {
        "Top 5 by Market Cap": top_5,
        "Average Price": avg_price,
        "Highest Change (24h)": highest_change,
        "Lowest Change (24h)": lowest_change
    }


def update_excel():
    data = fetch_top_50_cryptos()  
    analysis = analyze_data(data)  

    file_name = "Crypto_Live_Data.xlsx"

    with pd.ExcelWriter(file_name, mode='w', engine='openpyxl') as writer:
       
        data.to_excel(writer, index=False, sheet_name="Top 50 Cryptos")

      
        analysis_summary = {
            "Metric": ["Average Price", "Highest Change (24h)", "Lowest Change (24h)"],
            "Value": [
                f"${analysis['Average Price']:.2f}",
                f"{analysis['Highest Change (24h)']['Name']} ({analysis['Highest Change (24h)']['24h Change (%)']}%)",
                f"{analysis['Lowest Change (24h)']['Name']} ({analysis['Lowest Change (24h)']['24h Change (%)']}%)"
            ]
        }
        analysis_df = pd.DataFrame(analysis_summary)
        analysis["Top 5 by Market Cap"].to_excel(writer, index=False, sheet_name="Top 5 by Market Cap")
        analysis_df.to_excel(writer, index=False, sheet_name="Analysis Summary")

    print(f"Excel updated at {time.ctime()}")

schedule.every(5).minutes.do(update_excel)

update_excel()

print("Live updating started...")
while True:
    schedule.run_pending()
    time.sleep(1)
