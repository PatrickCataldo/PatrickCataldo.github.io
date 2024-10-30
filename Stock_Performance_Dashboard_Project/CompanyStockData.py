# --------------------------------------------------------- Program Information
# Name:             Company Stock Data Retrieval Program
# Author:           Patrick Cataldo
# Date Created:     10/28/2024
# Date Modified:    10/28/2024
# Purpose:          To retrieve stock and financial data for selected companies
#                   to be used in the stock analysis dashboard.
# -----------------------------------------------------------------------------

# --------------------------------------------------- Import Required Libraries
import yfinance as yf         # Used to get financial data
import pandas as pd           # Used to store the data as csv files
import os                     # Used to save files to selected locations
import csv                    # Used to initialize CSV files
import datetime               # Used for date ranges

import requests
from bs4 import BeautifulSoup

# ----------------------------------------------------- 1. Initialize csv files
dataFolder = os.path.join(os.getcwd(), "data")
files = {
    "CompanyData.csv": 
        ["Name", "Ticker", "Sector", "Industry", "MarketCap", 
         "No. of Employees", "Address", "Revenue", "Net Income", 
         "P/E Ratio (Trailing)", "Logo URL"], 
    "StockData.csv":
        ["Ticker", "Date", "Open", "High", "Low", "Close", "Adj Close", 
         "Volume", "Daily Return", "Cumulative Return", "MA50", "MA200", 
         "RSI", "YTD Return", "52_Week_High", "52_Week_Low"],
        "SP500Data.csv":
            ["Ticker", "Date", "Open", "High", "Low", "Close", "Adj Close", 
             "Volume", "Daily Return", "Cumulative Return", "MA50", "MA200", 
             "RSI", "YTD Return", "52_Week_High", "52_Week_Low"]}

if not os.path.exists(dataFolder):    
    os.makedirs(dataFolder)
    
os.chdir(dataFolder)

for file, headerList in zip(files, files.values()):
    #print(file)
    #print(headerList)
    filePath = os.path.join(dataFolder, file)
    
    with open(filePath, "w", newline ="") as f:
        dw = csv.DictWriter(f, delimiter=",",
                              fieldnames=headerList)
        dw.writeheader()
        
print("CSV Files Initialized")
        
# --------------------------------------------------------- 2. Get Company Data
def getCompanyInfo(ticker):
    
    print(f"Getting Company Info...")
    # Load in company information
    companyInfo = yf.Ticker(ticker)
    # Name
    companyName = companyInfo.info["longName"]
    
    # Sector and Industry
    sector = companyInfo.info["sector"]
    industry = companyInfo.info["industry"]
    
    # Market Cap
    marketCap = companyInfo.info["marketCap"]
    
    # Number of Employees
    numEmployees = companyInfo.info["fullTimeEmployees"]
    
    # Address
    street = companyInfo.info["address1"]
    city = companyInfo.info["city"]
    state = companyInfo.info["state"]
    zipCode = companyInfo.info["zip"]
    fullAddress = f"{street} {city}, {state} {zipCode}"
    
    # Financial Info
    revenue = companyInfo.info["totalRevenue"]
    netIncome = companyInfo.info["netIncomeToCommon"]
    peRatio = companyInfo.info["trailingPE"]
    
    # Company Logo
    shortName = companyInfo.info["shortName"]
    first_part = ''
    for char in shortName:
        if char.isalpha() or char == '-':
            first_part += char
        else:
            break
    
    # Construct the URL to the company's logo page
    company_name_formatted = first_part.replace(' ', '-')
    logo_page_url = f"https://companieslogo.com/{company_name_formatted}/logo/"
    print(logo_page_url)

    # Retrieve the logo download link
    logo_download_link = getLogoDownloadLink(logo_page_url)

    
    # temp message
    linebreak = "-" * 60
    message = f"""{linebreak}\n\tCompany Name:\t\t{companyName}\n{linebreak}
    Ticker Symbol:\t\t{ticker}\n
    Sector:\t\t\t\t{sector}
    Industry:\t\t\t{industry}\n
    Market Cap:\t\t\t{marketCap}
    Employees:\t\t\t{numEmployees}\n
    Address:\n\t{fullAddress}\n{linebreak}"""
    
    # Put info into a csv file using pandas
    companyData = {
        "Name": companyName,
        "Ticker": ticker,
        "Sector": sector,
        "Industry": industry,
        "Market Cap": marketCap,
        "No. of Employees": numEmployees,
        "Address": fullAddress,
        "Revenue": revenue,
        "Net Income": netIncome,
        "P/E Ratio (Trailing)": peRatio,
        "Logo URL": logo_download_link
        }
    
    df = pd.DataFrame([companyData])
    
    df.to_csv(os.path.join(dataFolder, "CompanyData.csv"), mode="a", 
              index=False, header=False)
    
    #print(companyInfo.info)
    
    print("Successfully stored company info!")
    return message

# ----------------------------------------------------------- 3. Get Stock Data
def getStockData(ticker):
    print(f"Getting Stock Info...")
    # Define a 5-year date range
    endYear = datetime.datetime.now().year - 1
    startYear = endYear - 4

    endDate = f"{endYear}-12-31"
    startDate = f"{startYear}-01-01"

    # Get Stock Data in daily intervals (default)
    data = yf.download(ticker, start=startDate, end=endDate)
    data.reset_index(inplace=True)

    # Flatten MultiIndex columns if necessary
    if isinstance(data.columns, pd.MultiIndex):
        data.columns = data.columns.get_level_values(0)

    # Ensure 'Date' is datetime and sort data
    data['Date'] = pd.to_datetime(data['Date'])
    data.sort_values('Date', inplace=True)
    data.reset_index(drop=True, inplace=True)

    # Calculate Daily Returns
    data["Daily Return"] = data["Adj Close"].pct_change()

    # Calculate Cumulative Returns
    data["Cumulative Return"] = (data["Adj Close"] / data["Adj Close"].iloc[0]) - 1

    # Calculate 50 & 200 period moving averages
    data["MA50"] = data["Adj Close"].rolling(window=50).mean()
    data["MA200"] = data["Adj Close"].rolling(window=200).mean()

    # Calculate RSI over a 14-day period
    period = 14
    delta = data["Adj Close"].diff(1)
    gain = delta.clip(lower=0)
    loss = -delta.clip(upper=0)
    avg_gain = gain.rolling(window=period).mean()
    avg_loss = loss.rolling(window=period).mean()
    rs = avg_gain / avg_loss
    data['RSI'] = 100 - (100 / (1 + rs))

    # Calculate YTD Return for each year
    data['Year'] = data['Date'].dt.year
    grouped = data.groupby('Year')

    ytd_list = []

    for year, group in grouped:
        # Get the Adjusted Close price on the first trading day of the year
        first_price = group.iloc[0]['Adj Close']

        # Calculate YTD Return for each date in the year
        group['YTD Return'] = (group['Adj Close'] - first_price) / first_price

        ytd_list.append(group)

    # Concatenate all yearly dataframes
    data = pd.concat(ytd_list)
    data.sort_values('Date', inplace=True)
    data.reset_index(drop=True, inplace=True)

    # Remove 'Year' column if not needed
    data.drop('Year', axis=1, inplace=True)

    # Calculate 52-week high and low
    data['52_Week_High'] = data['Adj Close'].rolling(window=252).max()
    data['52_Week_Low'] = data['Adj Close'].rolling(window=252).min()

    # Add Ticker Symbol
    data["Ticker"] = ticker

    # Rearrange Columns
    columns_order = ["Ticker", "Date", "Open", "High", "Low", "Close", "Adj Close",
                     "Volume", "Daily Return", "Cumulative Return", "MA50", "MA200",
                     "RSI", "YTD Return", "52_Week_High", "52_Week_Low"]
    data = data[columns_order]

    # Append to StockData.csv using pandas
    
    if ticker == "^GSPC":
        data.to_csv(os.path.join(dataFolder, "SP500Data.csv"), mode='a',
                    index=False, header=False)
    else:
        data.to_csv(os.path.join(dataFolder, "StockData.csv"), mode='a',
                index=False, header=False)
    
    print("Successfully stored stock data!")
    
# ------------------------------------------------- 5. Get Company Logo (EXTRA)
def getLogoDownloadLink(logo_page_url, logo_type='icon'):
    try:
        # Fetch the logo page
        response = requests.get(logo_page_url)
        response.raise_for_status()

        # Parse the HTML content
        soup = BeautifulSoup(response.text, 'html.parser')

        # Determine the section to search based on logo_type
        if logo_type == 'icon':
            # Find the h2 tag with text containing "icon/symbol"
            h2_tags = soup.find_all('h2', class_='logo-title')
            for h2 in h2_tags:
                if 'icon/symbol' in h2.text:
                    # The logo-section div following the h2 tag
                    logo_section = h2.find_next_sibling('div', class_='logo-section')
                    break
            else:
                print("Icon/symbol logo section not found.")
                return 'Not Found'
        else:
            # For full-size logo, find the first h2 tag
            logo_section = soup.find('h2', class_='logo-title').find_next_sibling('div', class_='logo-section')

        # Find the SVG logo download link within the selected logo section
        download_buttons = logo_section.find_all('a', href=True)
        for button in download_buttons:
            href = button['href']
            if '.svg' in href and 'download=true' in href:
                logo_download_link = 'https://companieslogo.com' + href
                return logo_download_link

        # If SVG logo not found, return 'Not Found'
        return 'Not Found'

    except Exception as e:
        print(f"Error retrieving logo download link: {e}")
        return ''

    
# ------------------------------------------------------------- 4. Main Program
if __name__ == "__main__":
    companies = ["MKL", "AAPL", "KO", "AMZN", "TSLA", "SBUX", "NVDA", "TGT", "NKE"]
    linebreak = "-" * 79
    print(f"{linebreak}\nSelected Companies: {companies}\n{linebreak}")

counter = 0
for company in companies:
    counter += 1
    print(linebreak)
    print(f"{counter}. {company}")
    getCompanyInfo(company)
    print()
    getStockData(company)
    print()

print(linebreak)
counter += 1
print(f"{counter}. S&P500")
getStockData("^GSPC")