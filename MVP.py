import streamlit as st
import pandas as pd
from yahooquery import Ticker
from datetime import datetime
from docx import Document
import base64
from download_button_function import download_button
import io
from create_word_doc import create_doc
import matplotlib.pyplot as plt
import seaborn as sns
import requests
from numerize import numerize
from PIL import Image

# To run use streamlit run streamlit.py
# to exit, use control-c
st.set_page_config(layout="wide")

st.sidebar.subheader("""DDR Edge Stock Analysis""")
user_input_ticker = st.sidebar.text_input("Enter the ticker of the stock","AAPL")
author = st.sidebar.text_input("Your name", "Rob Geurts")
user_input_competitors = st.sidebar.text_input("Enter names/tickers of competitor stocks, separate stocks with a comma without space:", "MSFT,TSLA")  # Changed from Dutch to English
# Date = current date
date = datetime.today().date

col1, col2, col3 = st.columns(3)





def main():
    col1.subheader('Company Information')
    col2.subheader('Competition and Price Chart')
    col3.subheader('Dividend Payouts')

    data_1_laden = col1.text('Loading...')
    ticker = user_input_ticker
    blueprint = Ticker(user_input_ticker)

    #company name
    companyName = 'Name: ' + blueprint.quote_type[ticker]['longName']
    shortcompanyName = blueprint.quote_type[ticker]['shortName']
    # Sector
    sector = 'Sector: ' + blueprint.asset_profile[ticker]['sector']
    # Industry
    industry = 'Industry: ' + blueprint.asset_profile[ticker]['industry']
    # Current price
    current_price = 'Current price: ' + str(round(blueprint.financial_data[ticker]['currentPrice'],2))
    # 52 weeks high and low
    fiftyTwoWeek = "52 weeks l/h: " + str(round(blueprint.summary_detail[ticker]['fiftyTwoWeekLow'],2)) + ' - ' + str(round(blueprint.summary_detail[ticker]['fiftyTwoWeekHigh'],2))
    # 1 year target
    targetMeanPrice = "1 Year Target: " + str(blueprint.financial_data[ticker]['targetMeanPrice'])
    # Market cap
    marketCap = "Market Cap: " + str(numerize.numerize(blueprint.summary_detail[ticker]['marketCap']))

    if type(blueprint.summary_detail[ticker]['trailingAnnualDividendRate']) == type(None):
        dividendRate = "Dividend N/A"
    else:
        dividendRate = "Forward Dividend: " + str(blueprint.summary_detail[ticker]['trailingAnnualDividendRate']) + " (" + str(
            round(blueprint.summary_detail[ticker]['trailingAnnualDividendYield'] * 100, 2)) + "%)"
        
    # Beta
    if type(blueprint.summary_detail[ticker]['beta']) == type(None):
        beta = "Beta N/A"
    else:
        beta = "Beta: " + str(round(blueprint.summary_detail[ticker]['beta'], 3))

    # Company info 
    companyInfo = blueprint.asset_profile[ticker]['longBusinessSummary']  # company info


    col1.write(companyName)
    col1.write(sector)
    col1.write(industry)
    col1.write(current_price)
    col1.write(fiftyTwoWeek)
    col1.write(targetMeanPrice)
    col1.write(marketCap)
    col1.write(beta)
    col1.write(dividendRate)
    col1.write(companyInfo)
    data_1_laden.text('Loading... ready')
    data_2_laden = col2.text('Loading...')
    user_list = [user_input_ticker]
    competitorlist = user_input_competitors.split(",")
    comparingList = user_list + competitorlist

    if user_input_competitors == "":
        comparingList = user_list

    #Initialize DF for comparison
    compare_df = pd.DataFrame(columns=[
        "Company name", "Total D/E", "Current Ratio", "Trailing P/E",
        "Return on Equity", "Profit Margin", "Trailing Annual Dividend Yield", "Enterprise value/EBITDA"
        ])

    for company in comparingList:
        items = []
        tickerb = company
        blueprintb = Ticker(company)
        items.append(blueprintb.quote_type[tickerb]['shortName'])
        items.append(blueprintb.financial_data[tickerb]['debtToEquity'])
        items.append(blueprintb.financial_data[tickerb]['currentRatio'])
        items.append(blueprintb.summary_detail[tickerb]['trailingPE'])
        items.append(blueprintb.financial_data[tickerb]['returnOnEquity'])
        items.append(blueprintb.financial_data[tickerb]['profitMargins'])
        items.append(blueprintb.summary_detail[tickerb]['trailingAnnualDividendYield'])
        items.append(blueprintb.key_stats[tickerb]['enterpriseToEbitda'])
        compare_df.loc[len(compare_df)] = items


    # Graph of stock for 2 years ( year input by user)

    ## Short positions
    # Short Ratio
    shortRatio = "Short Ratio: " + str(blueprint.key_stats[ticker]['shortRatio'])
    # Short % of shares outstanding
    shortPercentage = "Short % of Shares Outstanding: " + str(blueprint.key_stats[ticker]['shortPercentOfFloat'])

    print(compare_df)
    # News regarding company
    #news = ticker.news
    news = "no news"

    col2.dataframe(compare_df)

    #Graph
    # Reset the index of the DataFrame
    graph_data_reset = graph_data.reset_index()

    # Create a new DataFrame with 'date' and 'adjclose' columns
    adjclose_df = graph_data_reset[['date', 'adjclose']]

    # Then pass it to the line_chart method
    col2.line_chart(adjclose_df.set_index('date'))



    data_2_laden.text('Loading... ready')

    #graph to use in document
    memfile = io.BytesIO()
    sns.lineplot(data = graph_data.Close)
    plt.xticks(rotation=30)
    plt.savefig(memfile)



    # Dividend history
    data_3_laden = col3.text('Loading...')
    dividend_df = blueprint.history(period='2y')
    dividend_df = dividend_df[dividend_df['dividends']>0]
    dividend_df.reset_index(inplace=True)
    dividend_df.set_index(pd.to_datetime(dividend_df['date']), inplace=True, drop=True)
    dividend_df = dividend_df[['dividends']]
    dividend_df.index = dividend_df.index.strftime('%Y-%m-%d')
    dividend_df = dividend_df.sort_values(by='date', ascending=False)
    col3.dataframe(dividend_df)

    col3.subheader("Short ratios")
    col3.write(shortRatio)
    col3.write(shortPercentage)


    data_3_laden.text('loading... ready')
    document = create_doc(companyName, sector, industry, current_price,
               fiftyTwoWeek, targetMeanPrice,
               marketCap, beta, dividendRate, companyInfo,
               shortRatio, shortPercentage, dividend_df,
               news, compare_df,memfile, author)

    file_stream = io.BytesIO()
    # Save the .docx to the buffer
    document.save(file_stream)
    # Reset the buffer's file-pointer to the beginning of the file
    file_stream.seek(0)
    # convert doc to b64
    b64 = base64.b64encode(file_stream.getvalue()).decode()
    filename = "Analysis " + shortcompanyName + ".docx"
    download_button_str = download_button(b64, filename, f'Click here to download {filename}')
    st.sidebar.markdown(download_button_str, unsafe_allow_html=True)


if st.sidebar.button('GO'):
    main()

st.sidebar.write("Once the document is ready, a download link will appear here")
#if st.button('test'):
    #main()

#main()

#if __name__ == "__main__":
   # main()



