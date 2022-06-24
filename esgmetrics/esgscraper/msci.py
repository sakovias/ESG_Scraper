"""
MSCI website Scrape

This script allows the user to scrape the companies' ESG ratings from the MSCI
website. Website link:
"https://www.msci.com/research-and-insights/esg-ratings-corporate-search-tool"

This tool accepts Company's names list in comma separated value
file (.csv) format as input.

This script requires that `pandas` be installed within the Python
environment you are running this script in.

The output is a .csv file with Company name and its corresponding ESG ratings
"""

import pandas as pd
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.keys import Keys
from time import sleep
#from tqdm import tqdm
from .scraper import WebScraper
import traceback


def get_esg_score(bot, header_name, df, i):
    # Starting the search by finding the search bar & searching for the company
    search_bar = bot.send_request_to_search_bar(
        header_name, df, i, xpath='//*[@id="_esgratingsprofile_keywords"]')
    search_bar.send_keys(Keys.DOWN, Keys.RETURN)
    sleep(4)

    xpath = '//*[@id="_esgratingsprofile_esg-ratings-profile-header"]/div[1]/div[1]/div[2]/div[1]'
    esg_score_element = bot.find_element(xpath)
    esg_score = esg_score_element.get_attribute('class').split('-')[-1].upper()
    print('score', esg_score)

    company_msci_name = bot.find_element('//*[@class="header-company-title"]').text
    return {
        'MSCI_Company': [company_msci_name],
        'MSCI_ESG': [esg_score]
    }

def init_bot(chrome_path):
    # Set up the webdriver
    URL = "https://www.msci.com/research-and-insights/esg-ratings-corporate-search-tool"
    bot = WebScraper(URL, chrome_path)

    # Accept cookies on the website
    cookies_xpath = '//*[@id="onetrust-accept-btn-handler"]'
    bot.accept_cookies(cookies_xpath)
    return bot

# Read input companies dataset
companies_filename = WebScraper._get_filename()
header_name = WebScraper._get_headername()
export_path = WebScraper._get_exportpath()
df = pd.read_csv(companies_filename)
data_length = len(df)

chrome_path = input('Please specify the chromedriver path : ')
bot = init_bot(chrome_path)

# Extract company names and their ESG score and store it in the dictionary
for i in range(data_length):
    company = df.loc[i][header_name]
    print('processing', i, 'of', data_length, company)
    try:
        msci_data = get_esg_score(bot, header_name, df, i)
        # Save the data into a csv file
        bot.convert_dict_to_csv(msci_data, export_path)

    except Exception as e:
        print('Failed on', company)
        print('Exception', e.__class__.__name__, e)
        traceback.print_exc()
        bot.take_screenshot()

        # try closing the customer intake modal
        try:
            bot.try_closing_modal('//*[@class="yui3-widget-hd modal-header"]//div[1]//button')
            
            # reinstantiate bot
            bot = init_bot(chrome_path)
            msci_data = get_esg_score(bot, header_name, df, i)
            # Save the data into a csv file
            bot.convert_dict_to_csv(msci_data, export_path)
        except Exception as e:
            print('Yet another exception', e.__class__.__name__, e)
            traceback.print_exc()
            break
        

