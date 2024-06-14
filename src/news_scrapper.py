"""
******************************************************************************
*   PROJECT     :   rpa-challenge-orac
*   UTIL NAME   :   tasks
*   AUTHOR      :   oscaralvarez
*   CR DATE     :   6/4/24
******************************************************************************
*   OBJECTIVE   :   tasks
******************************************************************************
"""

import os
import re
import json
import logging
from datetime import datetime, timedelta
from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files
from RPA.HTTP import HTTP

class NewsScraper:
    def __init__(self, config_path):
        self.browser = Selenium()
        self.excel = Files()
        self.http = HTTP()
        self.config = self.load_config(config_path)
        self.output_dir = "output"
        os.makedirs(self.output_dir, exist_ok=True)
        self.setup_logging()

    def setup_logging(self):
        log_filename = os.path.join(self.output_dir, 'scraper.log')
        logging.basicConfig(filename=log_filename, level=logging.INFO,
                            format='%(asctime)s - %(levelname)s - %(message)s')
        logging.info('NewsScraper initialized.')

    def load_config(self, config_path):
        try:
            with open(config_path, 'r') as file:
                config = json.load(file)
                logging.info('Configuration loaded.')
                return config
        except Exception as e:
            logging.error(f"Error loading configuration, details: {e}")
            raise

    def search_news(self):
        news_data = []
        try:
            search_phrase = self.config["search_phrase"]
            category = self.config["category"]
            months = self.config["months"]
            source = self.config["source"]

            logging.info(f'Searching news for phrase: {search_phrase}, category: {category}, source: {source}')

            self.browser.open_available_browser(source)
            self.browser.click_element("class:SearchOverlay-search-button")
            self.browser.wait_until_element_is_visible("name:q", 5)
            self.browser.input_text("name:q", search_phrase)
            self.browser.click_element("class:SearchOverlay-search-submit")
            self.browser.wait_until_page_contains("Results for")

            self.dismiss_overlays()
            results = self.browser.find_element("class:SearchResultsModule-results")
            articles = self.browser.find_elements("class:PageList-items-item", results)

            for article in articles:
                try:
                    title_element = article.find_element("css selector", ".PagePromoContentIcons-text")
                    date_element = article.find_element("css selector", ".Timestamp-template")
                    description_article = article.find_element("css selector", ".PagePromo-description")
                    description_element = description_article.find_element("css selector", ".PagePromoContentIcons-text")
                    image_container = article.find_element("css selector", ".PagePromo-media")
                    image_element = image_container.find_element("css selector", ".Image")

                    title = title_element.text
                    date_str = date_element.get_attribute("datetime")
                    description = description_element.text if description_element else ""
                    image_url = image_element.get_attribute("src")

                    article_date = datetime.fromisoformat(date_str[:-1]) if date_str else datetime.now()

                    if self.is_within_months(article_date, months):
                        image_filename = self.download_image(image_url)
                        news_article = {
                            "title": self.clean_text(title),
                            "date": article_date, #.strftime("%Y-%m-%d"),
                            "description": self.clean_text(description),
                            "image_filename": image_filename,
                            "search_phrase_count": self.count_search_phrases(title, description, search_phrase),
                            "contains_money": self.contains_money(title, description)
                        }
                        print(f'News article found: {news_article}')

                        logging.info(f'News article found: {news_article}')
                        news_data.append(news_article)
                except Exception as e:
                    logging.error(f"Error processing article, details: {e}")

            logging.info('News search completed successfully.')
        except Exception as e:
            logging.error(f"Error processing the news search, details: {e}")
        finally:
            self.save_to_excel(news_data)

    def is_within_months(self, article_date, months):
        current_date = datetime.now()
        past_date = current_date - timedelta(days=30 * months)
        return past_date <= article_date <= current_date

    def count_search_phrases(self, title, description, search_phrase):
        return title.lower().count(search_phrase.lower()) + description.lower().count(search_phrase.lower())

    def contains_money(self, title, description):
        money_pattern = re.compile(r"\$[\d,]+(\.\d+)?|\b\d+\s+(dollars|usd)\b", re.IGNORECASE)
        return bool(money_pattern.search(title) or money_pattern.search(description))

    def download_image(self, url):
        try:
            filename = os.path.join(self.output_dir, url.split("/")[-1]) + ".jpg"
            self.http.download(url, filename)
            logging.info(f"Downloaded image from {url} to {filename}.")
            return filename
        except Exception as e:
            logging.error(f"Error downloading image from {url}, details: {e}")
            return ""

    def save_to_excel(self, news_data):
        try:
            output_file = os.path.join(self.output_dir, "news_data.xlsx")
            self.excel.create_workbook(output_file)
            self.excel.create_worksheet("News")
            self.excel.append_rows_to_worksheet([["Title", "Date", "Description", "Image Filename", "Search Phrase Count", "Contains Money"]], "News")
            for data in news_data:
                self.excel.append_rows_to_worksheet([[
                    data["title"],
                    data["date"],
                    data["description"],
                    data["image_filename"],
                    data["search_phrase_count"],
                    data["contains_money"]
                ]], "News")
            self.excel.save_workbook()
            logging.info(f"Saved news data to {output_file}.")
        except Exception as e:
            logging.error(f"Error saving data to Excel, details: {e}")

    def clean_text(self, text):
        return re.sub(r'[^a-zA-Z0-9\s]', '', text)

    def close_browser(self):
        try:
            self.browser.close_browser()
            logging.info('Browser closed.')
        except Exception as e:
            logging.error(f"Error closing browser, details: {e}")


    def dismiss_overlays(self):
        try:
            self.browser.wait_until_element_is_visible("css:.onetrust-close-btn-handler", timeout=10)
            self.browser.click_element("css:.onetrust-close-btn-handler")
        except Exception as e:
            logging.info(f"No overlay to dismiss, details: {e}")