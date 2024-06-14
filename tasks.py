from robocorp.tasks import task
from src.news_scrapper import NewsScraper
@task
def scrapper_task():
    scraper = NewsScraper("config/config.json")
    scraper.search_news()
    scraper.close_browser()
