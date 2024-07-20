import requests
from bs4 import BeautifulSoup
import io
import gzip
import chardet
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import csv
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time


scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("webscraper-429822-2166f9791a8b.json", scope)
client = gspread.authorize(creds)

spreadsheet = client.open("Webscraper Project")
worksheet = spreadsheet.sheet1 

worksheet.update_cell(1, 1,"Website/Category")
worksheet.update_cell(1, 2, "service offerings")
worksheet.update_cell(1, 3, "case studies")
worksheet.update_cell(1, 4, "client testimonials")
worksheet.update_cell(1, 5, "thought leadership")
worksheet.update_cell(1, 6, "market insights")
worksheet.update_cell(1, 7, "latest news articles")
worksheet.update_cell(1, 8, "market reports")
worksheet.update_cell(1, 9, "financial analysis")
worksheet.update_cell(1, 10, "global economic trends")
worksheet.update_cell(1, 11, "industry analysis")
worksheet.update_cell(1, 12, "statistics")
worksheet.update_cell(1, 13, "market research reports")
worksheet.update_cell(1, 14, "trends")
worksheet.update_cell(1, 15, "global news")

worksheet.update_cell(2, 1, "Boston Consulting Group")
worksheet.update_cell(3, 1, "McKinsey & Company")
worksheet.update_cell(4, 1, "Deloitte")
worksheet.update_cell(5, 1, "Bloomberg")
worksheet.update_cell(6, 1, "Reuters")
worksheet.update_cell(7, 1, "Financial Times")
worksheet.update_cell(8, 1, "Statista")
worksheet.update_cell(9, 1, "MarketResearch")
worksheet.update_cell(10, 1, "IBISWorld")
worksheet.update_cell(11, 1, "Economic Times")
worksheet.update_cell(12, 1, "Business Standard")
worksheet.update_cell(13, 1, "Hindustan Times")
worksheet.update_cell(14, 1, "China Daily")
worksheet.update_cell(15, 1, "Global Times")
worksheet.update_cell(16, 1, "Caixin Global")
worksheet.update_cell(17, 1, "Japan Times")
worksheet.update_cell(18, 1, "Asia Nikkei")
worksheet.update_cell(19, 1, "NHK World")
worksheet.update_cell(20, 1, "Straits Times")
worksheet.update_cell(21, 1, "Channel News Asia")
worksheet.update_cell(22, 1, "Korea Herald")
worksheet.update_cell(23, 1, "Korea Times")
worksheet.update_cell(24, 1, "Yonhap News")
worksheet.update_cell(25, 1, "Gulf News")
worksheet.update_cell(26, 1, "Khaleej Times")
worksheet.update_cell(27, 1, "The National")
worksheet.update_cell(28, 1, "Taipei Times")
worksheet.update_cell(29, 1, "Focus Taiwan")
worksheet.update_cell(30, 1, "Taiwan News")
worksheet.update_cell(31, 1, "Weblinks")

headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
    'Accept-Language': 'en-US,en;q=0.5',
    'Accept-Encoding': 'gzip, deflate, br',
    'Connection': 'keep-alive',
    'Referer': 'https://www.google.com/'
}

# Lists of websites to scrape

competitorWebsitesUrls = {
    # "Boston Consulting Group": "https://www.bcg.com",
    # "McKinsey & Company": "https://www.mckinsey.com",
    # "Deloitte": "https://www.deloitte.com"
}

industryNewsReportsUrls = {
    # "Bloomberg": "https://www.bloomberg.com",
    # "Reuters": "https://www.reuters.com", 
    # "Financial Times": "https://www.ft.com"
}

marketResearchPortalsUrls = {
    # "Statista": "https://www.statista.com",
    # "MarketResearch.com": "https://www.marketresearch.com",
    # "IBISWorld": "https://www.ibisworld.com"
}

globalNewsOutletsUrls = {
    # # India:
    # "The Economic Times": "https://economictimes.indiatimes.com",
    # "Business Standard": "https://www.business-standard.com", 
    # "Hindustan Times": "https://www.hindustantimes.com", 

    # # China:
    # "China Daily": "https://www.chinadaily.com.cn",
    # "Global Times": "https://www.globaltimes.cn",
    # "Caixin Global": "https://www.caixinglobal.com",

    # # Japan:
    # "The Japan Times": "https://www.japantimes.co.jp",
    # "Nikkei Asia": "https://asia.nikkei.com",
    # "NHK World": "https://www.nhk.or.jp/nhkworld",

    # # Singapore:
    # "The Straits Times": "https://straitstimes.com", 
    "Channel News Asia": "https://www.channelnewsasia.com",
    "The Business Times": "https://www.businesstimes.com.sg",

    # South Korea:
    "The Korea Herald": "https://www.koreaherald.com",
    "The Korea Times": "https://www.koreatimes.co.kr",
    "Yonhap News Agency": "https://www.yna.co.kr",

    # Dubai (UAE):
    "Gulf News": "https://www.gulfnews.com",
    "Khaleej Times": "https://www.khaleejtimes.com",
    "The National": "https://www.thenationalnews.com",

    # Taiwan:
    "Taipei Times": "https://www.taipeitimes.com",
    "Focus Taiwan": "https://focustaiwan.tw",
    "Taiwan News": "https://www.taiwannews.com.tw" 
}

def retriveContent(response,url):
    if response.headers.get('Content-Encoding') == 'gzip':
            try:
                buf = io.BytesIO(response.content)
                f = gzip.GzipFile(fileobj=buf)
                rawContent = f.read().decode('utf-8')
            except Exception as e:
                        rawContent = response.content
    else:
            rawContent = response.content
                
    # Detect encoding using chardet
    detected_encoding = chardet.detect(rawContent)['encoding']
    # print("Detected encoding for ",url,": ",detected_encoding)

    if detected_encoding == None:
        soup = BeautifulSoup(response.text, 'html.parser')
        return soup 
    else:
        content = rawContent.decode(detected_encoding)
        soup = BeautifulSoup(response.text, 'html.parser')
        return soup
            
#Function to sent HTTP requests
def scrapeWebsite(url):
    try:
        session = requests.Session()
        session.headers.update(headers)

        # Send HTTP GET request
        response = requests.get(url)
        
        # Check if the request was successful
        if response.status_code == 200:
                # print("Successfully connected to", url)
                soup = retriveContent(response,url)
                return soup
        else:
                response = requests.get(url)
                soup = retriveContent(response,url)
                if soup:
                    print("Successfully connected to", url)
                    return soup
                else:
                    print("Failed to retrieve", url,": Status code:",response.status_code)                        
                    return None
    except requests.exceptions.RequestException as e:
        print("An error occurred:", e)
        return None

def scrapeDynamic(url, className):
    driver = webdriver.Chrome()
    driver.get(url)
    html_content = driver.page_source
    elements = driver.find_elements(By.CLASS_NAME,className)
    # for element in elements:
    #     print(element.text)
    data = []
    for element in elements:
        data.append(element.text)
    driver.quit()
    return data

    
def main():

    data = {
    "service_offerings": [],
    "case_studies": [],
    "client_testimonials": [],
    "thought_leadership": [],
    "market_insights": [],
    "latest_news_articles": [],
    "market_reports": [],
    "financial_analysis": [],
    "global_economic_trends": [],
    "industry_analysis":[],
    "statistics":[],
    "market_research_reports":[],
    "trends":[],
    "global_news":[]
    }

    #parsing competitorWebsitesUrls 
    # Service offerings, case studies, client testimonials, thought leadership articles, and market insights.
    for name, url in competitorWebsitesUrls.items():
        print("----------\nScraping",name)
        soup = scrapeWebsite(url)
        
        # extracting information from each website
        if name == "Boston Consulting Group":
            data = {
                "service_offerings": [],
                "case_studies": [],
                "client_testimonials": [],
                "thought_leadership": [],
                "market_insights": []
            }
            # market insights
            driver = webdriver.Chrome()
            driver.get("https://www.bcg.com/publications")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR, '.article-block__title')
            for element in elements:
                data["market_insights"].append(element.text)
            data1 = ",".join(data["market_insights"])
            worksheet.update_cell(2,6,data1)
            linkList = []
            contentDiv = driver.find_elements(By.CSS_SELECTOR, '.article-block__content')
            for i in contentDiv:
                items = i.find_elements(By.XPATH, './/a[@href]')
                for item in items:
                    link = item.get_attribute('href')
                    print("item is",link)
                    linkList.append(link)
            data1 = ",".join(linkList)
            worksheet.update_cell(31,6,data1)
            driver.quit()

            # case studies
            driver = webdriver.Chrome()
            driver.get("https://www.bcg.com")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR, '.two-panel-slide-left-hero-promo__title')
            for element in elements:
                data["case_studies"].append(element.text)
            data1 = ",".join(data["case_studies"])
            worksheet.update_cell(2,3,data1)
            linkList = []
            contentDiv = driver.find_elements(By.CSS_SELECTOR, '.two-panel-slide-left-hero-promo__content-inner-wrapper')
            for i in contentDiv:
                items = i.find_elements(By.XPATH, './/a[@href]')
                for item in items:
                    link = item.get_attribute('href')
                    print("item is",link)
                    linkList.append(link)
            data1 = ",".join(linkList)
            worksheet.update_cell(31,3,data1)
            driver.quit()
            
            # service offerings
            response = requests.get("https://www.bcg.com/capabilities")
            soup = BeautifulSoup(response.content, 'html.parser')
            services = soup.find_all('p', class_='featured-collection__title')
            for service in services:
                 data["service_offerings"].append(service.get_text().strip())
            data1 = ",".join(data["service_offerings"])
            worksheet.update_cell(2,2,data1)
            linkList = [ ]
            contentDiv = soup.find_all('div', class_='featured-collection__wrapper')
            for i in contentDiv:
                item = i.find('a', href=True)
                if item:
                    link = item['href']
                    linkList.append(link)
            print("item linklist is",linkList)
            data1 = ",".join(linkList)
            worksheet.update_cell(31,2,data1)

            #client testimonials
            response = requests.get("https://www.bcg.com/capabilities/digital-technology-data/client-success")
            soup = BeautifulSoup(response.content, 'html.parser')
            elements = soup.find_all('div', class_='article-block__body')
            for element in elements:
                data["client_testimonials"].append(element.text)
            data1 = ",".join(data["client_testimonials"])
            worksheet.update_cell(2,4,data1)
            linkList = []
            contentDiv = soup.find_all('div', class_='article-block__content')
            for i in contentDiv:
                item = i.find('a', href=True)
                if item:
                    link = item['href']
                    print("item is",link)
                    linkList.append(link)
            data1 = ",".join(linkList)
            worksheet.update_cell(31,4,data1)

            #thought leadership articles
            articles = scrapeDynamic(url,"HomepageCardsCarouselPromo__title")
            articles = [article for article in articles if article ]
            for a in articles:
                    data["thought_leadership"].append(a)
            data1 = ",".join(data["thought_leadership"])
            worksheet.update_cell(2,5,data1)
            linkList = []
            contentDiv = soup.find_all('div', class_='glide__slides')
            for i in contentDiv:
                item = i.find('a', href=True)
                if item:
                    link = item['href']
                    print("item is",link)
                    linkList.append(link)
            data1 = ",".join(linkList)
            worksheet.update_cell(31,5,data1)

        elif name == "McKinsey & Company":
            data = {
                "service_offerings": [],
                "case_studies": [],
                "client_testimonials": [],
                "thought_leadership": [],
                "market_insights": []
            }
            # market insights
            response = requests.get("https://www.mckinsey.com/capabilities/growth-marketing-and-sales/how-we-help-clients/insights-and-analytics")
            soup = BeautifulSoup(response.content, 'html.parser')
            elements = soup.find_all('p')
            for element in elements:
                data["market_insights"].append(element.text)
            data1 = ",".join(data["market_insights"])
            worksheet.update_cell(3,6,data1)
            linkList = []
            contentDiv = soup.find_all('div', class_='mdc-c-description___SrnQP_3fbefcf')
            for i in contentDiv:
                item = i.find('a', href=True)
                if item:
                    link = item['href']
                    print("item is",link)
                    linkList.append(link)
            data1 = worksheet.cell(31,6).value + ",".join(linkList)
            worksheet.update_cell(31,6,data1)

            # case studies
            driver = webdriver.Chrome()
            driver.get("https://www.mckinsey.com/about-us/case-studies")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR, '.mdc-c-content-block___7p6Lu_3fbefcf span')
            for element in elements:
                data["case_studies"].append(element.text)
            data1 = ",".join(data["case_studies"])
            worksheet.update_cell(3,3,data1)
            linkList = []
            contentDiv = driver.find_elements(By.CSS_SELECTOR, '.mdc-c-content-block___7p6Lu_3fbefcf')
            for i in contentDiv:
                items = i.find_elements(By.XPATH, './/a[@href]')
                for item in items:
                    link = item.get_attribute('href')
                    print("item is",link)
                    linkList.append(link)
            data1 = worksheet.cell(31,3).value + ",".join(linkList)
            worksheet.update_cell(31,3,data1)
            driver.quit()


            #service offering
            driver = webdriver.Chrome()
            driver.get("https://www.mckinsey.com/locations/mckinsey-client-capabilities-network/our-work")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR, '.LinkListItems_mck-c-ll-items__list__LSn64 span')
            for element in elements:
                data["service_offerings"].append(element.text)
            data1 = ",".join(data["service_offerings"])
            worksheet.update_cell(3,2,data1)
            linkList = []
            contentDiv = driver.find_elements(By.CSS_SELECTOR, '.LinkListItems_mck-c-ll-items__list__LSn64')
            for i in contentDiv:
                items = i.find_elements(By.XPATH, './/a[@href]')
                for item in items:
                    link = item.get_attribute('href')
                    print("item is",link)
                    linkList.append(link)
            data1 = worksheet.cell(31,2).value + ",".join(linkList)
            worksheet.update_cell(31,2,data1)
            driver.quit()

            # thought leadership articles
            driver = webdriver.Chrome()
            driver.get("https://www.mckinsey.com")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR, '.mdc-c-heading___0fM1W_3fbefcf mdc-u-ts-6 h3')
            for element in elements:
                data["thought_leadership"].append(element.text)
            data1 = ",".join(data["thought_leadership"])
            worksheet.update_cell(3,5,data1)
            linkList = []
            contentDiv = driver.find_elements(By.CSS_SELECTOR, '.LinkListItems_mck-c-ll-items__list__LSn64')
            for i in contentDiv:
                items = i.find_elements(By.XPATH, './/a[@href]')
                for item in items:
                    link = item.get_attribute('href')
                    print("item is",link)
                    linkList.append(link)
            data1 = worksheet.cell(31,5).value + ",".join(linkList)
            worksheet.update_cell(31,5,data1)
            driver.quit()

            #client testimonials
            driver = webdriver.Chrome()
            driver.get("https://www.mckinsey.com/capabilities/operations/how-we-help-clients/operations-excellence-program")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR, '.mdc-c-heading___0fM1W_3fbefcf h2')
            for element in elements:
                data["client_testimonials"].append(element.text)
            data1 = ",".join(data["client_testimonials"])
            worksheet.update_cell(3,4,data1)
            linkList = []
            contentDiv = driver.find_elements(By.CSS_SELECTOR, '.mdc-c-link-container___xefGu_3fbefcf')
            for i in contentDiv:
                items = i.find_elements(By.XPATH, './/a[@href]')
                for item in items:
                    link = item.get_attribute('href')
                    print("item is",link)
                    linkList.append(link)
            data1 = worksheet.cell(31,4).value + ",".join(linkList)
            worksheet.update_cell(31,4,data1)
            driver.quit()

        elif name == "Deloitte":
            data = {
                "service_offerings": [],
                "case_studies": [],
                "client_testimonials": [],
                "thought_leadership": [],
                "market_insights": []
            }
            # market insights
            driver = webdriver.Chrome()
            driver.get(url)
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR, '.trending-item h4')
            for element in elements:
                data["market_insights"].append(element.text)
            data1 = ",".join(data["market_insights"])
            worksheet.update_cell(4,6,data1)
            linkList = []
            contentDiv = driver.find_elements(By.CSS_SELECTOR, '.featuredpromo')
            for i in contentDiv:
                items = i.find_elements(By.XPATH, './/a[@href]')
                for item in items:
                    link = item.get_attribute('href')
                    print("item is",link)
                    linkList.append(link)
            data1 = worksheet.cell(31,6).value + ",".join(linkList)
            worksheet.update_cell(31,6,data1)
            

            # case studies
            elements = driver.find_elements(By.CSS_SELECTOR, '.promo-focus h2')
            for element in elements:
                data["case_studies"].append(element.text)
            data1 = ",".join(data["case_studies"])
            worksheet.update_cell(4,3,data1)
            linkList = []
            contentDiv = driver.find_elements(By.CSS_SELECTOR, '.featuredpromo')
            for i in contentDiv:
                items = i.find_elements(By.XPATH, './/a[@href]')
                for item in items:
                    link = item.get_attribute('href')
                    print("item is",link)
                    linkList.append(link)
            data1 = worksheet.cell(31,3).value + ",".join(linkList)
            worksheet.update_cell(31,3,data1)
            

            #service offerings
            elements = driver.find_elements(By.CSS_SELECTOR,"#footer__links-services a")
            for element in elements:
                data["service_offerings"].append(element.text)
            data1 = ",".join(data["service_offerings"])
            worksheet.update_cell(4,2,data1)
            linkList = []
            contentDiv = driver.find_elements(By.CSS_SELECTOR, '#footer__links-services')
            for i in contentDiv:
                items = i.find_elements(By.XPATH, './/a[@href]')
                for item in items:
                    link = item.get_attribute('href')
                    print("item is",link)
                    linkList.append(link)
            data1 = worksheet.cell(31,2).value + ",".join(linkList)
            worksheet.update_cell(31,2,data1)
            

            #thought leadership 
            elements = driver.find_elements(By.CSS_SELECTOR,".trending-item h5")
            for element in elements:
                data["thought_leadership"].append(element.text)
            data1 = ",".join(data["thought_leadership"])
            worksheet.update_cell(4,5,data1)
            linkList = []
            contentDiv = driver.find_elements(By.CSS_SELECTOR, '.trending-item')
            for i in contentDiv:
                items = i.find_elements(By.XPATH, './/a[@href]')
                for item in items:
                    link = item.get_attribute('href')
                    print("item is",link)
                    linkList.append(link)
            data1 = worksheet.cell(31,5).value + ",".join(linkList)
            worksheet.update_cell(31,5,data1)
            driver.quit()
            

            #client testimonial
            driver = webdriver.Chrome()
            driver.get("https://www2.deloitte.com/us/en/pages/about-deloitte/articles/client-stories-and-successes.html?icid=bottom_client-stories-and-successes")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,"#featured-client-story h2")
            for element in elements:
                 data["client_testimonials"].append(element.text)
            data1 = ",".join(data["client_testimonials"])
            worksheet.update_cell(4,4,data1)
            linkList = []
            contentDiv = driver.find_elements(By.CSS_SELECTOR, '#stories')
            for i in contentDiv:
                items = i.find_elements(By.XPATH, './/a[@href]')
                for item in items:
                    link = item.get_attribute('href')
                    print("item is",link)
                    linkList.append(link)
            data1 = worksheet.cell(31,4).value + ",".join(linkList)
            worksheet.update_cell(31,4,data1)
            driver.quit()
       
    # print("case studies\n",data['case_studies'])
    # print("market insights\n",data["market_insights"])
    # print("service offerings\n",data["service_offerings"])
    # print("thought leadership articles\n",data["thought_leadership"])
    # print("client testimonials\n",data["client_testimonials"])

    #parsing industryNewsReportsUrls
    # Latest news articles, market reports, financial analysis, and global economic trends.
    for name, url in industryNewsReportsUrls.items():
        print("----------\nScraping",name)
       
        
        if name == "Bloomberg":
            data = {
                "latest_news_articles": [],
                "market_reports": [],
                "financial_analysis": [],
                "global_economic_trends": []
            }
            # latest news
            driver = webdriver.Chrome()
            driver.get(url)
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,".StoryListLatestFiltered_storyBlock___szNS a")
            for element in elements:
                    data["latest_news_articles"].append(element.text)
            data1 = ",".join(data["latest_news_articles"])
            worksheet.update_cell(5,7,data1)
            linkList = []
            contentDiv = driver.find_elements(By.CSS_SELECTOR, '.StoryListLatestFiltered_storyBlock___szNS')
            for i in contentDiv:
                items = i.find_elements(By.XPATH, './/a[@href]')
                for item in items:
                    link = item.get_attribute('href')
                    print("item is",link)
                    linkList.append(link)
            data1 = ",".join(linkList)
            worksheet.update_cell(31,7,data1)
            driver.quit()
          
            # market reports
            driver = webdriver.Chrome()
            driver.get("https://www.bloomberg.com/markets")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,".Headline_phoenix__tgVV3 span")
            for element in elements:
                 data["market_reports"].append(element.text)
            data1 = ",".join(data["market_reports"])
            worksheet.update_cell(5,8,data1)
            linkList = []
            contentDiv = driver.find_elements(By.CSS_SELECTOR, '.StoryBlock_storyLink__w8f2K')
            for i in contentDiv:
                items = i.find_elements(By.XPATH, './/a[@href]')
                for item in items:
                    link = item.get_attribute('href')
                    print("item is",link)
                    linkList.append(link)
            data1 = ",".join(linkList)
            worksheet.update_cell(31,8,data1)
            driver.quit()

            # financial analysis
            driver = webdriver.Chrome()
            driver.get("https://www.bloomberg.com/search?query=financial%20markets")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,".headline__3a97424275")
            for element in elements:
                 data["financial_analysis"].append(element.text)
            data1 = ",".join(data)
            worksheet.update_cell(5,9,data1)
            linkList = []
            contentDiv = driver.find_elements(By.CSS_SELECTOR, '.headline__3a97424275')
            for i in contentDiv:
                items = i.find_elements(By.XPATH, './/a[@href]')
                for item in items:
                    link = item.get_attribute('href')
                    print("item is",link)
                    linkList.append(link)
            data1 = ",".join(linkList)
            worksheet.update_cell(31,9,data1)
            driver.quit()
        
            # global economic trends
            driver = webdriver.Chrome()
            driver.get("https://www.bloomberg.com/economics")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,".Headline_phoenix__tgVV3 span")
            for element in elements:
                 data["global_economic_trends"].append(element.text)
            data1 = ",".join(data["global_economic_trends"])
            worksheet.update_cell(5,10,data1)
            linkList = []
            contentDiv = driver.find_elements(By.CSS_SELECTOR, '.StoryBlock_storyLink__w8f2K')
            for i in contentDiv:
                items = i.find_elements(By.XPATH, './/a[@href]')
                for item in items:
                    link = item.get_attribute('href')
                    print("item is",link)
                    linkList.append(link)
            data1 = ",".join(linkList)
            worksheet.update_cell(31,10,data1)
            driver.quit()

        elif name == "Reuters":
            data = {
                "latest_news_articles": [],
                "market_reports": [],
                "financial_analysis": [],
                "global_economic_trends": []
            }
            # latest news
            driver = webdriver.Chrome()
            driver.get(url)
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,".basic-card__container__1y8wi span")
            for element in elements:
                 data["latest_news_articles"].append(element.text)
            data["latest_news_articles"] = [article for article in data["latest_news_articles"] if article ]
            data1 = ",".join(data["latest_news_articles"])
            worksheet.update_cell(6,7,data1)
            linkList = []
            contentDiv = driver.find_elements(By.CSS_SELECTOR, '#stories')
            for i in contentDiv:
                items = i.find_elements(By.XPATH, './/a[@href]')
                for item in items:
                    link = item.get_attribute('href')
                    print("item is",link)
                    linkList.append(link)
            data1 = worksheet.cell(31,7).value + ",".join(linkList)
            worksheet.update_cell(31,7,data1)
            driver.quit()

            #market reports
            # driver = webdriver.Chrome()
            # driver.get(url)
            # html_content = driver.page_source
            # elements = driver.find_elements(By.CSS_SELECTOR,".basic-card__container__1y8wi span")
            # for element in elements:
            #      data["latest_news_articles"].append(element.text)
            # data["latest_news_articles"] = [article for article in data["latest_news_articles"] if article ]
            # data1 = ",".join(data["latest_news_articles"])
            # worksheet.update_cell(6,7,data1)
            # linkList = []
            # contentDiv = driver.find_elements(By.CSS_SELECTOR, '#stories')
            # for i in contentDiv:
            #     items = i.find_elements(By.XPATH, './/a[@href]')
            #     for item in items:
            #         link = item.get_attribute('href')
            #         print("item is",link)
            #         linkList.append(link)
            # data1 = worksheet.cell(31,7).value + ",".join(linkList)
            # worksheet.update_cell(31,7,data1)
            # driver.quit()


        if name == "Financial Times":
            data = {
                "latest_news_articles": [],
                "market_reports": [],
                "financial_analysis": [],
                "global_economic_trends": []
            }
            #latest news
            driver = webdriver.Chrome()
            driver.get(url)
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,".headline span")
            for element in elements:
                 data["latest_news_articles"].append(element.text)
            data["latest_news_articles"] = [article for article in data["latest_news_articles"] if article]
            data1 = ",".join(data["latest_news_articles"])
            worksheet.update_cell(7,7,data1)
            linkList = []
            contentDiv = driver.find_elements(By.CSS_SELECTOR, '.headline')
            for i in contentDiv:
                items = i.find_elements(By.XPATH, './/a[@href]')
                for item in items:
                    link = item.get_attribute('href')
                    print("item is",link)
                    linkList.append(link)
            data1 = worksheet.cell(31,7).value + ",".join(linkList)
            worksheet.update_cell(31,7,data1)
            driver.quit()

            #market reports
            driver = webdriver.Chrome()
            driver.get("https://www.ft.com/markets")
            html_content = driver.page_source
            reports = driver.find_elements(By.CSS_SELECTOR,".js-teaser-heading-link")
            for report in reports:
                 data["market_reports"].append(report.text)
            data["market_reports"] = [article for article in data["market_reports"] if article]
            data["market_reports"] = [
                item for item in data["market_reports"] 
                if item.strip() and "opinion" not in item.lower() and "people" not in item.lower() and "FT" not in item.lower()
            ]
            data1 = ",".join(data["market_reports"])
            worksheet.update_cell(7,8,data1)
            linkList = []
            contentDiv = driver.find_elements(By.CSS_SELECTOR, '.o-teaser__heading')
            for i in contentDiv:
                items = i.find_elements(By.XPATH, './/a[@href]')
                for item in items:
                    link = item.get_attribute('href')
                    print("item is",link)
                    linkList.append(link)
            data1 = ",".join(linkList)
            worksheet.update_cell(31,8,data1)
            driver.quit()

            #financial analysis
            driver = webdriver.Chrome()
            driver.get("https://www.ft.com/search?sort=relevance&q=analysis")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,".o-teaser__heading")
            for element in elements:
                 data["financial_analysis"].append(element.text)
            data1 = ",".join(data)
            worksheet.update_cell(7,9,data1)
            linkList = []
            contentDiv = driver.find_elements(By.CSS_SELECTOR, '.o-teaser__heading')
            for i in contentDiv:
                items = i.find_elements(By.XPATH, './/a[@href]')
                for item in items:
                    link = item.get_attribute('href')
                    print("item is",link)
                    linkList.append(link)
            data1 = ",".join(linkList)
            worksheet.update_cell(31,9,data1)
            driver.quit()

            #global economic trends
            driver = webdriver.Chrome()
            driver.get("https://www.ft.com/world")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,".js-teaser-heading-link")
            for element in elements:
                 data["global_economic_trends"].append(element.text)
            data["global_economic_trends"]= [article for article in data["global_economic_trends"] if article]  
            data1 = ",".join(data["global_economic_trends"])
            worksheet.update_cell(7,10,data1) 
            linkList = []
            contentDiv = driver.find_elements(By.CSS_SELECTOR, '.o-teaser__heading')
            for i in contentDiv:
                items = i.find_elements(By.XPATH, './/a[@href]')
                for item in items:
                    link = item.get_attribute('href')
                    print("item is",link)
                    linkList.append(link)
            data1 = ",".join(linkList)
            worksheet.update_cell(31,10,data1)
            driver.quit()
        # print("Successfully scraped",name,"\n----------")
        # print("latest news\n",data["latest_news_articles"])
        # print("market reports\n",data["market_reports"])
        # print("financial analysis\n",data["financial_analysis"])
        # print("global economic trends\n",data["global_economic_trends"])
    
    # #parsing marketResearchPortalsUrls
    for name, url in marketResearchPortalsUrls.items():
        print("----------\nScraping",name)
        
        if name == "Statista":
            data = {
            "industry_analysis":[],
            "statistics":[],
            "market_research_reports":[],
            "trends":[]
            }
            #market research reports
            driver = webdriver.Chrome()
            driver.get("https://www.statista.com/studies-and-reports/industries-and-markets")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,".reportResult__content h3")
            for element in elements:
                data["market_research_reports"].append(element.text)
            data["market_research_reports"]= [article for article in data["market_research_reports"] if article]
            data1 = ",".join(data["market_research_reports"])
            worksheet.update_cell(8,11,data1)
            linkList = []
            contentDiv = driver.find_elements(By.CSS_SELECTOR, '.reportResult')
            for item in contentDiv:
                    link = item.get_attribute('href')
                    print("item is",link)
                    linkList.append(link)
            data1 = ",".join(linkList)
            worksheet.update_cell(31,11,data1)
            driver.quit()

            #industry analysis
            driver = webdriver.Chrome()
            driver.get("https://www.statista.com/studies-and-reports/companies-and-products")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,".reportResult__content h3")
            for element in elements:
                data["industry_analysis"].append(element.text)
            data["industry_analysis"]= [article for article in data["industry_analysis"] if article]
            data1 = ",".join(data["industry_analysis"])
            worksheet.update_cell(8,12,data1)
            linkList = []
            contentDiv = driver.find_elements(By.CSS_SELECTOR, '.reportResult')
            for item in contentDiv:
                    link = item.get_attribute('href')
                    print("item is",link)
                    linkList.append(link)
            data1 = ",".join(linkList)
            worksheet.update_cell(31,12,data1)
            driver.quit()

            #statistics
            driver = webdriver.Chrome()
            driver.get("https://www.statista.com/statistics/popular/")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,".statList__title")
            for element in elements:
                data["statistics"].append(element.text)
            data["statistics"]= [article for article in data["statistics"] if article]
            data1 = ",".join(data["statistics"])
            worksheet.update_cell(8,13,data1)
            linkList = []
            contentDiv = driver.find_elements(By.CSS_SELECTOR, '.starList__item')
            for i in contentDiv:
                items = i.find_elements(By.XPATH, './/a[@href]')
                for item in items:
                    link = item.get_attribute('href')
                    print("item is",link)
                    linkList.append(link)
            data1 = ",".join(linkList)
            worksheet.update_cell(31,13,data1)
            driver.quit()

            #trends
            driver = webdriver.Chrome()
            driver.get("https://www.statista.com/studies-and-reports/digital-and-trends")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,".reportResult__content h3")
            for element in elements:
                data["trends"].append(element.text)
            data["trends"]= [article for article in data["trends"] if article]
            data1 = ",".join(data["trends"])
            worksheet.update_cell(8,14,data1)
            linkList = []
            contentDiv = driver.find_elements(By.CSS_SELECTOR, '.reportResult')
            for item in contentDiv:
                    link = item.get_attribute('href')
                    print("item is",link)
                    linkList.append(link)
            data1 = ",".join(linkList)
            worksheet.update_cell(31,14,data1)
            driver.quit()

        elif name == "MarketResearch.com":
            data = {
            "industry_analysis":[],
            "statistics":[],
            "market_research_reports":[],
            "trends":[]
            }
            #market research reports
            driver = webdriver.Chrome()
            driver.get("https://www.marketresearch.com/Marketing-Market-Research-c70/")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,".order-1 a")
            for element in elements:
                data["market_research_reports"].append(element.text)
            data["market_research_reports"]= [article for article in data["market_research_reports"] if article]
            data1 = ",".join(data["market_research_reports"])
            worksheet.update_cell(9,11,data1)
            linkList = []
            contentDiv = driver.find_elements(By.CSS_SELECTOR, '.vhp-card-list')
            for i in contentDiv:
                items = i.find_elements(By.XPATH, './/a[@href]')
                for item in items:
                    link = item.get_attribute('href')
                    print("item is",link)
                    linkList.append(link)
            data1 = worksheet.cell(31,11).value + ",".join(linkList)
            worksheet.update_cell(31,11,data1)

            #industry analysis
            driver = webdriver.Chrome()
            driver.get("https://www.marketresearch.com/Technology-Media-c1599/")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,".order-1 a")
            for element in elements:
                data["industry_analysis"].append(element.text)
            data["industry_analysis"]= [article for article in data["industry_analysis"] if article]
            data1 = ",".join(data["industry_analysis"])
            worksheet.update_cell(9,12,data1)
            linkList = []
            contentDiv = driver.find_elements(By.CSS_SELECTOR, '.vhp-card-list')
            for i in contentDiv:
                items = i.find_elements(By.XPATH, './/a[@href]')
                for item in items:
                    link = item.get_attribute('href')
                    print("item is",link)
                    linkList.append(link)
            data1 = worksheet.cell(31,12).value + ",".join(linkList)
            worksheet.update_cell(31,12,data1)

            #statistics
            driver = webdriver.Chrome()
            driver.get("https://www.marketresearch.com/Marketing-Market-Research-c70/Market-Research-c72/")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,".order-1 a")
            for element in elements:
                data["statistics"].append(element.text)
            data["statistics"]= [article for article in data["statistics"] if article]
            data1 = ",".join(data["statistics"])
            worksheet.update_cell(9,13,data1)
            linkList = []
            contentDiv = driver.find_elements(By.CSS_SELECTOR, '.list-reports li')
            for i in contentDiv:
                items = i.find_elements(By.XPATH, './/a[@href]')
                for item in items:
                    link = item.get_attribute('href')
                    print("item is",link)
                    linkList.append(link)
            data1 = ",".join(linkList)
            worksheet.update_cell(31,13,data1)
            

            #trends
            driver = webdriver.Chrome()
            driver.get(url)
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,".featureDetail")
            for element in elements:
                data["trends"].append(element.text)
            data["trends"]= [article for article in data["trends"] if article]
            data1 = ",".join(data["trends"])
            worksheet.update_cell(9,14,data1)
            linkList = []
            contentDiv = driver.find_elements(By.CSS_SELECTOR, '.featureItem')
            for i in contentDiv:
                items = i.find_elements(By.XPATH, './/a[@href]')
                for item in items:
                    link = item.get_attribute('href')
                    print("item is",link)
                    linkList.append(link)
            data1 = worksheet.cell(31,14).value + ",".join(linkList)
            worksheet.update_cell(31,14,data1)

        elif name == "IBISWorld":
            data = {
            "industry_analysis":[],
            "statistics":[],
            "market_research_reports":[],
            "trends":[]
            }
            #market research reports
            driver = webdriver.Chrome()
            driver.get("https://www.ibisworld.com/market-research-reports/#global")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,".country-reports span")
            for element in elements:
                data["market_research_reports"].append(element.text)
            data["market_research_reports"]= [article for article in data["market_research_reports"] if article]
            data1 = ",".join(data["market_research_reports"])
            worksheet.update_cell(10,11,data1)
            linkList = []
            contentDiv = driver.find_elements(By.CSS_SELECTOR, '#usIndustryList')
            for i in contentDiv:
                items = i.find_elements(By.XPATH, './/a[@href]')
                for item in items:
                    link = item.get_attribute('href')
                    print("item is",link)
                    linkList.append(link)
            data1 = worksheet.cell(31,11).value + ",".join(linkList)
            worksheet.update_cell(31,11,data1)

            #industry analysis
            driver = webdriver.Chrome()
            driver.get("https://www.ibisworld.com/united-states/list-of-industries/#specialized-reports")
            elements = driver.find_elements(By.CSS_SELECTOR,"#sectorsIndustryList a")
            for item in elements:
                data["industry_analysis"].append(item.text)
            data1 = ",".join(data["industry_analysis"])
            worksheet.update_cell(10,12,data1)
            linkList = []
            contentDiv = driver.find_elements(By.CSS_SELECTOR, '.Sector')
            # total_chars = len(worksheet.cell(31,12).value)
            # char_limit = 49999
            # for i in contentDiv:
            #     items = i.find_elements(By.XPATH, './/a[@href]')
            #     for item in items:
            #         link = item.get_attribute('href')
            #         item_length = len(link)
            #         print(link)
            #         if total_chars + item_length + 1 <= char_limit:
            #             linkList.append(link)
            #             total_chars += item_length
            #         else:
            #             break
            # data1 = worksheet.cell(31,12).value + ",".join(linkList)
            # worksheet.update_cell(31,12,data1)
            # weblink data entered manually. 

            #statistics
            driver = webdriver.Chrome()
            driver.get("https://www.ibisworld.com/search/default.aspx?st=statistics")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,".hit-name span")
            total_chars = len(worksheet.cell(31,13).value)
            char_limit = 49999
            for item in elements:
                item_length = len(item.text)
                if total_chars + item_length <= char_limit:
                    data["statistics"].append(item.text)
                    total_chars += item_length
                else:
                    break
            data1 = ",".join(data["statistics"])
            worksheet.update_cell(10,13,data1)
            linkList = []
            contentDiv = driver.find_elements(By.CSS_SELECTOR, '.ais-Hits-item')
            for i in contentDiv:
                items = i.find_elements(By.XPATH, './/a[@href]')
                for item in items:
                    link = item.get_attribute('href')
                    print(link)
                    linkList.append(link)
            data1 = worksheet.cell(31,13).value + ",".join(linkList)
            worksheet.update_cell(31,13,data1)

            #trends
            driver = webdriver.Chrome()
            driver.get("https://www.ibisworld.com/global/list-of-industries/")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,"#sectorsIndustryList")
            total_chars = len(worksheet.cell(31,14).value)
            char_limit = 49999
            for item in elements:
                item_length = len(item.text)
                if total_chars + item_length <= char_limit:
                    data["trends"].append(item.text)
                    total_chars += item_length
                else:
                    break
            data1 = ",".join(data["trends"])
            worksheet.update_cell(10,14,data1)
            linkList = []
            contentDiv = driver.find_elements(By.CSS_SELECTOR, '.Sector')
            for i in contentDiv:
                items = i.find_elements(By.XPATH, './/a[@href]')
                for item in items:
                    link = item.get_attribute('href')
                    print("item is",link)
                    linkList.append(link)
            data1 = worksheet.cell(31,14).value + ",".join(linkList)
            worksheet.update_cell(31,14,data1)

        # print("market research reports\n",data["market_research_reports"])
        # print("industry analysis\n",data["industry_analysis"])
        # print("statistics\n",data["statistics"])
        # print("trends\n",data["trends"])

    for name,url in globalNewsOutletsUrls.items():
        if name == "The Economic Times":
            data = {
                "global_news":[]
            }
            driver = webdriver.Chrome()
            driver.get("https://economictimes.indiatimes.com/news/economy")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,".top-news li")
            for element in elements:
                data["global_news"].append(element.text)
            data["global_news"]= [article for article in data["global_news"] if article]
            data1 = ",".join(data["global_news"])
            worksheet.update_cell(11,15,data1)
            linkList = []
            contentDiv = driver.find_elements(By.CSS_SELECTOR, '.top-news')
            for i in contentDiv:
                items = i.find_elements(By.XPATH, './/a[@href]')
                for item in items:
                    link = item.get_attribute('href')
                    print(link)
                    linkList.append(link)
            data1 = ",".join(linkList)
            worksheet.update_cell(31,15,data1)
        elif name == "Business Standard":
            data = {
                "global_news":[]
            }
            driver = webdriver.Chrome()
            driver.get("https://www.business-standard.com/")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,".cardlist p")
            for element in elements:
                data["global_news"].append(element.text)
            data["global_news"]= [article for article in data["global_news"] if article]
            data1 = ",".join(data["global_news"])
            worksheet.update_cell(12,15,data1)
            linkList = []
            contentDiv = driver.find_elements(By.CSS_SELECTOR, '.cardlist')
            for i in contentDiv:
                items = i.find_elements(By.XPATH, './/a[@href]')
                for item in items:
                    link = item.get_attribute('href')
                    print(link)
                    linkList.append(link)
            data1 = worksheet.cell(31,15).value + ",".join(linkList)
            worksheet.update_cell(31,15,data1)
        elif name == "Hindustan Times":
            data = {
                "global_news":[]
            }
            driver = webdriver.Chrome()
            driver.get("https://www.hindustantimes.com")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,".cartHolder a")
            for element in elements:
                data["global_news"].append(element.text)
            data["global_news"]= [article for article in data["global_news"] if article]
            data1 = ",".join(data["global_news"])
            worksheet.update_cell(13,15,data1)
            linkList = []
            contentDiv = driver.find_elements(By.CSS_SELECTOR, '.cartHolder')
            for i in contentDiv:
                items = i.find_elements(By.XPATH, './/a[@href]')
                for item in items:
                    link = item.get_attribute('href')
                    print(link)
                    linkList.append(link)
            data1 = worksheet.cell(31,15).value + ",".join(linkList)
            worksheet.update_cell(31,15,data1)
        elif name == "China Daily":
            data = {
                "global_news":[]
            }
            driver = webdriver.Chrome()
            driver.get("https://www.chinadaily.com.cn/")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,".txt1")
            for element in elements:
                data["global_news"].append(element.text)
            data["global_news"]= [article for article in data["global_news"] if article]
            data1 = ",".join(data["global_news"])
            worksheet.update_cell(14,15,data1)
            linkList = []
            contentDiv = driver.find_elements(By.CSS_SELECTOR, '.txtBox')
            for i in contentDiv:
                items = i.find_elements(By.XPATH, './/a[@href]')
                for item in items:
                    link = item.get_attribute('href')
                    print(link)
                    linkList.append(link)
            data1 = worksheet.cell(31,15).value + ",".join(linkList)
            worksheet.update_cell(31,15,data1)
        elif name == "Global Times":
            data = {
                "global_news":[]
            }
            driver = webdriver.Chrome()
            driver.get("https://www.globaltimes.cn")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,".row p")
            for element in elements:
                data["global_news"].append(element.text)
            data["global_news"]= [article for article in data["global_news"] if article]
            data1 = ",".join(data["global_news"])
            worksheet.update_cell(15,15,data1)
            linkList = []
            contentDiv = driver.find_elements(By.CSS_SELECTOR, '.new_title_m')
            for i in contentDiv:
                items = i.find_elements(By.XPATH, './/a[@href]')
                for item in items:
                    link = item.get_attribute('href')
                    print(link)
                    linkList.append(link)
            data1 = worksheet.cell(31,15).value + ",".join(linkList)
            worksheet.update_cell(31,15,data1)
        elif name == "Caixin Global":
            data = {
                "global_news":[]
            }
            driver = webdriver.Chrome()
            driver.get("https://www.caixinglobal.com")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,"h2 a")
            for element in elements:
                data["global_news"].append(element.text)
            data["global_news"]= [article for article in data["global_news"] if article]
            data1 = ",".join(data["global_news"])
            worksheet.update_cell(16,15,data1)
            linkList = []
            contentDiv = driver.find_elements(By.CSS_SELECTOR, 'h1')
            for i in contentDiv:
                items = i.find_elements(By.XPATH, './/a[@href]')
                for item in items:
                    link = item.get_attribute('href')
                    print(link)
                    linkList.append(link)
            data1 = worksheet.cell(31,15).value + ",".join(linkList)
            worksheet.update_cell(31,15,data1)
        elif name == "The Japan Times":
            data = {
                "global_news":[]
            }
            driver = webdriver.Chrome()
            driver.get("https://www.japantimes.co.jp")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,".article-title a")
            for element in elements:
                data["global_news"].append(element.text)
            data["global_news"]= [article for article in data["global_news"] if article]
            data1 = ",".join(data["global_news"])
            worksheet.update_cell(17,15,data1)
            linkList = []
            contentDiv = driver.find_elements(By.CSS_SELECTOR, '.article-title')
            for i in contentDiv:
                items = i.find_elements(By.XPATH, './/a[@href]')
                for item in items:
                    link = item.get_attribute('href')
                    print(link)
                    linkList.append(link)
            data1 = worksheet.cell(31,15).value + ",".join(linkList)
            worksheet.update_cell(31,15,data1)
        elif name == "Nikkei Asia":
            data = {
                "global_news":[]
            }
            driver = webdriver.Chrome()
            driver.get("https://asia.nikkei.com")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,"h4 span")
            for element in elements:
                data["global_news"].append(element.text)
            data["global_news"]= [article for article in data["global_news"] if article]
            data1 = ",".join(data["global_news"])
            worksheet.update_cell(18,15,data1)
            linkList = []
            contentDiv = driver.find_elements(By.CSS_SELECTOR, '.article-title')
            for i in contentDiv:
                items = i.find_elements(By.XPATH, './/a[@href]')
                for item in items:
                    link = item.get_attribute('href')
                    print(link)
                    linkList.append(link)
            data1 = worksheet.cell(31,15).value + ",".join(linkList)
            worksheet.update_cell(31,15,data1)
        elif name == "NHK World":
            data = {
                "global_news":[]
            }
            driver = webdriver.Chrome()
            driver.get("https://www.nhk.or.jp/nhkworld")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,"h3 a")
            for element in elements:
                data["global_news"].append(element.text)
            data["global_news"]= [article for article in data["global_news"] if article]
            data1 = ",".join(data["global_news"])
            worksheet.update_cell(19,15,data1)
            linkList = []
            contentDiv = driver.find_elements(By.CSS_SELECTOR, '.tNewsListItem__title')
            for i in contentDiv:
                items = i.find_elements(By.XPATH, './/a[@href]')
                for item in items:
                    link = item.get_attribute('href')
                    print(link)
                    linkList.append(link)
            data1 = worksheet.cell(31,15).value + ",".join(linkList)
            worksheet.update_cell(31,15,data1)
        elif name == "The Straits Times":
            data = {
                "global_news":[]
            }
            driver = webdriver.Chrome()
            driver.get("https://straitstimes.com")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,"h5 a")
            for element in elements:
                data["global_news"].append(element.text)
            data["global_news"]= [article for article in data["global_news"] if article]
            data1 = ",".join(data["global_news"])
            worksheet.update_cell(20,15,data1)
            linkList = []
            contentDiv = driver.find_elements(By.CSS_SELECTOR, '.card-title')
            for i in contentDiv:
                items = i.find_elements(By.XPATH, './/a[@href]')
                for item in items:
                    link = item.get_attribute('href')
                    print(link)
                    linkList.append(link)
            data1 = worksheet.cell(31,15).value + ",".join(linkList)
            worksheet.update_cell(31,15,data1)
        elif name == "Channel News Asia":
            data = {
                "global_news":[]
            }
            driver = webdriver.Chrome()
            driver.get("https://www.channelnewsasia.com")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,"h6 a")
            for element in elements:
                data["global_news"].append(element.text)
            data["global_news"]= [article for article in data["global_news"] if article]
            data1 = ",".join(data["global_news"])
            worksheet.update_cell(21,15,data1)
            linkList = []
            contentDiv = driver.find_elements(By.CSS_SELECTOR, '.list-object__heading')
            for i in contentDiv:
                items = i.find_elements(By.XPATH, './/a[@href]')
                for item in items:
                    link = item.get_attribute('href')
                    print(link)
                    linkList.append(link)
            data1 = ",".join(linkList)
            worksheet.update_cell(32,15,data1)
        elif name == "The Business Times":
            data = {
                "global_news":[]
            }
            driver = webdriver.Chrome()
            driver.get("https://www.businesstimes.com.sg")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,".inherit .word-break")
            for element in elements:
                data["global_news"].append(element.text)
            data["global_news"]= [article for article in data["global_news"] if article]
            data1 = ",".join(data["global_news"])
            worksheet.update_cell(22,15,data1)
            linkList = []
            contentDiv = driver.find_elements(By.CSS_SELECTOR, '.my-1 _title_d4ajx_12')
            for i in contentDiv:
                items = i.find_elements(By.XPATH, './/a[@href]')
                for item in items:
                    link = item.get_attribute('href')
                    print(link)
                    linkList.append(link)
            data1 = worksheet.cell(32,15).value + ",".join(linkList)
            worksheet.update_cell(32,15,data1)
        elif name == "The Korea Herald":
            data = {
                "global_news":[]
            }
            driver = webdriver.Chrome()
            driver.get("https://www.koreaherald.com")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,".news_txt")
            for element in elements:
                data["global_news"].append(element.text)
            data["global_news"]= [article for article in data["global_news"] if article]
            data1 = ",".join(data["global_news"])
            worksheet.update_cell(23,15,data1)
            linkList = []
            contentDiv = driver.find_elements(By.CSS_SELECTOR, '.news_txt')
            for i in contentDiv:
                items = i.find_elements(By.XPATH, './/a[@href]')
                for item in items:
                    link = item.get_attribute('href')
                    print(link)
                    linkList.append(link)
            data1 = worksheet.cell(32,15).value + ",".join(linkList)
            worksheet.update_cell(32,15,data1)
        elif name == "The Korea Times":
            data = {
                "global_news":[]
            }
            driver = webdriver.Chrome()
            driver.get("https://www.koreatimes.co.kr/www/section_129.html")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,".LoraMedium")
            for element in elements:
                data["global_news"].append(element.text)
            data["global_news"]= [article for article in data["global_news"] if article]
            data1 = ",".join(data["global_news"])
            worksheet.update_cell(24,15,data1)
            linkList = []
            contentDiv = driver.find_elements(By.CSS_SELECTOR, '.list2_article_headline_top2')
            for i in contentDiv:
                items = i.find_elements(By.XPATH, './/a[@href]')
                for item in items:
                    link = item.get_attribute('href')
                    print(link)
                    linkList.append(link)
            data1 = worksheet.cell(32,15).value + ",".join(linkList)
            worksheet.update_cell(32,15,data1)
        elif name == "Gulf News":
            data = {
                "global_news":[]
            }
            driver = webdriver.Chrome()
            driver.get("https://www.gulfnews.com")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,"h2 a")
            for element in elements:
                data["global_news"].append(element.text)
            data["global_news"]= [article for article in data["global_news"] if article]
            data1 = ",".join(data["global_news"])
            worksheet.update_cell(25,15,data1)
            linkList = []
            contentDiv = driver.find_elements(By.CSS_SELECTOR, '.card-title')
            for i in contentDiv:
                items = i.find_elements(By.XPATH, './/a[@href]')
                for item in items:
                    link = item.get_attribute('href')
                    print(link)
                    linkList.append(link)
            data1 = worksheet.cell(32,15).value + ",".join(linkList)
            worksheet.update_cell(32,15,data1)
        elif name == "Khaleej Times":
            data = {
                "global_news":[]
            }
            driver = webdriver.Chrome()
            driver.get("https://www.khaleejtimes.com")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,"h2 a")
            for element in elements:
                data["global_news"].append(element.text)
            data["global_news"]= [article for article in data["global_news"] if article]
            data1 = ",".join(data["global_news"])
            worksheet.update_cell(26,15,data1)
            linkList = []
            contentDiv = driver.find_elements(By.CSS_SELECTOR, '.post-title')
            for i in contentDiv:
                items = i.find_elements(By.XPATH, './/a[@href]')
                for item in items:
                    link = item.get_attribute('href')
                    print(link)
                    linkList.append(link)
            data1 = worksheet.cell(32,15).value + ",".join(linkList)
            worksheet.update_cell(32,15,data1)
        elif name == "The National":
            data = {
                "global_news":[]
            }
            driver = webdriver.Chrome()
            driver.get("https://www.thenationalnews.com")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,".order-0 p")
            for element in elements:
                data["global_news"].append(element.text)
            data["global_news"]= [article for article in data["global_news"] if article]
            data1 = ",".join(data["global_news"])
            worksheet.update_cell(27,15,data1)
            linkList = []
            contentDiv = driver.find_elements(By.CSS_SELECTOR, '.defaultstyled__StyledCardDetails-sc-1dedoj7-1 erznAs')
            for i in contentDiv:
                items = i.find_elements(By.XPATH, './/a[@href]')
                for item in items:
                    link = item.get_attribute('href')
                    print(link)
                    linkList.append(link)
            data1 = worksheet.cell(32,15).value + ",".join(linkList)
            worksheet.update_cell(32,15,data1)
        elif name == "Taipei Times":
            data = {
                "global_news":[]
            }
            driver = webdriver.Chrome()
            driver.get("https://www.taipeitimes.com")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,".bf2")
            for element in elements:
                data["global_news"].append(element.text)
            data["global_news"]= [article for article in data["global_news"] if article]
            data1 = ",".join(data["global_news"])
            worksheet.update_cell(28,15,data1)
            linkList = []
            contentDiv = driver.find_elements(By.CSS_SELECTOR, '.boxTitle')
            for i in contentDiv:
                items = i.find_elements(By.XPATH, './/a[@href]')
                for item in items:
                    link = item.get_attribute('href')
                    print(link)
                    linkList.append(link)
            data1 = worksheet.cell(32,15).value + ",".join(linkList)
            worksheet.update_cell(32,15,data1)
        elif name == "Focus Taiwan":
            data = {
                "global_news":[]
            }
            driver = webdriver.Chrome()
            driver.get("https://focustaiwan.tw/business")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,".h2 span")
            for element in elements:
                data["global_news"].append(element.text)
            data["global_news"]= [article for article in data["global_news"] if article]
            data1 = ",".join(data["global_news"])
            worksheet.update_cell(29,15,data1)
            linkList = []
            contentDiv = driver.find_elements(By.CSS_SELECTOR, '.listStyle')
            for i in contentDiv:
                items = i.find_elements(By.XPATH, './/a[@href]')
                for item in items:
                    link = item.get_attribute('href')
                    print(link)
                    linkList.append(link)
            data1 = worksheet.cell(32,15).value + ",".join(linkList)
            worksheet.update_cell(32,15,data1)
        elif name == "Taiwan News":
            data = {
                "global_news":[]
            }
            driver = webdriver.Chrome()
            driver.get("https://www.taiwannews.com.tw/category/Business")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,".w-full p")
            for element in elements:
                data["global_news"].append(element.text)
            data["global_news"]= [article for article in data["global_news"] if article]
            data1 = ",".join(data["global_news"])
            worksheet.update_cell(30,15,data1)
            linkList = []
            contentDiv = driver.find_elements(By.CSS_SELECTOR, '.listStyle')
            for i in contentDiv:
                items = i.find_elements(By.XPATH, './/a[@href]')
                for item in items:
                    link = item.get_attribute('href')
                    print(link)
                    linkList.append(link)
            data1 = worksheet.cell(32,15).value + ",".join(linkList)
            worksheet.update_cell(32,15,data1)
    # print(data["global_news"])
             
    all_values = worksheet.get_all_values()
    def count_commas(s):
        return s.count(',')

    comma_counts_per_row = []
    for row in all_values:
        comma_count = sum(count_commas(cell) for cell in row)
        comma_counts_per_row.append(comma_count)

    for i, count in enumerate(comma_counts_per_row, start=1):
        print(f'Row {i} has {count} commas.')

    if len(all_values) >= 31:
        row_31 = all_values[30]  # Row 31 in 0-indexed list
        comma_counts_per_column_row_31 = [count_commas(cell) for cell in row_31]

        # Print comma counts for each column in row 31
        for i, count in enumerate(comma_counts_per_column_row_31, start=1):
            print(f'Column {i} in row 31 has {count} entries.')
    else:
        print('The sheet does not have 31 rows.')

    ibisArticlesStats = count_commas(worksheet.cell(33, 12).value)
    print("ibisArticleStats has entries:", ibisArticlesStats)
    

    
if __name__ == "__main__":
    main()
