import requests
from bs4 import BeautifulSoup
import chardet
import io
import gzip
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import csv


import time

headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
    'Accept-Language': 'en-US,en;q=0.5',
    'Accept-Encoding': 'gzip, deflate, br',
    'Connection': 'keep-alive',
    'Referer': 'https://www.google.com/'
}

# Lists of websites to scrape

competitorWebsitesUrls = {
    "Boston Consulting Group": "https://www.bcg.com",
    "McKinsey & Company": "https://www.mckinsey.com",
    "Deloitte": "https://www.deloitte.com"
}

industryNewsReportsUrls = {
    "Bloomberg": "https://www.bloomberg.com",
    "Reuters": "https://www.reuters.com", 
    "Financial Times": "https://www.ft.com"
}

marketResearchPortalsUrls = {
    "Statista": "https://www.statista.com",
    "MarketResearch.com": "https://www.marketresearch.com",
    "IBISWorld": "https://www.ibisworld.com"
}

globalNewsOutletsUrls = {
    # India:
    "The Economic Times": "https://economictimes.indiatimes.com",
    "Business Standard": "https://www.business-standard.com", 
    "Hindustan Times": "https://www.hindustantimes.com", 

    # China:
    "China Daily": "https://www.chinadaily.com.cn",
    "Global Times": "https://www.globaltimes.cn",
    "Caixin Global": "https://www.caixinglobal.com",

    # Japan:
    "The Japan Times": "https://www.japantimes.co.jp",
    "Nikkei Asia": "https://asia.nikkei.com",
    "NHK World": "https://www.nhk.or.jp/nhkworld",

    # Singapore:
    "The Straits Times": "https://straitstimes.com", 
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
    print("Detected encoding for ",url,": ",detected_encoding)

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
                print("Successfully connected to", url)
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
    # data structure for extracted data
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
            # market insights
            insights = soup.find_all('h3',class_="article-block__title")
            if insights == []:
                insights = scrapeDynamic("https://www.bcg.com/publications", "article-block__title")
                insights = [insight for insight in insights if insight !=""]
                for i in insights:
                    data["market_insights"].append(i)
            else:
                for i in insights:
                    data["market_insights"].append(i.get_text())

            # case studies
            cases = soup.find_all('h1',class_="homepageAnimatedHero-title")
            if cases == []:
                cases = scrapeDynamic(url, "homepageAnimatedHero-title" )
            for case in cases:
                data["case_studies"].append(case.text)
            
            # service offerings
            services = soup.find_all('li',class_="glide__slide")
            for service in services:
                 data["service_offerings"].append(service.get_text().strip())

            #client testimonials
            driver = webdriver.Chrome()
            driver.get("https://www.bcg.com/capabilities/digital-technology-data/client-success")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR, '.article-block__title h3')
            for element in elements:
                data["client_testimonials"].append(element.text)
            driver.quit()

            #thought leadership articles
            articles = scrapeDynamic(url,"HomepageCardsCarouselPromo__title")
            articles = [article for article in articles if article ]
            for a in articles:
                    data["thought_leadership"].append(a)
        elif name == "McKinsey & Company":
            # market insights
            data['market_insights'].append([div.get_text() for div in soup.select(".HighlightBar_mck-c-highlight-bar__item-title__g9Fdc span")])

            # case studies
            cases = soup.find_all('h1',class_="mck-c-hwpm__heading--text-m")
            for case in cases:
                data["case_studies"].append(case.get_text())
            
            #service offering
            driver = webdriver.Chrome()
            driver.get(url)
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR, '[data-testid="hamburger-list-Capabilities"] span')
            for element in elements:
                data["service_offerings"].append(element.text)
            driver.quit()

            #thought leadership articles
            articles = soup.find_all("h3",class_="mdc-c-heading___0fM1W_d0bf44e")
            for article in articles:
                data["thought_leadership"].append(article.get_text())
            
            #client testimonials
            soup = scrapeWebsite("https://www.mckinsey.com/capabilities/operations/how-we-help-clients/operations-excellence-program")
            ts = soup.select('.mdc-c-heading___0fM1W_d0bf44e[data-component="mdc-c-heading"]')
            uncleaned = []
            for t in ts:
                uncleaned.append(t.get_text())
            cleaned_list = [
                item for item in uncleaned 
                if item.strip() and "join" not in item.lower() and "testimonial" not in item.lower()
            ]
            for item in cleaned_list:
                 data["client_testimonials"].append(item)

        elif name == "Deloitte":
            # market insights
            driver = webdriver.Chrome()
            driver.get(url)
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR, '.trending-item h4')
            for element in elements:
                data["market_insights"].append(element.text)
           
            # case studies
            elements = driver.find_elements(By.CSS_SELECTOR, '.promo-focus h2')
            for element in elements:
                data["case_studies"].append(element.text)
            
            #service offerings
            elements = driver.find_elements(By.CSS_SELECTOR,"#footer__links-services a")
            for element in elements:
                data["service_offerings"].append(element.text)
            
            #thought leadership 
            elements = driver.find_elements(By.CSS_SELECTOR,".showcase-content-wrap h1")
            for element in elements:
                data["thought_leadership"].append(element.text)
            driver.quit()

            #client testimonial
            driver = webdriver.Chrome()
            driver.get("https://www2.deloitte.com/us/en/pages/about-deloitte/articles/client-stories-and-successes.html?icid=bottom_client-stories-and-successes")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,".glass__effect h2")
            for element in elements:
                 data["client_testimonials"].append(element.text)
            driver.quit()
        print("Successfully scraped",name,"\n----------")
      
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
            # latest news
            driver = webdriver.Chrome()
            driver.get(url)
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,".StoryListLatestFiltered_storyBlock___szNS a")
            for element in elements:
                    data["latest_news_articles"].append(element.text)
            driver.quit()
          
            # market reports
            driver = webdriver.Chrome()
            driver.get("https://www.bloomberg.com/markets")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,".styles_storyBlock__l5VzV a")
            for element in elements:
                 data["market_reports"].append(element.text)
            driver.quit()

            # financial analysis
            driver = webdriver.Chrome()
            driver.get("https://www.bloomberg.com/search?query=financial%20markets")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,".headline__3a97424275 a")
            for element in elements:
                 data["market_reports"].append(element.text)
            driver.quit()
        
            # global economic trends
            driver = webdriver.Chrome()
            driver.get("https://www.bloomberg.com/economics")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,".styles_storyBlock__l5VzV a")
            for element in elements:
                 data["market_reports"].append(element.text)
            driver.quit()

        elif name == "Reuters":
            # latest news
            driver = webdriver.Chrome()
            driver.get(url)
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,".basic-card__container__1y8wi span")
            for element in elements:
                 data["latest_news_articles"].append(element.text)
            data["latest_news_articles"] = [article for article in data["latest_news_articles"] if article ]
            driver.quit()

        if name == "Financial Times":
            #latest news
            driver = webdriver.Chrome()
            driver.get(url)
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,".story-group-stacked__primary-story span")
            for element in elements:
                 data["latest_news_articles"].append(element.text)
            data["latest_news_articles"] = [article for article in data["latest_news_articles"] if article]
            driver.quit()

            #market reports
            driver = webdriver.Chrome()
            driver.get(url)
            html_content = driver.page_source
            reports = driver.find_elements(By.CSS_SELECTOR,".js-teaser-headline span")
            for report in reports:
                 data["market_reports"].append(report.text)
            data["market_reports"] = [article for article in data["market_reports"] if article]
            data["market_reports"] = [
                item for item in data["market_reports"] 
                if item.strip() and "opinion" not in item.lower() and "people" not in item.lower() and "FT" not in item.lower()
            ]
            driver.quit()

            #financial analysis
            driver = webdriver.Chrome()
            driver.get("https://markets.ft.com/data")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,".mod-news span")
            for element in elements:
                 data["financial_analysis"].append(element.text)
            data["financial_analysis"] = [article for article in data["financial_analysis"] if article]
            driver.quit()

            #global economic trends
            driver = webdriver.Chrome()
            driver.get("https://www.ft.com/world")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,".o-teaser__heading a")
            for element in elements:
                 data["global_economic_trends"].append(element.text)
            data["global_economic_trends"]= [article for article in data["global_economic_trends"] if article]
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
            #market research reports
            driver = webdriver.Chrome()
            driver.get("https://www.statista.com/studies-and-reports/industries-and-markets")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,".reportResult__content h3")
            for element in elements:
                data["market_research_reports"].append(element.text)
            data["market_research_reports"]= [article for article in data["market_research_reports"] if article]
            driver.quit() 

            #industry analysis
            driver = webdriver.Chrome()
            driver.get("https://www.statista.com/studies-and-reports/companies-and-products")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,".reportResult__content h3")
            for element in elements:
                data["industry_analysis"].append(element.text)
            data["industry_analysis"]= [article for article in data["industry_analysis"] if article]
            driver.quit() 

            #statistics
            driver = webdriver.Chrome()
            driver.get("https://www.statista.com/statistics/popular/")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,".statList__title")
            for element in elements:
                data["statistics"].append(element.text)
            data["statistics"]= [article for article in data["statistics"] if article]
            driver.quit() 

            #trends
            driver = webdriver.Chrome()
            driver.get("https://www.statista.com/studies-and-reports/digital-and-trends")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,".reportResult__content h3")
            for element in elements:
                data["trends"].append(element.text)
            data["trends"]= [article for article in data["trends"] if article]
            driver.quit() 

        elif name == "MarketResearch.com":
            #market research reports
            driver = webdriver.Chrome()
            driver.get("https://www.marketresearch.com/Marketing-Market-Research-c70/")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,".order-1 a")
            for element in elements:
                data["market_research_reports"].append(element.text)
            data["market_research_reports"]= [article for article in data["market_research_reports"] if article]
            driver.quit() 

            #industry analysis
            driver = webdriver.Chrome()
            driver.get("https://www.marketresearch.com/Technology-Media-c1599/")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,".order-1 a")
            for element in elements:
                data["industry_analysis"].append(element.text)
            data["industry_analysis"]= [article for article in data["industry_analysis"] if article]
            driver.quit() 

            #statistics
            driver = webdriver.Chrome()
            driver.get("https://www.marketresearch.com/Marketing-Market-Research-c70/Market-Research-c72/")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,".order-1 a")
            for element in elements:
                data["statistics"].append(element.text)
            data["statistics"]= [article for article in data["statistics"] if article]
            driver.quit() 

            #trends
            driver = webdriver.Chrome()
            driver.get(url)
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,".featureDetail")
            for element in elements:
                data["trends"].append(element.text)
            data["trends"]= [article for article in data["trends"] if article]
            driver.quit() 

        elif name == "IBISWorld":
            #market research reports
            driver = webdriver.Chrome()
            driver.get("https://www.ibisworld.com/market-research-reports/#global")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,".country-reports span")
            for element in elements:
                data["market_research_reports"].append(element.text)
            data["market_research_reports"]= [article for article in data["market_research_reports"] if article]
            driver.quit() 

            #industry analysis
            driver = webdriver.Chrome()
            driver.get("https://www.ibisworld.com/united-states/list-of-industries/#specialized-reports")
            elements = driver.find_elements(By.CSS_SELECTOR,"#sectorsIndustryList a")
            for element in elements:
                data["industry_analysis"].append(element.text)
            driver.quit() 

            #statistics
            driver = webdriver.Chrome()
            driver.get("https://www.ibisworld.com/search/default.aspx?st=statistics")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,".hit-name span")
            for element in elements:
                data["statistics"].append(element.text)
            data["statistics"]= [article for article in data["statistics"] if article]
            driver.quit() 

            #trends
            driver = webdriver.Chrome()
            driver.get("https://www.ibisworld.com/global/list-of-industries/")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,"#sectorsIndustryList")
            for element in elements:
                data["trends"].append(element.text)
            data["trends"]= [article for article in data["trends"] if article]
            driver.quit() 

        print("market research reports\n",data["market_research_reports"])
        print("industry analysis\n",data["industry_analysis"])
        print("statistics\n",data["statistics"])
        print("trends\n",data["trends"])

    for name,url in globalNewsOutletsUrls.items():
        if name == "The Economic Times":
            driver = webdriver.Chrome()
            driver.get("https://economictimes.indiatimes.com/news/economy")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,".top-news li")
            for element in elements:
                data["global_news"].append(element.text)
            data["global_news"]= [article for article in data["global_news"] if article]
            driver.quit() 
        elif name == "Business Standard":
            driver = webdriver.Chrome()
            driver.get("https://www.business-standard.com/")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,".cardlist p")
            for element in elements:
                data["global_news"].append(element.text)
            data["global_news"]= [article for article in data["global_news"] if article]
            driver.quit() 
        elif name == "Hindustan Times":
            driver = webdriver.Chrome()
            driver.get("https://www.hindustantimes.com")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,".cartHolder a")
            for element in elements:
                data["global_news"].append(element.text)
            data["global_news"]= [article for article in data["global_news"] if article]
            driver.quit() 
        elif name == "China Daily":
            driver = webdriver.Chrome()
            driver.get("https://www.chinadaily.com.cn/")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,".txt1")
            for element in elements:
                data["global_news"].append(element.text)
            data["global_news"]= [article for article in data["global_news"] if article]
            driver.quit() 
        elif name == "Global Times":
            driver = webdriver.Chrome()
            driver.get("https://www.globaltimes.cn")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,".row p")
            for element in elements:
                data["global_news"].append(element.text)
            data["global_news"]= [article for article in data["global_news"] if article]
            driver.quit()
        elif name == "Caixin Global":
            driver = webdriver.Chrome()
            driver.get("https://www.caixinglobal.com")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,"h2 a")
            for element in elements:
                data["global_news"].append(element.text)
            data["global_news"]= [article for article in data["global_news"] if article]
            driver.quit()
        elif name == "The Japan Times":
            driver = webdriver.Chrome()
            driver.get("https://www.japantimes.co.jp")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,".article-title a")
            for element in elements:
                data["global_news"].append(element.text)
            data["global_news"]= [article for article in data["global_news"] if article]
            driver.quit()
        elif name == "Nikkei Asia":
            driver = webdriver.Chrome()
            driver.get("https://asia.nikkei.com")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,"h4 span")
            for element in elements:
                data["global_news"].append(element.text)
            data["global_news"]= [article for article in data["global_news"] if article]
            driver.quit()
        elif name == "NHK World":
            driver = webdriver.Chrome()
            driver.get("https://www.nhk.or.jp/nhkworld")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,"h3 a")
            for element in elements:
                data["global_news"].append(element.text)
            data["global_news"]= [article for article in data["global_news"] if article]
            driver.quit()
        elif name == "The Straits Times":
            driver = webdriver.Chrome()
            driver.get("https://straitstimes.com")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,"h5 a")
            for element in elements:
                data["global_news"].append(element.text)
            data["global_news"]= [article for article in data["global_news"] if article]
            driver.quit()
        elif name == "Channel News Asia":
            driver = webdriver.Chrome()
            driver.get("https://www.channelnewsasia.com")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,"h6 a")
            for element in elements:
                data["global_news"].append(element.text)
            data["global_news"]= [article for article in data["global_news"] if article]
            driver.quit()
        elif name == "The Business Times":
            driver = webdriver.Chrome()
            driver.get("https://www.businesstimes.com.sg")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,".inherit .word-break")
            for element in elements:
                data["global_news"].append(element.text)
            data["global_news"]= [article for article in data["global_news"] if article]
            driver.quit()
        elif name == "The Korea Herald":
            driver = webdriver.Chrome()
            driver.get("https://www.koreaherald.com")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,".ellipsis2")
            for element in elements:
                data["global_news"].append(element.text)
            data["global_news"]= [article for article in data["global_news"] if article]
            driver.quit()
        elif name == "The Korea Times":
            driver = webdriver.Chrome()
            driver.get("https://www.koreatimes.co.kr/www/section_129.html")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,".LoraMedium")
            for element in elements:
                data["global_news"].append(element.text)
            data["global_news"]= [article for article in data["global_news"] if article]
            driver.quit()
        elif name == "Gulf News":
            driver = webdriver.Chrome()
            driver.get("https://www.gulfnews.com")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,"h2 a")
            for element in elements:
                data["global_news"].append(element.text)
            data["global_news"]= [article for article in data["global_news"] if article]
            driver.quit()
        elif name == "Khaleej Times":
            driver = webdriver.Chrome()
            driver.get("https://www.khaleejtimes.com")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,"h2 a")
            for element in elements:
                data["global_news"].append(element.text)
            data["global_news"]= [article for article in data["global_news"] if article]
            driver.quit()
        elif name == "The National":
            driver = webdriver.Chrome()
            driver.get("https://www.thenationalnews.com")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,".order-0 p")
            for element in elements:
                data["global_news"].append(element.text)
            data["global_news"]= [article for article in data["global_news"] if article]
            driver.quit()
        elif name == "Taipei Times":
            driver = webdriver.Chrome()
            driver.get("https://www.taipeitimes.com")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,".bf2")
            for element in elements:
                data["global_news"].append(element.text)
            data["global_news"]= [article for article in data["global_news"] if article]
            driver.quit()
        elif name == "Focus Taiwan":
            driver = webdriver.Chrome()
            driver.get("https://focustaiwan.tw/business")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,".h2 span")
            for element in elements:
                data["global_news"].append(element.text)
            data["global_news"]= [article for article in data["global_news"] if article]
            driver.quit()
        elif name == "Taiwan News":
            driver = webdriver.Chrome()
            driver.get("https://www.taiwannews.com.tw/category/Business")
            html_content = driver.page_source
            elements = driver.find_elements(By.CSS_SELECTOR,".w-full p")
            for element in elements:
                data["global_news"].append(element.text)
            data["global_news"]= [article for article in data["global_news"] if article]
            driver.quit()
    print(data["global_news"])
             
        

    csv_file_path = 'output_vertical.csv'
    
    with open(csv_file_path, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        writer.writerow(['Category', 'Item'])

        for category, items in data.items():
            for item in items:
                writer.writerow([category, item])
    import os

    csv_file_path = 'output_vertical.csv'
    print("File exists:", os.path.exists(csv_file_path))
    print("Writable directory:", os.access('.', os.W_OK))

    transposed_data = {}
    max_length = max(len(items) for items in data.values())
    for category, items in data.items():
        for i in range(max_length):
            if i < len(items):
                transposed_data.setdefault(i, []).append(items[i])
            else:
                transposed_data.setdefault(i, []).append('')

    csv_file_path = 'output_horizontal.csv'

    # Write the transposed data to a CSV file
    with open(csv_file_path, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        # Write the header
        writer.writerow(data.keys())
        # Write the rows
        for row in transposed_data.values():
            writer.writerow(row)

    print("File exists:", os.path.exists(csv_file_path))
    print("Writable directory:", os.access('.', os.W_OK))

    
if __name__ == "__main__":
    main()
