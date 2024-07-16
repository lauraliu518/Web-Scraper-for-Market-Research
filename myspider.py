import scrapy
from scrapy.crawler import CrawlerProcess

class ConsultingFirmsSpider(scrapy.Spider):
    name = 'consulting_firms'
    allowed_domains = ['bcg.com', 'mckinsey.com', 'deloitte.com']
    start_urls = [
        'https://www.bcg.com',
        'https://www.mckinsey.com',
        'https://www.deloitte.com'
    ]

    def parse(self, response):
        if 'bcg.com' in response.url:
            return self.parse_bcg(response)
        elif 'mckinsey.com' in response.url:
            return self.parse_mckinsey(response)
        elif 'deloitte.com' in response.url:
            return self.parse_deloitte(response)

    def parse_bcg(self, response):
        # Market Insights
        for insight in response.css('h3.article-block__title::text').getall():
            yield {
                'site': 'BCG',
                'category': 'Market Insight',
                'text': insight.strip()
            }

        # Case Studies
        for case in response.css('h1.homepageAnimatedHero-title::text').getall():
            yield {
                'site': 'BCG',
                'category': 'Case Study',
                'text': case.strip()
            }

        # Service Offerings
        for service in response.css('li.glide__slide::text').getall():
            yield {
                'site': 'BCG',
                'category': 'Service Offering',
                'text': service.strip()
            }

    def parse_mckinsey(self, response):
        # Market Insights
        for insight in response.css(".HighlightBar_mck-c-highlight-bar__item-title__g9Fdc span::text").getall():
            yield {
                'site': 'McKinsey',
                'category': 'Market Insight',
                'text': insight.strip()
            }

        # Case Studies
        for case in response.css('h1.mck-c-hwpm__heading--text-m::text').getall():
            yield {
                'site': 'McKinsey',
                'category': 'Case Study',
                'text': case.strip()
            }

        # Service Offerings
        for service in response.css('[data-testid="hamburger-list-Capabilities"] span::text').getall():
            yield {
                'site': 'McKinsey',
                'category': 'Service Offering',
                'text': service.strip()
            }

        # Thought Leadership Articles
        for article in response.css("h3.mdc-c-heading___0fM1W_d0bf44e::text").getall():
            yield {
                'site': 'McKinsey',
                'category': 'Thought Leadership',
                'text': article.strip()
            }

    def parse_deloitte(self, response):
        # Market Insights
        for insight in response.css('.trending-item h4::text').getall():
            yield {
                'site': 'Deloitte',
                'category': 'Market Insight',
                'text': insight.strip()
            }

        # Case Studies
        for case in response.css('.promo-focus h2::text').getall():
            yield {
                'site': 'Deloitte',
                'category': 'Case Study',
                'text': case.strip()
            }

        # Service Offerings
        for service in response.css("#footer__links-services a::text").getall():
            yield {
                'site': 'Deloitte',
                'category': 'Service Offering',
                'text': service.strip()
            }

        # Thought Leadership
        for article in response.css(".showcase-content-wrap h1::text").getall():
            yield {
                'site': 'Deloitte',
                'category': 'Thought Leadership',
                'text': article.strip()
            }

class IndustryNewsSpider(scrapy.Spider):
    name = 'industry_news'
    allowed_domains = ['bloomberg.com', 'reuters.com', 'ft.com']
    start_urls = [
        'https://www.bloomberg.com',
        'https://www.reuters.com',
        'https://www.ft.com'
    ]

    custom_settings = {
        'USER_AGENT': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
    }

    def parse(self, response):
        if 'bloomberg.com' in response.url:
            if response.url == 'https://www.bloomberg.com':
                # Latest news
                for a in response.css('.StoryListLatestFiltered_storyBlock___szNS a::text').getall():
                    yield {'site': 'Bloomberg', 'category': 'Latest News', 'title': a.strip()}

                # Follow links to market reports and financial analysis
                yield response.follow('https://www.bloomberg.com/markets', callback=self.parse_bloomberg_markets)
                yield response.follow('https://www.bloomberg.com/search?query=financial%20markets', callback=self.parse_bloomberg_financial)

            # Global economic trends
            elif 'economics' in response.url:
                for a in response.css('.styles_storyBlock__l5VzV a::text').getall():
                    yield {'site': 'Bloomberg', 'category': 'Global Economic Trends', 'title': a.strip()}

        elif 'reuters.com' in response.url:
            # Latest news
            for span in response.css('.basic-card__container__1y8wi span::text').getall():
                yield {'site': 'Reuters', 'category': 'Latest News', 'title': span.strip()}

        elif 'ft.com' in response.url:
            if response.url == 'https://www.ft.com':
                # Latest news and Market reports
                for span in response.css('.story-group-stacked__primary-story span::text, .js-teaser-headline span::text').getall():
                    yield {'site': 'Financial Times', 'category': 'Latest News/Market Reports', 'title': span.strip()}

                # Follow links to financial analysis and global economic trends
                yield response.follow('https://markets.ft.com/data', callback=self.parse_ft_financial)
                yield response.follow('https://www.ft.com/world', callback=self.parse_ft_world)

    def parse_bloomberg_markets(self, response):
        for a in response.css('.styles_storyBlock__l5VzV a::text').getall():
            yield {'site': 'Bloomberg', 'category': 'Market Reports', 'title': a.strip()}

    def parse_bloomberg_financial(self, response):
        for a in response.css('.headline__3a97424275 a::text').getall():
            yield {'site': 'Bloomberg', 'category': 'Financial Analysis', 'title': a.strip()}

    def parse_ft_financial(self, response):
        for span in response.css('.mod-news span::text').getall():
            yield {'site': 'Financial Times', 'category': 'Financial Analysis', 'title': span.strip()}

    def parse_ft_world(self, response):
        for a in response.css('.o-teaser__heading a::text').getall():
            yield {'site': 'Financial Times', 'category': 'Global Economic Trends', 'title': a.strip()}

class MarketResearchSpider(scrapy.Spider):
    name = 'market_research'
    allowed_domains = ['statista.com', 'marketresearch.com', 'ibisworld.com']
    start_urls = [
        'https://www.statista.com/studies-and-reports/industries-and-markets',
        'https://www.marketresearch.com/Marketing-Market-Research-c70/',
        'https://www.ibisworld.com/market-research-reports/#global'
    ]

    def parse(self, response):
        # Identify which site we're scraping based on the start URL
        if 'statista.com' in response.url:
            self.parse_statista(response)
        elif 'marketresearch.com' in response.url:
            self.parse_marketresearch(response)
        elif 'ibisworld.com' in response.url:
            self.parse_ibisworld(response)

    def parse_statista(self, response):
        # Handle different sections of the Statista website
        for h3 in response.css('.reportResult__content h3::text').getall():
                yield {'site': 'Statista', 'category': 'Market Research Reports', 'title': h3.strip()}

            # Continue to industry analysis page
        yield response.follow('https://www.statista.com/studies-and-reports/companies-and-products', self.parse_statista)

        for h3 in response.css('.reportResult__content h3::text').getall():
                yield {'site': 'Statista', 'category': 'Industry Analysis', 'title': h3.strip()}

            # Continue to statistics page
        yield response.follow('https://www.statista.com/statistics/popular/', self.parse_statista)

        for title in response.css('.statList__title::text').getall():
                yield {'site': 'Statista', 'category': 'Statistics', 'title': title.strip()}

            # Continue to trends page
        yield response.follow('https://www.statista.com/studies-and-reports/digital-and-trends', self.parse_statista)

        for h3 in response.css('.reportResult__content h3::text').getall():
                yield {'site': 'Statista', 'category': 'Trends', 'title': h3.strip()}

    def parse_marketresearch(self, response):
        # Handle different sections of the MarketResearch.com website
        for a in response.css('.order-1 a::text').getall():
            yield {'site': 'MarketResearch.com', 'category': 'General Market Research', 'title': a.strip()}

    def parse_ibisworld(self, response):
        # Handle different sections of the IBISWorld website
        for span in response.css('.country-reports span::text').getall():
                yield {'site': 'IBISWorld', 'category': 'Market Research Reports', 'title': span.strip()}

            # Continue to industry analysis page
        yield response.follow('https://www.ibisworld.com/united-states/list-of-industries/#specialized-reports', self.parse_ibisworld)

        for a in response.css('#sectorsIndustryList a::text').getall():
                yield {'site': 'IBISWorld', 'category': 'Industry Analysis', 'title': a.strip()}

            # Continue to statistics page
        yield response.follow('https://www.ibisworld.com/search/default.aspx?st=statistics', self.parse_ibisworld)

        for span in response.css('.hit-name span::text').getall():
                yield {'site': 'IBISWorld', 'category': 'Statistics', 'title': span.strip()}

            # Continue to global industry trends page
        yield response.follow('https://www.ibisworld.com/global/list-of-industries/', self.parse_ibisworld)

        for text in response.css('#sectorsIndustryList::text').getall():
                yield {'site': 'IBISWorld', 'category': 'Trends', 'title': text.strip()}


class GlobalNewsSpider(scrapy.Spider):
    name = 'global_news'
    allowed_domains = ['economictimes.indiatimes.com', 'business-standard.com', 'hindustantimes.com', 'chinadaily.com.cn', 'globaltimes.cn', 'caixinglobal.com', 'japantimes.co.jp', 'nikkei.com', 'nhk.or.jp', 'straitstimes.com', 'channelnewsasia.com', 'businesstimes.com.sg', 'koreaherald.com', 'koreatimes.co.kr', 'gulfnews.com', 'khaleejtimes.com', 'thenationalnews.com', 'taipeitimes.com', 'focustaiwan.tw', 'taiwannews.com.tw']
    start_urls = [
        'https://economictimes.indiatimes.com/news/economy',
        'https://www.business-standard.com/',
        'https://www.hindustantimes.com',
        'https://www.chinadaily.com.cn/',
        'https://www.globaltimes.cn',
        'https://www.caixinglobal.com',
        'https://www.japantimes.co.jp',
        'https://asia.nikkei.com',
        'https://www.nhk.or.jp/nhkworld',
        'https://straitstimes.com',
        'https://www.channelnewsasia.com',
        'https://www.businesstimes.com.sg',
        'https://www.koreaherald.com',
        'https://www.koreatimes.co.kr/www/section_129.html',
        'https://www.gulfnews.com',
        'https://www.khaleejtimes.com',
        'https://www.thenationalnews.com',
        'https://www.taipeitimes.com',
        'https://focustaiwan.tw/business',
        'https://www.taiwannews.com.tw/category/Business'
    ]

    def parse(self, response):
        site_name = response.url.split('/')[2].replace('www.', '').replace('.com', '')
        category_selector_mapping = {
            'economictimes.indiatimes.com': ".top-news li::text",
            'business-standard.com': ".cardlist p::text",
            'hindustantimes.com': ".cartHolder a::text",
            'chinadaily.com.cn': ".txt1::text",
            'globaltimes.cn': ".row p::text",
            'caixinglobal.com': "h2 a::text",
            'japantimes.co.jp': ".article-title a::text",
            'nikkei.com': "h4 span::text",
            'nhk.or.jp': "h3 a::text",
            'straitstimes.com': "h5 a::text",
            'channelnewsasia.com': "h6 a::text",
            'businesstimes.com.sg': ".inherit .word-break::text",
            'koreaherald.com': ".ellipsis2::text",
            'koreatimes.co.kr': ".LoraMedium::text",
            'gulfnews.com': "h2 a::text",
            'khaleejtimes.com': "h2 a::text",
            'thenationalnews.com': ".order-0 p::text",
            'taipeitimes.com': ".bf2::text",
            'focustaiwan.tw': ".h2 span::text",
            'taiwannews.com.tw': ".w-full p::text"
        }

        selector = category_selector_mapping.get(response.url.split('/')[2])
        if selector:
            for text in response.css(selector).getall():
                yield {
                    'site': site_name,
                    'category': 'Global News',
                    'content': text.strip()
                }

# To run the spider directly if needed
if __name__ == "__main__":
    process = CrawlerProcess(settings={
        'FEED_URI': 'output.csv',
        'FEED_FORMAT': 'csv'
    })
    process.crawl(ConsultingFirmsSpider)
    process.crawl(IndustryNewsSpider)
    process.crawl(MarketResearchSpider)
    process.crawl(GlobalNewsSpider)
    process.start()
