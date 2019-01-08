# common imports
from lxml import etree
from selenium import webdriver
# from selenium.webdriver.chrome.options import Options
import time
import pandas as pd
from tqdm import tqdm


fireFoxOptions = webdriver.FirefoxOptions()
fireFoxOptions.add_argument('--headless')
browser = webdriver.Firefox(executable_path="/usr/local/bin/geckodriver", firefox_options=fireFoxOptions)
# browser = webdriver.Chrome(chrome_options=chrome_options)
"""
chrome_options = Options()
chrome_options.add_argument('--no-sandbox')  # 解决DevToolsActivePort文件不存在的报错
# chrome_options.add_argument('window-size=1920x3000')  # 指定浏览器分辨率
chrome_options.add_argument('--disable-gpu')  # 谷歌文档提到需要加上这个属性来规避bug
chrome_options.add_argument('--hide-scrollbars')  # 隐藏滚动条, 应对一些特殊页面
chrome_options.add_argument('blink-settings=imagesEnabled=false')  # 不加载图片, 提升速度
# chrome_options.add_argument('--headless')  # 浏览器不提供可视化页面. linux下如果系统不支持可视化不加这条会启动失败
# 添加代理 proxy
# chrome_options.add_argument('--proxy-server=http://' + proxy)
browser = webdriver.Chrome(chrome_options=chrome_options)
"""


def get_baidu(keyword):
    browser.get('https://www.baidu.com/')
    search = browser.find_element_by_xpath('//*[@id="kw"]')
    search.send_keys(keyword)
    time.sleep(1)
    browser.find_element_by_xpath('//*[@id="su"]').click()
    time.sleep(1)


def get_page():
    html = browser.page_source
    return html


def next(next_url):
    url = 'https://www.baidu.com' + next_url
    browser.get(url)


def scrape_multi(keywords):
    MAX_PAGE = 5
    result = pd.DataFrame(columns=('keyword', 'rank', 'title', 'domain'))
    for keyword in keywords:
        try:
            domains = []
            titles = []
            keyword_list = []
            get_baidu(keyword)
            print('---开始抓取[{k}]---'.format(k=keyword))
            for i in tqdm(range(1, MAX_PAGE + 1)):
                print('第{i}页'.format(i=i))
                html = get_page()
                next_url = parse_multi(html, domains, titles, keyword_list, keyword)
                next(next_url)
            tmp_frame = frame(titles, domains, keyword_list)
            result = pd.concat([result, tmp_frame], ignore_index=True)
        except Exception:
            pass
    print(result)
    result.to_excel('./data/keyword-multi.xlsx')


def parse_multi(html, domains, titles, keyword_list, keyword):
    data = etree.HTML(html)
    items = data.xpath('//*[@id="content_left"]/div')
    try:
        next = data.xpath('//*[@id="page"]/a[@class="n"]/@href')[0]
        for item in items:
            title_list = item.xpath('./h3/a//text()')
            title = "".join(title_list)
            print(title)
            if title == '':
                pass
            else:
                domain = "".join(item.xpath('.//*[@class="c-showurl"]/text()'))
                domain = domain.replace('\xa0', '')
                domains.append(domain)
                titles.append(title)
                keyword_list.append(keyword)

        print('=' * 20)
        return next
    except Exception:
        pass


def frame(titles, domains, keywords_list):
    ranks = list(range(1, len(titles)+1))
    frame_dict = {
        'keyword': keywords_list,
        'rank': ranks,
        'title': titles,
        'domain': domains,
    }
    frame = pd.DataFrame(frame_dict)
    return frame

