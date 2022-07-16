from selenium import webdriver
import requests
from bs4 import BeautifulSoup
from selenium.webdriver.common.keys import Keys

driver = webdriver.Chrome(executable_path=input('path:'))
linklist = []
url_list = ['https://www.ranker.com/list-of/film?ref=mainnav','https://www.ranker.com/list-of/tv?ref=mainnav']

def get_html2(url, params=None):
    try:
        html = driver.page_source
    except:
        html = ''
        pass
    return html


def get_html(url, params):
    r = requests.get(url, params=params)
    return r


lists_websites = []


def get_all_links(url):
    driver.get(url)
    html = driver.find_element_by_tag_name('html')
    for i in range(10000):
        html.send_keys(Keys.PAGE_DOWN)
        print(
            i)
    html = get_html2(url, params=None)
    soup = BeautifulSoup(html, 'lxml')
    items = soup.find_all('section')

    for item in items:
        for i in item:
            try:
               linkk = i.get('href')
               lists_websites.append('https://www.ranker.com/' + linkk)
            except:
                pass


for url in url_list:
    get_all_links(url)
print(len(lists_websites))
a = open('urllist11.txt', 'w')
for link in lists_websites:
    a.write(link + '\n')
a.close()

