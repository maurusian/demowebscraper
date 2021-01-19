from bs4 import BeautifulSoup
import requests
from openpyxl import Workbook
import urllib.parse as up

def build_url(url_pattern,keyword,page_number):
    return url_pattern.format(keyword,page_number)

def get_page_html(url):
    page = requests.get(url)
    return page.content

def get_search_result_item_list(page_html):
    soup = BeautifulSoup(page_html,'html.parser')
    tables = soup.find_all('table')
    items = []
    for i in range(len(tables)):
        if tables[i].text is not None and tables[i].text != "":
            children = tables[i].children

            for child in children:
                if child.name == "tr":
                    for cc in child.children:
                        if cc.name == "td":
                            for ee in cc.descendants:
                                if ee.name == "a":
                                    product_name = ee.text
                                    product_url = ee['href'].strip()
                                    if (product_name != 'Versand' and product_name is not None and product_name != '' and 
                                    product_url is not None and product_url !='' and len(product_url)>6 and product_url[:6] != 'search' and product_url[0]!='/'):
                                        items.append({'product':product_name,'url':product_url})
                                    
    return items

def get_search_page_count(page_html):
    soup = BeautifulSoup(page_html,'html.parser')
    tds = soup.find_all('td')
    count = 0
    for td in tds:
        if 'class' in td.attrs.keys() and 'srchapages' in td['class']:
            for child in td.children:
                if child.name in ['b','a'] and child.text.strip() is not None:
                    count+=1
    return count

def get_item_details(item_dict):
    item_page_html = get_page_html(item_dict['url'])
    soup = BeautifulSoup(item_page_html,'html.parser')
    divs = soup.find_all('div')
    for div in divs:
        if div is not None and 'class' in div.attrs.keys() and div['class'] is not None:
            if 'pull-left' in div['class']:
                for div_child in div.children:
                    if div_child.name == 'span' and 'class' in div_child.attrs.keys() and 'text-primary' in div_child['class']:
                        
                        article_number = div_child.text.strip()
                        print(div.get_text())
                        man_number = div.get_text().strip()[len(article_number):]
                        
                        item_dict['article_num'] = article_number
                        item_dict['man_num'] = man_number
                        break
            elif 'prod-caption' in div['class']:
                item_dict['product'] = div.text.strip()
            
            elif 'pricewrap' in div['class']:
                for descendant in div.descendants:
                    if descendant.name == 'span' and 'content' in descendant.attrs.keys() and 'itemprop' in descendant.attrs.keys() and descendant['itemprop'] == 'priceCurrency':
                        item_dict['currency'] = descendant['content']
                    elif descendant.name == 'span' and 'content' in descendant.attrs.keys() and 'itemprop' in descendant.attrs.keys() and descendant['itemprop'] == 'price':
                        item_dict['price'] = float(descendant['content'])
                        

    return item_dict
    


def export_items_to_xlsx(item_list,keyword):
    wb = Workbook()

    sheet = wb.active

    sheet['A1'] = 'Produkt'
    sheet['B1'] = 'Produkt nummer 1'
    sheet['C1'] = 'Produkt nummer 2'
    sheet['D1'] = 'Preis'
    sheet['E1'] = 'Waerung'
    sheet['F1'] = 'URL'

    for i in range(len(item_list)):
        
        sheet['A'+str(i+2)] = item_list[i]['product']
        
        if 'article_num' in item_list[i].keys():
            sheet['B'+str(i+2)] = item_list[i]['article_num']
      
        if 'man_num' in item_list[i].keys():
            sheet['C'+str(i+2)] = item_list[i]['man_num']

        if 'price' in item_list[i].keys():
            sheet['D'+str(i+2)] = item_list[i]['price']
            
        if 'currency' in item_list[i].keys():
            sheet['E'+str(i+2)] = item_list[i]['currency']
            
        sheet['F'+str(i+2)] = item_list[i]['url']

    wb.save(keyword+'.xlsx')
        

if __name__ == '__main__':
    lapstore_url_pattern = "https://www.lapstore.de/search.php?shop=lapstore&searcharg={}&artcpage={}"
    keyword = input()
    
    items = []
    page_number = 0
    url = build_url(lapstore_url_pattern,up.quote(keyword),0)
    page_html = get_page_html(url)
    search_page_count = get_search_page_count(page_html)
    for page_number in range(search_page_count):
        print('page '+str(page_number+1))
        
        url = build_url(lapstore_url_pattern,up.quote(keyword),page_number)
        print(url)
        page_html = get_page_html(url)
       
        
        items += get_search_result_item_list(page_html)

    i = 0
    while i < len(items):
        print(items[i]['url'])
        items[i] = get_item_details(items[i])
 
        i+=1
    if len(items) > 0:
        export_items_to_xlsx(items,keyword)
    else:
        print("no results found for keyword \"{}\"".format(keyword))


        
