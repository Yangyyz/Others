import requests
from lxml import html

lists=[[] for i in range(10)]

website="http://www.onezh.com"
url='http://www.onezh.com/zhanhui/1_1_0_0_20190901/20190930/' 
page=requests.Session().get(url) 
tree=html.fromstring(page.text) 

subpage=tree.xpath('//div[@class="row"]/a/@href') 

page_sub=requests.Session().get(website+subpage[0])
tree_sub=html.fromstring(page_sub.text)

title=tree_sub.xpath('//div[@class="tuan-detail wrap"]/h1/text()')
date=tree_sub.xpath('//dl[@class="tuan-info"]/dd/div/text()')
pavilion=tree_sub.xpath('//div[@class="bao-key"]/a/span/text()')
host=tree_sub.xpath('//dl[@class="tuan-info mp5"]/dd/text()')
hostinfo=tree_sub.xpath('//div[@class="top_dealer_1"]/ul/li/text()')

print(title,date,pavilion,host,hostinfo)


