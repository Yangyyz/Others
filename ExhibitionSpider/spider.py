import requests
import re
from lxml import html

lists=[["none" for j in range(10)] for i in range(50)]

website="http://www.onezh.com"
url='http://www.onezh.com/zhanhui/1_1_0_0_20190901/20190930/' 
page=requests.Session().get(url) 
tree=html.fromstring(page.text) 

subpage=tree.xpath('//div[@class="row"]/a/@href') 

loopControlVar=0
for val in subpage:

    page_sub=requests.Session().get(website+val)
    tree_sub=html.fromstring(page_sub.text)

    title=tree_sub.xpath('//div[@class="tuan-detail wrap"]/h1/text()')
    date=tree_sub.xpath('//dl[@class="tuan-info"]/dd/div/text()')
    pavilion=tree_sub.xpath('//div[@class="bao-key"]/a/span/text()')
    host=tree_sub.xpath('//dl[@class="tuan-info mp5"]/dd/text()')
    hostinfo=tree_sub.xpath('//div[@class="top_dealer_1"]/ul/li/text()')

    # 标题
    pattern=re.compile(r'(\S+)')
    result=pattern.search(title[0])
    #print(result.group(1))
    if result:
        lists[loopControlVar][0]=result.group(1)

    #日期
    pattern=re.compile(r'(.+?)\xa0')
    result=pattern.search(date[0])
    #print(result.group(1))

    pattern=re.compile(r'(年|月)')
    result1=pattern.sub("/",result.group(1))
    #print(result1)

    pattern=re.compile(r'(日)')
    result2=pattern.sub("",result1)
    #print(result2)

    pattern=re.compile(r'---')
    result3=pattern.sub("---2019/",result2)
    #print(result3)
    if result:
        lists[loopControlVar][1]=result3

    #展馆
    #print(pavilion[1])

    if len(pavilion) > 1:
        lists[loopControlVar][2]=pavilion[1]


    #主办单位
    pattern=re.compile(r'：(.+)')
    result=pattern.search(host[0])
    #print(result.group(1))
    if result:
        lists[loopControlVar][3]=result.group(1)

    #承办单位
    pattern=re.compile(r'：(.+)')
    result=pattern.search(host[1])
    #print(result.group(1))
    if result:
        lists[loopControlVar][4]=result.group(1)

    #官网
    pattern=re.compile(r'((www|http).+?)\'')
    result=pattern.search(str(hostinfo))
    #print(result.group(1))
    if result:
        lists[loopControlVar][5]=result.group(1)
    
    loopControlVar+=1

print(lists)