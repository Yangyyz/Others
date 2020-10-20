import requests
import re
from lxml import html
import xlwings as xw
import xlrd
import xlwt

lists=[["none" for j in range(10)] for i in range(50)]

website="http://www.onezh.com"
url='http://www.onezh.com/zhanhui/2_998_0_0_20200101/20201231/' 
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
    result3=pattern.sub("---2020/",result2)

    pattern=re.compile(r'/([0-9])-')
    result4=pattern.sub(r"/0\1-",result3)

    pattern=re.compile(r'/([0-9])$')
    result5=pattern.sub(r"/0\1",result4)

    pattern=re.compile(r'/([0-9])/')
    result6=pattern.sub(r"/0\1/",result5)

    #print(result3)
    if result:
        lists[loopControlVar][1]=result6

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

    #分类
    for key in {"农业","种子","花卉"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "农业林业渔业及农副产品"
            lists[loopControlVar][7] = "农业"

    for key in {"林业","森林","树"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "农业林业渔业及农副产品"
            lists[loopControlVar][7] = "林业"
    
    for key in {"渔业","渔"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "农业林业渔业及农副产品"
            lists[loopControlVar][7] = "渔业"

    for key in {"畜牧业","牧"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "农业林业渔业及农副产品"
            lists[loopControlVar][7] = "畜牧业"

    for key in {"农副产品加工","农副产品","农产品"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "农业林业渔业及农副产品"
            lists[loopControlVar][7] = "农副产品加工"

    for key in {"食品制造","食品"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "食品酒饮及服务"
            lists[loopControlVar][7] = "食品制造"

    for key in {"酒饮料和精制茶制造","酒","茶"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "食品酒饮及服务"
            lists[loopControlVar][7] = "酒饮料和精制茶制造"

    for key in {"餐饮服务","餐饮","饭店","火锅","食材"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "食品酒饮及服务"
            lists[loopControlVar][7] = "餐饮服务"

    for key in {"煤炭开采及加工","煤"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "能源矿产"
            lists[loopControlVar][7] = "煤炭开采及加工"       

    for key in {"石油和天然气开采及加工","石油","天然气"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "能源矿产"
            lists[loopControlVar][7] = "石油和天然气开采及加工"       

    for key in {"黑色金属矿采选业","黑色金属"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "能源矿产"
            lists[loopControlVar][7] = "黑色金属矿采选业"       

    for key in {"有色金属矿采选业","有色金属"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "能源矿产"
            lists[loopControlVar][7] = "有色金属矿采选业"       

    for key in {"非金属矿采选业","矿采"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "能源矿产"
            lists[loopControlVar][7] = "非金属矿采选业"       

    for key in {"其他矿产开采及加工"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "能源矿产"
            lists[loopControlVar][7] = "其他矿产开采及加工"       

    for key in {"造纸及纸制品印刷","印刷","纸"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "工业科技"
            lists[loopControlVar][7] = "造纸及纸制品印刷"

    for key in {"化学原料和化学制品","化学"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "工业科技"
            lists[loopControlVar][7] = "化学原料和化学制品"

    for key in {"橡胶和塑料制品","橡胶","塑料"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "工业科技"
            lists[loopControlVar][7] = "橡胶和塑料制品"

    for key in {"非金属矿物制品","非金属"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "工业科技"
            lists[loopControlVar][7] = "非金属矿物制品"

    for key in {"通用设备制造","设备","装备"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "工业科技"
            lists[loopControlVar][7] = "通用设备制造"

    for key in {"电气机械和器材","电气","机械"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "工业科技"
            lists[loopControlVar][7] = "电气机械和器材"

    for key in {"电子器件"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "工业科技"
            lists[loopControlVar][7] = "电子器件"

    for key in {"仪器仪表","仪器","仪表"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "工业科技"
            lists[loopControlVar][7] = "仪器仪表"

    for key in {"人工智能"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "工业科技"
            lists[loopControlVar][7] = "人工智能"

    for key in {"环境保护及废弃资源综合利用","环保","资源利用","净水","新风"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "工业科技"
            lists[loopControlVar][7] = "环境保护及废弃资源综合利用"

    for key in {"金属制品机械和设备修理","设备维修"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "工业科技"
            lists[loopControlVar][7] = "金属制品机械和设备修理"

    for key in {"铁路交通运输","高铁","铁路","地铁"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "交通运输仓储和邮政"
            lists[loopControlVar][7] = "铁路交通运输"

    for key in {"道路交通运输","汽车","卡车","车"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "交通运输仓储和邮政"
            lists[loopControlVar][7] = "道路交通运输"

    for key in {"水上交通运输","船"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "交通运输仓储和邮政"
            lists[loopControlVar][7] = "水上交通运输"

    for key in {"航空航天","飞机","航空","航天"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "交通运输仓储和邮政"
            lists[loopControlVar][7] = "航空航天"

    for key in {"仓储","仓库","储存"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "交通运输仓储和邮政"
            lists[loopControlVar][7] = "仓储"

    for key in {"邮政","快递","物流"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "交通运输仓储和邮政"
            lists[loopControlVar][7] = "邮政"

    for key in {"其他运输设备及服务"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "交通运输仓储和邮政"
            lists[loopControlVar][7] = "其他运输设备及服务"

    for key in {"计算机通信和其他电子设备","计算机","通信"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "信息传输软件和信息技术"
            lists[loopControlVar][7] = "计算机通信和其他电子设备"

    for key in {"电信广播电视和卫星传输服务","电信","广播"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "信息传输软件和信息技术"
            lists[loopControlVar][7] = "电信广播电视和卫星传输服务"

    for key in {"互联网和相关服务","互联网"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "信息传输软件和信息技术"
            lists[loopControlVar][7] = "互联网和相关服务"

    for key in {"软件和信息技术","软件","信息技术"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "信息传输软件和信息技术"
            lists[loopControlVar][7] = "软件和信息技术"

    for key in {"医药制造","药"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "医疗健康"
            lists[loopControlVar][7] = "医药制造"

    for key in {"医疗用品及器材","医疗器材","医疗用品","医疗器械"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "医疗健康"
            lists[loopControlVar][7] = "医药制造"

    for key in {"护理及其他医疗健康服务","护理","医疗健康"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "医疗健康"
            lists[loopControlVar][7] = "护理及其他医疗健康服务"

    for key in {"货币金融服务","货币","金融"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "金融"
            lists[loopControlVar][7] = "货币金融服务"

    for key in {"保险业","保险"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "金融"
            lists[loopControlVar][7] = "保险业"

    for key in {"互联网及其他金融业","新零售"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "金融"
            lists[loopControlVar][7] = "互联网及其他金融业"

    for key in {"土木工程","土木"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "房屋建筑装修及经营服务"
            lists[loopControlVar][7] = "土木工程"

    for key in {"房屋建筑","建筑","房屋"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "房屋建筑装修及经营服务"
            lists[loopControlVar][7] = "房屋建筑"

    for key in {"建筑安装"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "房屋建筑装修及经营服务"
            lists[loopControlVar][7] = "建筑安装"

    for key in {"建筑装饰和其他建筑业","装修","家博会"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "房屋建筑装修及经营服务"
            lists[loopControlVar][7] = "建筑装饰和其他建筑业"

    for key in {"家装设计及家具","家具","家装","家居"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "房屋建筑装修及经营服务"
            lists[loopControlVar][7] = "家装设计及家具"

    for key in {"房地产"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "房屋建筑装修及经营服务"
            lists[loopControlVar][7] = "房地产"

    for key in {"租赁"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "租赁和商务服务"
            lists[loopControlVar][7] = "租赁"

    for key in {"咨询与调查","咨询","调查"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "租赁和商务服务"
            lists[loopControlVar][7] = "咨询与调查"

    for key in {"广告设计","广告"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "租赁和商务服务"
            lists[loopControlVar][7] = "广告设计"

    for key in {"人力资源"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "租赁和商务服务"
            lists[loopControlVar][7] = "人力资源"

    for key in {"旅行及相关服务","旅游"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "租赁和商务服务"
            lists[loopControlVar][7] = "旅行及相关服务"

    for key in {"安全保护服务","安全","保护","安防","安保"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "租赁和商务服务"
            lists[loopControlVar][7] = "安全保护服务"

    for key in {"电子商务"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "租赁和商务服务"
            lists[loopControlVar][7] = "电子商务"

    for key in {"其他商务服务业","加盟"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "租赁和商务服务"
            lists[loopControlVar][7] = "其他商务服务业"

    for key in {"纺织面料服装及服饰","服饰","时装"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "日用消费品及居民服务"
            lists[loopControlVar][7] = "纺织面料服装及服饰"

    for key in {"皮革及箱包","皮革","箱包"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "日用消费品及居民服务"
            lists[loopControlVar][7] = "皮革及箱包"

    for key in {"家用电器及电子产品","电器","手机","电子产品"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "日用消费品及居民服务"
            lists[loopControlVar][7] = "家用电器及电子产品"

    for key in {"玩具"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "日用消费品及居民服务"
            lists[loopControlVar][7] = "玩具"

    for key in {"化妆品卫生用品及美容美发服务","化妆品","卫生用品","美容","美发"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "日用消费品及居民服务"
            lists[loopControlVar][7] = "化妆品卫生用品及美容美发服务"

    for key in {"钟表眼镜","钟","表","眼镜"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "日用消费品及居民服务"
            lists[loopControlVar][7] = "钟表眼镜"

    for key in {"珠宝首饰","珠宝","首饰"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "日用消费品及居民服务"
            lists[loopControlVar][7] = "珠宝首饰"

    for key in {"办公设备及服务","办公"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "日用消费品及居民服务"
            lists[loopControlVar][7] = "办公设备及服务"

    for key in {"殡葬设施及服务","殡","葬"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "日用消费品及居民服务"
            lists[loopControlVar][7] = "殡葬设施及服务"

    for key in {"家政服务","家政"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "日用消费品及居民服务"
            lists[loopControlVar][7] = "家政服务"

    for key in {"婚庆设施及服务","婚"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "日用消费品及居民服务"
            lists[loopControlVar][7] = "婚庆设施及服务"

    for key in {"其他产品及服务","消费品"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "日用消费品及居民服务"
            lists[loopControlVar][7] = "其他产品及服务"

    for key in {"教育"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "教育"
            lists[loopControlVar][7] = "教育"

    for key in {"教育机构及培训"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "教育"
            lists[loopControlVar][7] = "教育机构及培训"

    for key in {"其他教育产品及服务"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "教育"
            lists[loopControlVar][7] = "其他教育产品及服务"

    for key in {"新闻和出版业","新闻","出版"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "文化体育和娱乐"
            lists[loopControlVar][7] = "新闻和出版业"

    for key in {"广播电视电影和影视制作","影视","电影","视频"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "文化体育和娱乐"
            lists[loopControlVar][7] = "广播电视电影和影视制作"

    for key in {"文化艺术","艺术","文化","佛"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "文化体育和娱乐"
            lists[loopControlVar][7] = "文化艺术"

    for key in {"体育","运动"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "文化体育和娱乐"
            lists[loopControlVar][7] = "体育"

    for key in {"娱乐"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "文化体育和娱乐"
            lists[loopControlVar][7] = "娱乐"

    loopControlVar+=1

workbook = xlwt.Workbook(encoding='utf-8') #创建workbook 对象
worksheet = workbook.add_sheet('sheet1')   #创建工作表sheet
for index in range(len(lists)):
    worksheet.write(index, 1, lists[index][0]) 
    worksheet.write(index, 2, lists[index][1]) 
    worksheet.write(index, 5, lists[index][2]) 
    worksheet.write(index, 8, lists[index][3]) 
    worksheet.write(index, 9, lists[index][4]) 
    worksheet.write(index, 14, lists[index][5])
    worksheet.write(index, 10, lists[index][6])
    worksheet.write(index, 11, lists[index][7])
    

workbook.save('xxxx.xls') 