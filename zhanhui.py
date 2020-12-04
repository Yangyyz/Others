import requests
import re
from lxml import html
import xlwings as xw
import xlrd
import xlwt
import regex as re

lists=[[" " for j in range(10)] for i in range(50)]

website="http://www.onezh.com"
url='http://www.onezh.com/zhanhui/1_1643_1683_0_20201001/20201031/' 
page=requests.Session().get(url) 
tree=html.fromstring(page.text) 

subpage=tree.xpath('//div[@class="row"]/a/@href') 

def number_translator(target):
     
    def word2number(s):
        '''
        可将[零-九]正确翻译为[0-9]

        :param s: 大写数字
        :return: 对应的整形数，如果不是数字返回-1
        '''
        if (s == u'零') or (s == '0'):
            return 0
        elif (s == u'一') or (s == '1') or (s == u'壹'):
            return 1
        elif (s == u'二') or (s == '2') or (s == u'贰') or (s == u'两'):
            return 2
        elif (s == u'三') or (s == '3') or (s == u'叁'):
            return 3
        elif (s == u'四') or (s == '4') or (s == u'肆'):
            return 4
        elif (s == u'五') or (s == '5') or (s == u'伍'):
            return 5
        elif (s == u'六') or (s == '6') or (s == u'陆'):
            return 6
        elif (s == u'七') or (s == '7') or (s == u'柒') or (s == u'天') or (s == u'日') or (s == u'末'):
            return 7
        elif (s == u'八') or (s == '8') or (s == u'捌'):
            return 8
        elif (s == u'九') or (s == '9') or (s == u'玖'):
            return 9
    #     elif (s == u'十') or (s == u'拾'):
    #         return 10
    #     elif (s == u'百') or (s == u'佰'):
    #         return 100
    #     elif (s == u'千') or (s == u'仟'):
    #         return 1000
    #     elif (s == u'万') or (s == u'萬'):
    #         return 10000
    #     elif (s == u'亿'):
    #         return 100000000
        else:
            return -1
        
    def str2int(s):
        '''
        将字符数字转换为int
        '''
        try:
            res = int(s)
        except:
            res = 0
        return res
    
    pattern = re.compile(u"[一二两三四五六七八九123456789]万[一二两三四五六七八九123456789](?!(亿|千|百|十))")
    match = pattern.finditer(target)
    for m in match:
        group = m.group()
        s = group.split(u"万")
        s = list(s)
        num = 0
        if len(s) == 2:
            num += word2number(s[0]) * 10000 + word2number(s[1]) * 1000
        target = pattern.sub(str(num), target, 1)
#     print(target)
    
    pattern = re.compile(u"[一二两三四五六七八九123456789]万[一二两三四五六七八九123456789](?!(千|百|十))")
    match = pattern.finditer(target)
    for m in match:
        group = m.group()
        s = group.split(u"万")
        s = list(s)
        num = 0
        if len(s) == 2:
            num += word2number(s[0]) * 10000 + word2number(s[1]) * 1000
        target = pattern.sub(str(num), target, 1)
#     print(target)

    pattern = re.compile(u"[一二两三四五六七八九123456789]千[一二两三四五六七八九123456789](?!(百|十))")
    match = pattern.finditer(target)
    for m in match:
        group = m.group()
        s = group.split(u"千")
        s = list(filter(None, s))
        num = 0
        if len(s) == 2:
            num += word2number(s[0]) * 1000 + word2number(s[1]) * 100
        target = pattern.sub(str(num), target, 1)
#     print(target)

    pattern = re.compile(u"[一二两三四五六七八九123456789]百[一二两三四五六七八九123456789](?!十)")
    match = pattern.finditer(target)
    for m in match:
        group = m.group()
        s = group.split(u"百")
        s = list(filter(None, s))
        num = 0
        if len(s) == 2:
            num += word2number(s[0]) * 100 + word2number(s[1]) * 10
        target = pattern.sub(str(num), target, 1)
#     print(target)

    pattern = re.compile(u"[零一二两三四五六七八九]")
    match = pattern.finditer(target)
    for m in match:
        target = pattern.sub(str(word2number(m.group())), target, 1)
#     print(target)

    pattern = re.compile(u"(?<=(周|星期|天|日))[天|日|末]")
    match = pattern.finditer(target)
    for m in match:
        target = pattern.sub(str(word2number(m.group())), target, 1)
#     print(target)

    pattern = re.compile(u"(?<!(周|星期))0?[0-9]?十[0-9]?")
    match = pattern.finditer(target)
    for m in match:
        group = m.group()
        s = group.split(u"十")
        num = 0
        ten = str2int(s[0])
        if ten == 0:
            ten = 1
        unit = str2int(s[1])
        num = ten * 10 + unit
        target = pattern.sub(str(num), target, 1)
#     print(target)

    pattern = re.compile(u"0?[1-9]百[0-9]?[0-9]?")
    match = pattern.finditer(target)
    for m in match:
        group = m.group()
        s = group.split(u"百")
        s = list(filter(None, s))
        num = 0
        if len(s) == 1:
            hundred = int(s[0])
            num += hundred * 100
        elif len(s) == 2:
            hundred = int(s[0])
            num += hundred * 100
            num += int(s[1])
        target = pattern.sub(str(num), target, 1)
#     print(target)

    pattern = re.compile(u"0?[1-9]千[0-9]?[0-9]?[0-9]?")
    match = pattern.finditer(target)
    for m in match:
        group = m.group()
        s = group.split(u"千")
        s = list(filter(None, s))
        num = 0
        if len(s) == 1:
            thousand = int(s[0])
            num += thousand * 1000
        elif len(s) == 2:
            thousand = int(s[0])
            num += thousand * 1000
            num += int(s[1])
        target = pattern.sub(str(num), target, 1)
#     print(target)

    pattern = re.compile(u"[0-9]+万[0-9]?[0-9]?[0-9]?[0-9]?")
    match = pattern.finditer(target)
    for m in match:
        group = m.group()
        s = group.split(u"万")
        s = list(filter(None, s))
        num = 0
        if len(s) == 1:
            tenthousand = int(s[0])
            num += tenthousand * 10000
        elif len(s) == 2:
            tenthousand = int(s[0])
            num += tenthousand * 10000
            num += int(s[1])
        target = pattern.sub(str(num), target, 1)
#     print(target)
    
    pattern = re.compile(u"[0-9]+亿[0-9]?[0-9]?[0-9]?[0-9]?[0-9]?[0-9]?[0-9]?[0-9]?")
    match = pattern.finditer(target)
    for m in match:
        group = m.group()
        s = group.split(u"亿")
        s = list(filter(None, s))
        num = 0
        if len(s) == 1:
            tenthousand = int(s[0])
            num += tenthousand * 100000000
        elif len(s) == 2:
            tenthousand = int(s[0])
            num += tenthousand * 100000000
            num += int(s[1])
        target = pattern.sub(str(num), target, 1)
#     print(target)

    return target


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

    # 届数
    pattern=re.compile(r'第(.+)届')
    result=pattern.search(lists[loopControlVar][0])
    if result:
        lists[loopControlVar][9]=number_translator(result.group(1)) 

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

    pattern=re.compile(r'\d.+\d$')
    result7=pattern.search(result6).group(0)

    #print(result3)
    if result:
        lists[loopControlVar][1]=result7

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
    if len(host) > 1:
        result=pattern.search(host[1])
        #print(result.group(1))
        if result:
            lists[loopControlVar][4]=result.group(1)

    #是否政府举办
    lists[loopControlVar][8] = "否"
    if "政府" in lists[loopControlVar][4]:
        lists[loopControlVar][8] = "是"
    if "政府" in lists[loopControlVar][3]:
        lists[loopControlVar][8] = "是"
              
    #官网
    pattern=re.compile(r'((www|http).+?)\'')
    result=pattern.search(str(hostinfo))
    #print(result.group(1))
    if result:
        lists[loopControlVar][5]=result.group(1)

    #分类
    for key in {"农业","种子","花卉","水果","农资","植保","农机"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "农业林业渔业及农副产品"
            lists[loopControlVar][7] = "农业"

    for key in {"林业","森林","树","园艺","竹产业"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "农业林业渔业及农副产品"
            lists[loopControlVar][7] = "林业"
    
    for key in {"渔业","渔","鱼","海洋"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "农业林业渔业及农副产品"
            lists[loopControlVar][7] = "渔业"

    for key in {"畜牧业","牧","宠","奶业","猪"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "农业林业渔业及农副产品"
            lists[loopControlVar][7] = "畜牧业"

    for key in {"农副产品加工","农副产品","农产品","乳业","乳制品","粮油"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "农业林业渔业及农副产品"
            lists[loopControlVar][7] = "农副产品加工"

    for key in {"食品制造","食品","烘培","烘焙"}:
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

    for key in {"煤炭开采及加工","煤","矿业"}:
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

    for key in {"其他矿产开采及加工","钛业","冶金"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "能源矿产"
            lists[loopControlVar][7] = "其他矿产开采及加工"       

    for key in {"造纸及纸制品印刷","印刷","纸"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "工业科技"
            lists[loopControlVar][7] = "造纸及纸制品印刷"

    for key in {"化学原料和化学制品","化学","化工","材料"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "工业科技"
            lists[loopControlVar][7] = "化学原料和化学制品"

    for key in {"橡胶和塑料制品","橡胶","塑料","橡塑","胶粘剂"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "工业科技"
            lists[loopControlVar][7] = "橡胶和塑料制品"

    for key in {"非金属矿物制品","非金属","陶瓷"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "工业科技"
            lists[loopControlVar][7] = "非金属矿物制品"

    for key in {"通用设备制造","设备","装备"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "工业科技"
            lists[loopControlVar][7] = "通用设备制造"

    for key in {"电气机械和器材","电气","机械","军博会","科技","机床","电力","照明","紧固件","线缆","LED","3D打印","防锈","工业","光电","科技","五金","磨具","热处理","变速机","燃料电池","泵阀"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "工业科技"
            lists[loopControlVar][7] = "电气机械和器材"

    for key in {"电子器件","电子","半导体"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "工业科技"
            lists[loopControlVar][7] = "电子器件"

    for key in {"仪器仪表","仪器","仪表","传感器","衡器"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "工业科技"
            lists[loopControlVar][7] = "仪器仪表"

    for key in {"人工智能"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "工业科技"
            lists[loopControlVar][7] = "人工智能"

    for key in {"环境保护及废弃资源综合利用","环保","资源利用","净水","新风","太阳能","排水","节能","新能源","风能"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "工业科技"
            lists[loopControlVar][7] = "环境保护及废弃资源综合利用"

    for key in {"金属制品机械和设备修理","设备维修","管材"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "工业科技"
            lists[loopControlVar][7] = "金属制品机械和设备修理"

    for key in {"铁路交通运输","高铁","铁路","地铁","轨道"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "交通运输仓储和邮政"
            lists[loopControlVar][7] = "铁路交通运输"

    for key in {"道路交通运输","汽车","卡车","车","汽摩","交通工程"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "交通运输仓储和邮政"
            lists[loopControlVar][7] = "道路交通运输"

    for key in {"水上交通运输","船","海事"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "交通运输仓储和邮政"
            lists[loopControlVar][7] = "水上交通运输"

    for key in {"航空航天","飞机","航空","航天","无人机"}:
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

    for key in {"计算机通信和其他电子设备","计算机","通信","智能硬件"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "信息传输软件和信息技术"
            lists[loopControlVar][7] = "计算机通信和其他电子设备"

    for key in {"电信广播电视和卫星传输服务","电信","广播","卫星"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "信息传输软件和信息技术"
            lists[loopControlVar][7] = "电信广播电视和卫星传输服务"

    for key in {"互联网和相关服务","互联网"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "信息传输软件和信息技术"
            lists[loopControlVar][7] = "互联网和相关服务"

    for key in {"软件和信息技术","软件","信息技术","数据","嵌入式系统","智能交通","智慧","物联网"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "信息传输软件和信息技术"
            lists[loopControlVar][7] = "软件和信息技术"

    for key in {"医药制造","药","生物技术","医学"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "医疗健康"
            lists[loopControlVar][7] = "医药制造"

    for key in {"医疗用品及器材","医疗器材","医疗用品","医疗器械","疫","口腔","眼科"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "医疗健康"
            lists[loopControlVar][7] = "医疗用品及器材"

    for key in {"护理及其他医疗健康服务","护理","医疗健康","康复","养老","老龄","健康产业","健康产品","营养","保健","健康","养生"}:
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

    for key in {"建筑装饰和其他建筑业","装修","家博会","涂料","供暖","全屋整装"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "房屋建筑装修及经营服务"
            lists[loopControlVar][7] = "建筑装饰和其他建筑业"

    for key in {"家装设计及家具","家具","家装","家居"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "房屋建筑装修及经营服务"
            lists[loopControlVar][7] = "家装设计及家具"

    for key in {"房地产","住宅"}:
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

    for key in {"旅行及相关服务","旅游","民宿","旅居"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "租赁和商务服务"
            lists[loopControlVar][7] = "旅行及相关服务"

    for key in {"安全保护服务","安全","安全保护","安防","安保","应急"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "租赁和商务服务"
            lists[loopControlVar][7] = "安全保护服务"

    for key in {"电子商务"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "租赁和商务服务"
            lists[loopControlVar][7] = "电子商务"

    for key in {"其他商务服务业","加盟","品牌","公益","环博会","丝绸之路"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "租赁和商务服务"
            lists[loopControlVar][7] = "其他商务服务业"

    for key in {"纺织面料服装及服饰","服饰","时装","服装","针织","纺织","纱线"}:
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

    for key in {"玩具","模型"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "日用消费品及居民服务"
            lists[loopControlVar][7] = "玩具"

    for key in {"化妆品卫生用品及美容美发服务","化妆品","卫生用品","美容","美发","美博会"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "日用消费品及居民服务"
            lists[loopControlVar][7] = "化妆品卫生用品及美容美发服务"

    for key in {"钟表眼镜","钟表","钟","眼镜"}:
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

    for key in {"婚庆设施及服务","婚","喜庆用品"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "日用消费品及居民服务"
            lists[loopControlVar][7] = "婚庆设施及服务"

    for key in {"其他产品及服务","消费品","商品","交易","特卖","展销","贸易","婴童"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "日用消费品及居民服务"
            lists[loopControlVar][7] = "其他产品及服务"

    for key in {"教育"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "教育"
            lists[loopControlVar][7] = "教育"

    for key in {"教育机构及培训","幼儿园","幼教"}:
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

    for key in {"广播电视电影和影视制作","影视","电影","视频","灯光音响","视听"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "文化体育和娱乐"
            lists[loopControlVar][7] = "广播电视电影和影视制作"

    for key in {"文化艺术","艺术","佛文化","石博会","乐器","佛事"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "文化体育和娱乐"
            lists[loopControlVar][7] = "文化艺术"

    for key in {"体育","运动","马术","健身","钓鱼","钓具"}:
        if key in lists[loopControlVar][0]:
            lists[loopControlVar][6] = "文化体育和娱乐"
            lists[loopControlVar][7] = "体育"

    for key in {"娱乐","动漫","游戏","酷狗蘑菇"}:
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
    worksheet.write(index, 27, "否")
    worksheet.write(index, 28, lists[index][8])
    worksheet.write(index, 20, lists[index][9])
    

workbook.save('xxxx.xls') 