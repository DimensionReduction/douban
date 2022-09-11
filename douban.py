# 导入库
import os
import requests
import parsel
import xlsxwriter

# 请求头
headers={'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/99.0.4844.84 Safari/537.36'}

# 创建excel，设置列宽
wb=xlsxwriter.Workbook('豆瓣电影.xlsx')
ws=wb.add_worksheet('豆瓣电影海报')
ws.set_column('A:A',7)
ws.set_column('B:B',30)
ws.set_column(2,3,50)
ws.set_column('G:G',30)

# 标题行
headings=['海报','名称','导演','主演','年份','国家','类型']
# 设置excel风格
ws.set_tab_color('red')
head_format=wb.add_format({'bold':1,'fg_color':'cyan','align':'center','font_name':u'微软雅黑','valign':'vcenter'})
cell_format=wb.add_format({'bold':0,'align':'center','font_name':u'微软雅黑','valign':'vcenter'})
ws.write_row('A1',headings,head_format)

# 创建空列表
j=0
actor_1=[]
actor_2=[]
year=[]
country=[]
movie_type=[]

# 发送请求，获取响应，遍历豆瓣电影信息
for i in range(0,250,25):
    url='https://movie.douban.com/top250?start='+str(i)
    response=requests.get(url,headers=headers)
    response.encoding=response.apparent_encoding
    selector=parsel.Selector(response.text)
    lis=selector.css('#content>div>div.article>ol>li>div>div.pic>a>img::attr(src)').getall()
    title=selector.css('#content>div>div.article>ol>li>div>div.info>div.hd>a>span:nth-child(1)::text').getall()
    director_info=selector.xpath('//*[@id="content"]/div/div[1]/ol/li/div/div[2]/div[2]/p[1]/text()[1]').getall()
    movie_info = selector.xpath('//*[@id="content"]/div/div[1]/ol/li/div/div[2]/div[2]/p[1]/text()[2]').getall()

    # 提取并清洗导演、主演标签信息
    for director in director_info:
        director=director.strip().replace('\xa0\xa0\xa0',' ').replace('...','').replace('导演: ','').split('主演: ')
        if len(director)>1:
            actor_1.append(director[0])
            actor_2.append(director[1])
        else:
            actor_1.append(director[0])
            actor_2.append('None')

    # 提取并清洗电影年份、国家、类型标签信息
    for detail in movie_info:
        detail=detail.strip().split('\xa0/\xa0')
        year.append(detail[0])
        country.append(detail[1])
        movie_type.append(detail[2])

    # 下载电影海报信息到本地
    if os.path.exists('./图片')==False:
        os.mkdir('./图片')
    for n in range(len(lis)):
        img=requests.get(lis[n]).content
        with open(f'./图片/{title[n]}.jpg','wb') as f:
            f.write(img)

    # 将获取到的各种标签信息写入excel
    for k in range(len(lis)):
        ws.set_row(k+1+j*25,60)
        ws.insert_image('A'+str(k+2+j*25),f'./图片/{title[k]}.jpg',{'x_scale':0.2,'y_scale':0.2})
        ws.write(k+1+j*25,1,title[k],cell_format)
        ws.write(k+1+j*25,2,actor_1[k+j*25],cell_format)
        ws.write(k+1+j*25,3,actor_2[k+j*25],cell_format)
        ws.write(k+1+j*25,4,year[k+j*25],cell_format)
        ws.write(k+1+j*25,5,country[k+j*25],cell_format)
        ws.write(k+1+j*25,6,movie_type[k+j*25],cell_format)
    j+=1
wb.close()
