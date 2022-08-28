import re
import urllib.error,urllib.request
import time
import xlwt

'''
https://www.cjlu.edu.cn/index/xww/jlyw1.htm
https://www.cjlu.edu.cn/index/xww/jlyw1/151.htm
https://www.cjlu.edu.cn/index/xww/jlyw1/150.htm
'''



def main():

    url1='https://www.cjlu.edu.cn/index/xww/jlyw1/'
    url3='.htm'

    savepath='中国计量大学新闻.xls'

    html=''
    for i in range(1,153):
        url2=153-i
        url=url1+str(url2)+url3
        if url=='https://www.cjlu.edu.cn/index/xww/jlyw1/152.htm':
            url='https://www.cjlu.edu.cn/index/xww/jlyw1.htm'
        html=html+ask(url)
        time.sleep(5)

    # print(html)
    datalist=request_data(html)
    save_data(datalist,savepath)

    print(datalist)


def ask(url):
    header={
        "User-Agent": ""
    }
    html=''
    request=urllib.request.Request(headers=header,url=url)

    try:
        respnse=urllib.request.urlopen(request)

        html=respnse.read().decode('utf-8')

    except urllib.error.URLError as e:
        if hasatter(e,'code'):
            print(e.code)
        if hasatter(e,'reason'):
            print(e.reason)

    return html


def request_data(html):
    datalist=[]
    #链接
    link=[]
    links=re.findall(r'../../info/1100/(.*?)"',str(html))
    for i in links:
        li='https://www.cjlu.edu.cn/info/1100/'+i
        link.append(li)

    link.insert(454,'https://www.cjlu.edu.cn/info/1099/21644.htm')

    #消息名称
    new=re.compile('target="_blank" style="font-size:14px(.*?)</a>',re.S)

    new_find=re.findall(new,str(html))

    # news=[x.strip() for x in new_find]
    news=[]

    for n in new_find:
        n=n.replace('">','')
        n=n.replace(';font-weight:bold;color:#000000;','')
        n=n.strip()

        news.append(n)



    #时间
    new_time=re.findall('<span>(.*?)</span>',str(html))

    a=0
    for i in news:
        data=[]
        data.append(news[a])
        data.append(new_time[a])
        data.append(link[a])
        print(news[a],new_time[a],link[a])
        datalist.append(data)
        a+=1


    return datalist



def save_data(datalist,savepath):
    book=xlwt.Workbook(encoding='utf-8',style_compression=0)

    sheet=book.add_sheet('中国计量大学新闻',cell_overwrite_ok=True)

    qres=[]
    for y in range(1, 32):
        qre = y * 150
        qres.append(qre)

    c=0
    a=0
    col=('消息','时间','链接')
    for i in range(0,31):
        for y in range(0,3):
            sheet.write(0,y+(3*i),'第'+str(i+1)+'列'+col[y])
    for data in datalist:

        for i in qres:
            if a==i:
                c+=1
        a+=1

        for i in range(0,3):
            sheet.write(a-(150*c),i+(3*c),data[i])
    book.save(savepath)






if __name__ == '__main__':
    main()
