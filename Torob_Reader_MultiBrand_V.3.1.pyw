import tkinter as tk
import requests
from bs4 import BeautifulSoup
import csv
import datetime as dt
import time 
import random
import pandas.io.excel as pd
import os

def openFile(fileName):
    os.startfile(f"{fileName}")
   
window= tk.Tk()
window.title('Torob reader')
canvas1 = tk.Canvas(window, width = 400, height = 350)
canvas1.pack()

brandsDic={'htc':1,
'nokia':2,
'sony':3,
'samsung':5,
'lg':6,
'glx':8,
'microsoft':13,
'apple':14,
'alcatel':24,
'asus':29,
'huawei':31,
'tp-link':32,
'dimo':38,
'fly':40,
'motorola':42,
'lenovo':43,
'blackberry':66,
'cat':67,
'zte':71,
'xiaomi':102,
'meizu':111,
'techno':220,
'hyundai':268,
'oneplus':285,
'google':299,
'blu':303,
'doogee':581,
'oppo':667,
'orod':677,
'marshal':757,
'energizer':3446,
'wiwu':3588,
'venus':4595,
'gplus':4616,
'smart':5615,
'ken xin da':6996,
'realme':7019,
'leagoo':10365,
'inoi':10465,
'honor':10661,
'dox':12630,
'concord plus':13233,
'lava':13573,
'infinix':14795,
'hope':17639,
'kgtel':53492,
'invens':53586,
'odscn':53589,
'multiphone':53617,
'vfone':53745,
'hisense':15214}

def timeToText(p):
    m=p//60
    s=p%60
    t1=''
    if m>0:
        t1=f' {p//60} minute'
        if m>1: t1+='s'
        if s>0: t1+=' and'
    t2=''
    if s>0 : 
        t2=f' {p%60} second'
        if s>1: t2+='s'
    
    return t1+t2

def arToEnNum(st):
   lis=[]
   dic = { 
        '۰':'0', 
        '۱':'1', 
        '۲':'2', 
        '۳':'3', 
        '۴':'4', 
        '۵':'5', 
        '۶':'6', 
        '۷':'7', 
        '۸':'8', 
        '۹':'9', 
        '٫':','
       }
   for char in st:
          if char in dic:
                 lis.append(dic[char])
          else:
                 lis.append(char)
   return "".join(lis)

def getData (linkList):
    dt1=dt.datetime.now()
    label_counter = tk.Label(window, text='0')
    canvas1.create_window(200, 200,window=label_counter)
    items1=[]
    items=[]
    j=0
    for link in linkList:
        if True:

            URL1 = link
            page1 = requests.get(URL1)
            items1=[]
            soup1 = BeautifulSoup(page1.content, "html.parser")
            table=soup1.find("div",class_="cards")
            divs1 = table.findChildren("div",recursive=False)
            for div in divs1:
                item={}
                model_name_elem = div.find("h2", class_="product-name")
                if model_name_elem:
                    item["model"] = arToEnNum(model_name_elem.text)
                    item["spec"]=""
                    if len(model_name_elem)>1:
                        item["spec"]=arToEnNum(model_name_elem[1])
                link_elem = div.find("a")
                if link_elem:
                    item["link"] = r'https://torob.com'+link_elem["href"]
                min_price_elem = div.find("div", class_="product-price-text")
                if min_price_elem:
                    item["price"] = arToEnNum(min_price_elem.text)
                shops_elem = div.find("div", class_="shops")
                if shops_elem:
                    item["shops"] = arToEnNum(shops_elem.text)
                item["DateTime"]=dt1.strftime('%Y/%m/%d %H:%M:%S')
                items1.append(item)
                
            for index1,item1 in enumerate(items1,start=0):
                if index1==5:
                    break
                else:
                    label_counter['text']=str(round((j/(5*len(linkList)))*100))+' %'
                    label_counter.update()
                    j+=1
                    URL = item1["link"]
                    page = requests.get(URL)
                    soup = BeautifulSoup(page.content, "html.parser")
                    specSection=soup.find("div",class_="specs-content")
                    specDivs=specSection.findChildren()
                    Specs=""
                    for div in specDivs:
                        #print(div)
                        title=div.find("div",class_="detail-title")
                        value=div.find("div",class_="detail-value")
                        if title:
                            Specs+=title.text+": "+value.text+"|"
                    divs = soup.find_all("div",class_="shop-card")
                    for div in divs:
                        item={}
                        if div:
                            #print(div.text)
                            shop=div.find("div", class_="name-wrapper")
                            if shop:
                                item["Store"]=shop.text
                                item["Spec Details"]=Specs
                                print(shop.text)
                            city=div.find("a",class_="city-name")
                            if city:
                                item["City"]=city.text
                                print(city.text)
                            adv=div.find("div",class_="click_vijhe")
                            if adv:
                                item["Ads"]=adv.text
                                print(adv.text)
                            model=div.find("div",class_="product-info").find("a",class_="seller-element").find("div",class_="product-name")
                            if model:
                                item["Description"]=model.text
                                print(model.text)
                            price=div.find("a",class_="price")
                            if price: 
                                item["Price"]=arToEnNum(price.text).strip('تومان').strip()
                            item["Model"]=item1["model"]
                            item["Specification"]=item1["spec"]
                            item["Min price"]=item1["price"]
                            if item1["price"]!="ناموجود":
                                item["Min price"]=item1["price"].strip('از').strip('تومان').strip()
                            item["In-stock stores"]=item1["shops"]
                            lastchange=div.find("div",class_="last_price_change_date")
                            if lastchange: 
                                item["Last price change"]=arToEnNum(lastchange.text.lstrip("آخرین تغییر قیمت فروشگاه: "))
                            item["DateTime"]=item1["DateTime"]
                            items.append(item)
                            print(item)
                    time.sleep(random.randint(5,9))
                    # time.sleep(random.randint(9,12))
            
    label_counter['text']=str(100)+' %'
    label_counter.update()
    dt2=dt.datetime.now()
    elapsedTime=dt2-dt1
    filename=dt.datetime.today().strftime("%Y%m%d-%H%M%S")+".csv"
    with open(filename, 'w', encoding="utf-8-sig", newline='') as f:
        fieldnames =["Model","Specification","Store","Price","Min price","Ads","City","Last price change","In-stock stores","Description","DateTime","Spec Details"]
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(items)
        if button1['state']=='disabled':
            button1.pack_forget()
            button2 = tk.Button(text='Click to see the result', command=lambda:openFile(filename), bg='green',fg='white',width=20)
            button2.pack()
            canvas1.create_window(200, 160, window=button2)
        label1 = tk.Label(window, text= 'Finished!', fg='red', font=('helvetica', 11, 'normal'))
        canvas1.create_window(200, 200, window=label1)
        label11 = tk.Label(window, text= 'The result file have been saved as: \"'+filename+'\" \n(Elapsed time: '+timeToText(elapsedTime.seconds)+')')
        canvas1.create_window(200, 240, window=label11)

        
def getLink(brand,minPrice,maxPrice,model):
    category='گوشی-موبایل-mobile'
    brand=str(brand)
    minPrice=str(minPrice)
    maxPrice=str(maxPrice)
    model=str(model)
    p1='https://torob.com/browse/94/'+category
    p2='/b/'+str(brandsDic[brand.lower()])+'/'+brand+'/?'
    p3='stock_status=new&available=true'
    p4='&price__gt='+str(minPrice)
    p5='&price__lt='+str(maxPrice)
    p6=''
    if model is not None or model!='':
        p6='&q='+model
    linkString=p1+p2+p3+p4+p5+p6
    return linkString

        
def buildLinkList():
    excel=pd.ExcelFile('FilterParams.xlsx')
    myList=pd.read_excel(excel,'Sheet1')
    linkList=[]
    for i in myList.index:
        link=getLink(myList['Brand'][i],myList['Min_price'][i],myList['Max_price'][i],myList['Model'][i])
        linkList.append(link)
    return linkList

def execute():
    if button1['state']=='normal':
        button1['state']='disabled'
        button1['text']='Please wait...'
        button1['bg']='light grey'
    linkList=buildLinkList()
    getData(linkList)
    
button1 = tk.Button(text='Start', command=execute, bg='brown',fg='white',width=20)
canvas1.create_window(200, 160, window=button1)
label0=tk.Label(window,text='\nUSER GUIDE: \n1- Build the \"FilterParams.xlsx\" file according to the sample format\n     and save it in root folder of the app.\n2- Make sure your computer has an active internet connection.\n3- Click the \"Start\" button.\n',justify='left')
canvas1.create_window(200,70,window=label0)
label2=tk.Label(window,text='h.khajei@gmail.com\nVersion: 3.1', fg='gray',font=('helvetica', 8, 'normal'))
canvas1.create_window(200,320,window=label2)
window.mainloop()
