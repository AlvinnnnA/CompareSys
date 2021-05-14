import time  #载入计时需要的包
loadstart=time.process_time()
import easygui as g  #载入GUI包
print("正在载入运行包")
import synonyms,os,sys,docx  #载入主要包和一些需要的包
from selenium import webdriver  #载入浏览器操作包
from webdriver_manager.microsoft import EdgeChromiumDriverManager  #载入浏览器自动控制驱动
loadend=time.process_time()
print("包载入用时%6.3f'秒" %(loadend-loadstart))
def wordGet(num): #word内容提取函数
    doc=docx.Document(g.fileopenbox(msg="请选择要比对的第"+str(num)+"份Word文档",filetypes=["*.docx"]))
    content=[]
    for i in doc.paragraphs:  #遍历全部段落
       contentstr=i.text 
       content.append(contentstr)
    return str(content)
def isMatch(a,b):  #网站匹配判断函数
    match=True
    if len(b)>len(a):
        max=b
        min=a
    else:
        max=a
        min=b
    i=0
    while i < len(min):
        if min[i]==max[i]:
            i+=1
        else:
            match=False
            break
    else:
        pass
    return match  #之前作业拿出来的函数懒得写注释了
def chooseMode():  #模式选择GUI 毫无必要的弱智写法
    choice=g.buttonbox(msg='请选择工作模式', title='选择工作模式', choices=('文本比对文本', '文本网络比对', '即时输入比对','即时输入与文件比对','文件夹与单文件比对(仅Word文件)','文件夹交叉比对','退出'), image=None)
    mode=0
    if choice=="文本比对文本":
        mode=1
    elif choice=="文本网络比对":
        mode=2
    elif choice=="即时输入比对":
        mode=3
    elif choice=="即时输入与文件比对":
        mode=4
    elif choice=='文件夹与单文件比对(仅Word文件)':
        mode=5
    elif choice=='文件夹交叉比对':
        mode=6
    return mode  #输出mode为一个整数
def webCompare(SourceDocString1):  #网络对比
    keywords =synonyms.keywords(SourceDocString1,topK=3)  #关键词列表 提取三个关键词搜索
    print("文本关键词为",keywords)
    site="m.51test.net"  #搜索站点 暂定"无忧考网"
    browser = webdriver.Edge(EdgeChromiumDriverManager().install())  #使用edge浏览器
    browser.get("https://cn.bing.com/search?q="+" "+keywords[0]+" "+keywords[1]+" "+keywords[2]+" site:"+site)  #必应搜索三个关键词
    search_links=[] #储存搜索结果的链接list
    resultcontent=[]  #储存搜索结果内容list
    r=[]  #结果数字list
    rstelmt=[]  #result element list
    results=dict() #输出用！字典 result带s！！
    key_list2=[]
    value_list2=[]    
    site_contentdict={}
    output={}
    sitestr=""
    result=browser.find_elements_by_css_selector("h2>a")  #提取搜索结果项源码
    for i in result[0:8]:  #前五搜索结果源码中提取链接
        if isMatch("https://"+site,i.get_attribute("href"))==True:  #排除非该网站的项
            search_links.append(i.get_attribute("href"))  #提取链接合并到list
    for j in search_links: #打开结果链接并提取内容
        browser.get(j)
        rstelmt=browser.find_elements_by_css_selector("div#content-txt>p")  #ResultElement
        for k in rstelmt:  #只有当前网页才能提取文本，故需要在循环中加循环嵌套
            if len(k.get_attribute("textContent"))>0:  #排除空内容
                sitestr+=k.get_attribute("textContent")
        resultcontent.append(sitestr)  #单网站内容集
        site_contentdict[j]=sitestr
    browser.quit()  #关闭浏览器
    for l in resultcontent[0:5]:  #取的数字太小
        rel=synonyms.compare(SourceDocString1,l, seg=True,ignore=True)
        print(rel)
        results[rel]=l #result字典中存储link对应的text（key：list）
        print(results)
        r.append(rel)  #语句比对
    ressorted=sorted(results.items(),key=lambda x:x[0],reverse=True) 
    for key,value in site_contentdict.items( ): #创造主字典内反向查找的条件
        key_list2.append(key)
        value_list2.append(value)
    print(ressorted)
    for m in ressorted[0:3]:
        val=m[1]  #查找：在results中取前三位的key值
        output[m[0]]=key_list2[value_list2.index(val)]  #将rel：link加入output字典之中
    return output  #要改return值
def compareText():   #主体判断与执行函数 (要做的：+文件与文件夹比对  +大量文件互相交叉比对12)
    mode=chooseMode()   #选择模式 
    if mode==0:
        sys.exit(0)   #点取消就退出程序
    elif mode==1:   #文件比对
        choice=g.buttonbox(msg='请选择文件格式', title='选择文件格式', choices=('TXT', 'Word(仅支持docx格式)'), image=None)
        if choice=='TXT':
            SourceRoute1=g.fileopenbox(msg="请选择需比对的第一份文件",filetypes = ["*.txt"])  #找到比对源文件，输出路径到SourceRoute1
            Source1=open(SourceRoute1,encoding='utf-8')  #打开比对源文件1
            try:
               SourceDocString1 = Source1.read()  #将txt内容输出到字符串SourceDocString1
            finally:  
               Source1.close()
            SourceRoute2=g.fileopenbox(msg="请选择需比对的第二份文件",filetypes = ["*.txt"])  #找到比对源文件，输出路径到SourceRoute2
            Source2=open(SourceRoute2,encoding='utf-8')  #打开比对源文件2
            try:
               SourceDocString2 = Source2.read()  #将txt内容输出到字符串SourceDocString2
            finally:  
               Source2.close()
        elif choice=='Word(仅支持docx格式)':
            SourceDocString1=wordGet(1)
            SourceDocString2=wordGet(2)
        start =time.process_time() #计时
        print("正在进行比对")
        r = synonyms.compare(SourceDocString1, SourceDocString2, seg=True,ignore=True) #语句比对
        end=time.process_time()
        Sim=str(r)  #将r值number转换为string
        g.msgbox(msg="输入的语句相似度为"+Sim, title='比对结果', ok_button='返回')
    elif mode==2:  #网络比对 
        output_list=[]
        output_list2=[]
        choice=g.buttonbox(msg='请选择文本源', title='选择文本源', choices=('输入文本', '文本文件'), image=None)  #选择文本比对/文件比对
        if choice=='输入文本':
            sen1=g.enterbox(msg='请输入需比对的语句', title='输入语句',  strip=True, image=None)
            start =time.process_time()  #计时
            output=webCompare(sen1) #比对函数
            end=time.process_time()
            for key,value in output.items( ): #创造输出的条件
                output_list.append(key)
                output_list2.append(value)
            g.msgbox(msg="相似度最高，为"+str(output_list[0])+"的网络链接为："+output_list2[0]+ "\n"+"相似度第二高，为"+str(output_list[1])+"的网络链接为："+output_list2[1]+ "\n"+"相似度第三高，为"+str(output_list[2])+"的网络链接为："+output_list2[2])
        elif choice=='文本文件':
            SourceRoute1=g.fileopenbox(msg="请选择需比对的文本文件",filetypes = ["*.txt"])  #找到比对源文件，输出路径到SourceRoute1
            Source1=open(SourceRoute1,encoding='utf-8')  #打开比对源文件1
            try:
                SourceDocString1 = Source1.read()  #将txt内容输出到字符串SourceDocString1
            finally:  
                Source1.close()
            start =time.process_time()  #计时
            output=webCompare(SourceDocString1)
            end=time.process_time()
            for key,value in output.items( ): #创造输出的条件
                output_list.append(key)
                output_list2.append(value)
            print(output)
            g.msgbox(msg="相似度最高，为"+str(output_list[0])+"的网络链接为："+output_list2[0]+ "\n"+"相似度第二高，为"+str(output_list[1])+"的网络链接为："+output_list2[1]+ "\n"+"相似度第三高，为"+str(output_list[2])+"的网络链接为："+output_list2[2])
    elif mode==3: #即时输入比对
        sen1=g.enterbox(msg='请输入需比对的第一语句', title='输入语句',  strip=True, image=None)
        sen2=g.enterbox(msg='请输入需比对的第二语句', title='输入语句',  strip=True, image=None)
        start =time.process_time()  #计时
        print("正在进行比对")
        r = synonyms.compare(sen1, sen2, seg=True,ignore=True)  #语句比对
        Sim=str(r)  #将r值number转换为string
        end=time.process_time()
        g.msgbox(msg="输入的语句相似度为"+Sim, title='比对结果', ok_button='返回')
    elif mode==4: #即时输入与文件比对
        SourceRoute1=g.fileopenbox(msg="请选择需比对的第一份文件",filetypes = ["*.txt"])  #找到比对源文件，输出路径到SourceRoute1
        Source1=open(SourceRoute1,encoding='utf-8')  #打开比对源文件1
        try:
            SourceDocString1 = Source1.read( )  #将txt内容输出到字符串SourceDocString1
        finally:  
            Source1.close( )
        sen1=g.enterbox(msg='请输入需比对的第一语句', title='输入语句',  strip=True, image=None)
        start =time.process_time() #计时
        print("正在进行比对")
        r = synonyms.compare(sen1,SourceDocString1, seg=True,ignore=True)  #语句比对
        Sim=str(r)  #将r值number转换为string
        end=time.process_time()
        g.msgbox(msg="输入的语句相似度为"+Sim, title='比对结果', ok_button='返回')
    elif mode==5:#文件+文件夹比对(Word(.docx) only)
        SourceDocString1=wordGet(1) #取文件
        Dir=g.diropenbox(msg="请选择需比对的文件目录",title="请选择需比对的文件目录") #取文件夹
        FileDir=os.listdir(Dir) #遍历文件
        filecontent=dict() #建立内容字典
        reldict=dict() #建立相似度字典
        key_list2=[]
        value_list2=[]
        output_list=[]
        output_list2=[]
        output=dict()
        for file in FileDir:
            fildir=Dir+"\\"+file
            if os.path.splitext(file)[1]=='.docx':#判断文件扩展名
               doc=docx.Document(fildir)
               content=''
               for i in doc.paragraphs:  #遍历全部段落
                  contentstr=i.text
                  if len(contentstr)>0: #排除空段
                      content+=contentstr #content字符串保存内容
               filecontent[file]=content 
            else:
                pass
        start =time.process_time()
        for filecon in filecontent.values(): #比对代码+反向查询+排序
            rel=synonyms.compare(filecon,SourceDocString1,seg=True,ignore=True)
            reldict[rel]=filecon #创造子字典（相似度：比对内容）
        ressorted=sorted(reldict.items(),key=lambda x:x[0],reverse=True)
        for key,value in filecontent.items( ): #创造主字典内反向查找的条件
            key_list2.append(key)
            value_list2.append(value)    
        for m in ressorted[0:3]:#beg:end=beg->(end-1)!!!!!注意数字含义！！
            key=m[1]  #查找：在results中取前三位的key值
            output[m[0]]=key_list2[value_list2.index(key)]
        for key,value in output.items( ): #创造输出的条件
            output_list.append(key)
            output_list2.append(value)
        end=time.process_time()
        g.msgbox(msg="相似度最高，为"+str(output_list[0])+"的文件为："+output_list2[0]+ "\n"+"相似度第二高，为"+str(output_list[1])+"的文件为："+output_list2[1]+ "\n"+"相似度第三高，为"+str(output_list[2])+"的文件为："+output_list2[2])
    elif mode==6:#文件夹交叉比对
        Dir=g.diropenbox(msg="请选择需比对的文件目录",title="请选择需比对的文件目录") #取文件夹
        FileDir=os.listdir(Dir) #遍历文件
        FullDir=[]
        RelDict={}
        output={}
        output_list=[]
        output_list2=[]
        for file in FileDir:#完整路径list
           fildir=Dir+"\\"+file
           FullDir.append(fildir)
        start =time.process_time()
        for sourcefile in FullDir[0:len(FullDir)-1]:#对除最后一个以外的每一个文件进行操作
               srcdoc=docx.Document(sourcefile)
               srccontent=''
               for i in srcdoc.paragraphs:  #遍历全部段落
                  srccontentstr=i.text
                  if len(srccontentstr)>0: #排除空段
                      srccontent+=srccontentstr #content字符串保存内容
               for targetfile in FullDir[FullDir.index(sourcefile)+1:]:#对该文件与其之后的文件进行比对
                  tgtdoc=docx.Document(targetfile)
                  tgtcontent=''
                  for i in tgtdoc.paragraphs:  #遍历全部段落
                     tgtcontentstr=i.text
                     if len(tgtcontentstr)>0: #排除空段
                        tgtcontent+=tgtcontentstr #content字符串保存内容
                  sim=synonyms.compare(srccontent,tgtcontent,seg=True,ignore=True)
                  RelDict[sim]=os.path.basename(targetfile)+"和"+os.path.basename(sourcefile)
        ressorted=sorted(RelDict.items(),key=lambda x:x[0],reverse=True)
        end=time.process_time()
        for m in ressorted[0:3]:#beg:end=beg->(end-1)!!!!!注意数字含义！！
            output[m[0]]=m[1]
        for key,value in output.items( ): #创造输出的条件
            output_list.append(key)
            output_list2.append(value)
        g.msgbox(msg="相似度最高，为"+str(output_list[0])+"的文件为："+output_list2[0]+ "\n"+"相似度第二高，为"+str(output_list[1])+"的文件为："+output_list2[1]+ "\n"+"相似度第三高，为"+str(output_list[2])+"的文件为："+output_list2[2])
    print("本次运行用时%6.3f'秒" %(end - start),sep='')
while True:
    compareText()
