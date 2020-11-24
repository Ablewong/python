# encoding:utf-8

__author__ = 'Able'
'''
systeminfo.py 的升级版本，主要用来把公司电脑信息收集分类写到一个excel中。
执行前需要把公司每个电脑的dxdiag信息采集到D:/Work/computercollect/20201119目录(当然这个可以改)，
dxdiag信息收集可以执行‘ start dxdiag.exe ’内容的BAT文件。然后保存电脑使用人员的名字(这个很重要，我用这名字来写入到EXCEL)给到你那个目录。

'''
import os
import os.path
from pathlib import Path
import xlwings as xw


info_dict={}  #把所有电脑信息存到这个字典中。
count=0

def walk_dir(dir,topdown=False):
    for root, dirs, files in os.walk(dir, topdown):
        
        for name in files:
            countdict={'Model':''}      #定义个空硬盘信息字典
            catinfo = open(os.path.join(root,name), 'r')    #打开个人文件
            hardwarename = ['   Processor', 'OS Memory', 'File System', 'Monitor Model', '          Card name',]
            #fileinfo.write(os.path.join(name).strip() + '\n')
            for readline in catinfo:
                for hard in hardwarename:
                    if hard == '   Processor' and hard in readline:  #这段是把CPU简化，只需知道型号就行。
                        cpu = (readline.strip().split(":")[-1]).strip()     #把CPU提取出来
                        if 'Intel' in cpu:
                            countdict[hard.strip()] =  ' '.join((cpu.split('CPU')[0]).split(' '))
                        else:
                            countdict[hard.strip()] = ' '.join(cpu.split(' ')[0:4])
                    elif hard == 'File System' and hard in readline:  #获取下一行硬盘信息
                        model=(next(catinfo).strip().split(":")[-1]).strip()    
                        if not model in countdict['Model']:     #分区的情况会获取到重复的硬盘信息，这里是去掉重复信息。
                            countdict['Model'] = countdict['Model'] + '|' + model
                            #countdict['Model'] = (next(catinfo).strip().split(":")[-1]).strip()
                    elif hard == 'OS Memory'and hard in readline:
                        mem = (readline.strip().split(":")[-1]).strip()
                        countdict[hard.strip()] = str(int(float(mem.split('MB')[0])/1000)) + 'GB'  #把MB转换成GB，为什么除1000呢，是因为获取的内存不一定是全内存，有可能有些被占用了。
                    elif hard in readline:
                        if hard in countdict:
                            countdict[hard.strip()] = (countdict[hard] +","+ readline.strip().split(":")[-1]).strip()
                        else:
                            countdict[hard.strip()] = (readline.strip().split(":")[-1]).strip()



            info_dict[str(name.split(".txt")[0])] = countdict
            #count+1
            catinfo.close   #个人文件关闭

                        

def wirte_xls(comdict):
    count1=1    
    app = xw.App(visible=True,add_book=False)
    #新建工作簿 (如果不接下一条代码的话，Excel只会一闪而过，卖个萌就走了）
    wb = app.books.add()
    sht = wb.sheets[0]
    sht.range('a1').value = ["序列","使用人员","部门","CPU","内存","硬盘","显示器","显卡","备注"]
    for keys in comdict:
        count1+=1
        print(count1,keys)
        
        sht.range('A'+str(count1)).value = [count1,keys,'',comdict[keys]['Processor'],comdict[keys]['OS Memory'],comdict[keys]['Model'],comdict[keys]['Monitor Model'],comdict[keys]['Card name'],'']   #把每个人的配置写到excle里。
        #第二个空为部门，部门信息需要自己定义，每个公司不一样，就不填上去了。
    wb.save("computercollect2020v5.xls")  #保存excel名字
    wb.close()      #关闭sh
    app.quit()      #退出


def main():
    dir = 'D:/Work/computercollect/20201119'    #这里注意要用linux路径写法/，而不是\
    
    walk_dir(dir)       #此函数用来处理路径下的个信电脑信息，再遍历打开获取到需要采集的信息，保存为一个字典。
    
    #print(info_dict)
    wirte_xls(info_dict)        #此函数用来把之前的字典遍历写到EXCLE中


main()
