# -*- coding: utf-8 -*-  
import whyspider  
  
# ��ʼ���������  
my_spider = whyspider.WhySpider()  
  
# ģ��GET����  
print my_spider.send_get('http://3.apitool.sinaapp.com/?why=GetString2333')    
  
# ģ��POST����  
print my_spider.send_post('http://3.apitool.sinaapp.com/','why=PostString2333')    
  
# ģ��GET����  
print my_spider.send_get('http://www.baidu.com/')    
  
# �л����ֻ�ģʽ  
my_spider.set_mobile()  
  
# ģ��GET����  
print my_spider.send_get('http://www.baidu.com/') 