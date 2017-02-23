# -*- coding: utf-8 -*-  
import whyspider  
  
# 初始化爬虫对象  
my_spider = whyspider.WhySpider()  
  
# 模拟GET操作  
print my_spider.send_get('http://3.apitool.sinaapp.com/?why=GetString2333')    
  
# 模拟POST操作  
print my_spider.send_post('http://3.apitool.sinaapp.com/','why=PostString2333')    
  
# 模拟GET操作  
print my_spider.send_get('http://www.baidu.com/')    
  
# 切换到手机模式  
my_spider.set_mobile()  
  
# 模拟GET操作  
print my_spider.send_get('http://www.baidu.com/') 