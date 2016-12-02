#!/usr/bin/env python
# -*- coding: utf-8 -*-

'''
 Copyright 2016 MaiXian Inc
 Description: this script help HR send mail to all staffs.
 Dependency: xlrd, Install:
 pip install xlrd

 Written by Able Huang 2016-11-12
 refer: http://code.activestate.com/recipes/578150-sending-non-ascii-emails-from-python-3/
'''

import sys
import smtplib
from email.header import Header
from email.mime.text import MIMEText
from email.utils import formataddr
from datetime import datetime
import xlrd
import time

#importlib.reload(sys)
#sys.setdefaultencoding('utf-8')

mail_host = 'smtp.126.com' #发送邮件的smtp地址
mail_user = 'hui@126.com' # 发送通知邮件的用户名
mail_pass = 'XXX' # 用户的密码

sender_name = 'Able' #发邮件者姓名
sender_addr = 'huibeyond@126.com' # 发邮件人的邮箱地址
subject = '2016年%s月买鲜网薪水发放通知单' %(int(time.strftime('%m'))-1) # ***邮件标题***
title = subject

html_template = """
<html>
<head>
<meta http-equiv="content-type" content="text/html;charset=utf-8" />
</head>
<body>
<h2 align="center">%s</h2>
<table border="1", cellpadding="2" cellspacing="0">
    <thead>
        <tr>
            <th>姓名</th>
            <th>部门</th>
            <th>出勤工时</th>
            <th>岗位津贴</th>
            <th>基本工资</th>
            <th>绩效工资</th>
            <th>加班补贴</th>
            <th>津贴</th>
            <th>请假工时</th>
            <th>请假工资</th>
            <th>乐捐</th>
            <th>应发合计</th>
            <th>公司承担社保</th>
            <th>公司承担住房公积金</th>
            <th>水电费</th>
            <th>税前工资</th>
            <th>个税</th>
            <th>实发工资</th>
        </tr>
    </thead>
    <tbody>
        <tr>
            <td> %s </td>
            <td> %s </td>
            <td> %s </td>
            <td> %s </td>
            <td> %s </td>
            <td> %s </td>
            <td> %s </td>
            <td> %s </td>
            <td> %s </td>
            <td> %s </td>
            <td> %s </td>
            <td> %s </td>
            <td> %s </td>
            <td> %s </td>
            <td> %s </td>
            <td> %s </td>
            <td> %s </td>
            <td> %s </td>
        </tr>
    </tbody>
</table>

<pre>
 <!--说明：
 1、以上薪水将会转入您提供给公司的银行账户中，如银行账户有更改请及时通知人力资源部。
 2、员工薪水信息属公司机密，请妥善保管，泄漏者将按公司有关制度处理。
 3、如对薪水支付数额有异议的，请于一周内以邮件的形式向人力资源部提出。
 -->
财务部
 %s


  ----------------------------------------------

              %s   财务
</pre>
</body>
</html>
"""

def write_mail_content(title, name, deparment,Attendance_hours,
                post_salary,base_salary,pay_for_performance,overtime_subsidies,post_allowance,
                sick_hours,sick_salary,donastion,total,company_pension_commitment,company_housing_fund_commitment,
                water_rate,Gross_salary,personal_income_tax,real_salary,today,sender_name):

    content = html_template % (title, name, deparment, Attendance_hours,
                post_salary, base_salary, pay_for_performance, overtime_subsidies, post_allowance,
                sick_hours, sick_salary, donastion, total, company_pension_commitment,
                company_housing_fund_commitment,
                water_rate, Gross_salary, personal_income_tax, real_salary, today, sender_name)

    return content

def send_mail(msg, sender, recipient):
    try:
        s = smtplib.SMTP()
        s.connect(mail_host)
        s.ehlo()
        s.starttls()
        s.login(mail_user,mail_pass)
        s.sendmail(sender, recipient, msg.as_string())
        s.close()
        return True
    except Exception as e:
        print(str(e))
    return False

def write_mail(sender, recipient, sub, content):
    name = Header(sender, 'utf-8').encode()
    msg = MIMEText(content, _subtype = 'html', _charset='utf-8')
    msg['Subject'] = Header(sub, 'utf-8')
    msg['From'] = formataddr((name, sender_addr))
    msg['To'] = recipient
    return msg

def main():

    if  len(sys.argv) < 2:
        print('错误，没有指定参数')
        print('用法：python Send_wages_Email.py xxx.xls')
        sys.exit()

    bk = xlrd.open_workbook(sys.argv[1])
    #bk.sheets()返回一个列表
    sh = bk.sheets()[0]  #读取第一张sheet
    #下面是按行读取excel表格内容
    for row in range(2, sh.nrows):
        name = sh.row(row)[0].value # 姓名
        deparment = sh.row(row)[1].value # 部门
        recipient_addr = sh.row(row)[2].value # Email
        Attendance_hours = sh.row(row)[3].value # 出勤工时
        post_salary = sh.row(row)[4].value # 岗位津贴
        base_salary = sh.row(row)[5].value # 基本工资
        pay_for_performance = sh.row(row)[7].value # 绩效工资
        overtime_subsidies = sh.row(row)[8].value # 加班补贴
        post_allowance = sh.row(row)[8].value # 津贴
        sick_hours = sh.row(row)[9].value # 请假工时
        sick_salary = sh.row(row)[10].value # 请假工时
        donastion = sh.row(row)[11].value # 乐捐
        total = sh.row(row)[12].value # 应发合计
        company_pension_commitment = sh.row(row)[13].value # 公司承担社保
        company_housing_fund_commitment = sh.row(row)[14].value # 公司承担住房公积金
        water_rate = sh.row(row)[16].value # 水电费
        Gross_salary= sh.row(row)[16].value # 实发工资
        personal_income_tax = sh.row(row)[17].value # 个税
        real_salary = sh.row(row)[18].value # 实发工资
        today = datetime.now().strftime('%Y/%m/%d')


        content = write_mail_content(title, name, deparment,Attendance_hours,
                post_salary,base_salary,pay_for_performance,overtime_subsidies,post_allowance,
                sick_hours,sick_salary,donastion,total,company_pension_commitment,company_housing_fund_commitment,
                water_rate,Gross_salary,personal_income_tax,real_salary,today,sender_name)

        msg = write_mail(sender_name, recipient_addr, subject, content)
        if send_mail(msg, sender_addr, recipient_addr):
            print(' 姓名：' + name + ' 发送成功.')
        else:
            print(' 姓名：' + name + ' 发送失败.')

    print('Send all finished! please check out the failed records.')

if __name__ == '__main__':
    main()