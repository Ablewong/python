__author__ = 'able'
# !/usr/bin/env python
#-*- encoding: utf-8 -*-

import platform
import wmi

def get_os_info():
    print (platform.platform())
    return True

c = wmi.WMI()
def sys_version():
    # 获取操作系统版本
    for sys in c.Win32_OperatinSystem():
        print ("Version:%s" % sys.Caption.encode("UTF8"),"Vernum:%s" % sys.BuildNumber)
        print (sys.OSArchitecture.encode("UTF8")) #系统是32位还是64位
        # print sys.NumberOfProcesses  当前系统运行的进程总数

def cpu_mem():
    #cpu类型和内存
    for processor in c.Win32_Processor():
        print ("Prosess Name % s" % processor.Name.strip())
    for Memory in c.Win32_PhysicalMemory():
        print ("Memory Capacity: %.FMB" % (int(Memory.Capacity)/1048576))
#    for display in c.Win32_DisplayConfiguration():
#        print ("Display Name: % s" % (display.Name.strip()))
def disk():
    # 获取硬盘使用情况
    all = 0
    for disk in c.Win32_logicalDisk (DriveType=3):
        print (disk.Caption, "%0.2f%% free" % (100.0 * long (disk.FreeSpace) / long (disk.Size)))
       # print "The disk size is %.2f GB(s)" % (float(disk.Size) / 1024 / 1024 / 1024)
        c1 = (float(disk.Size)/ 1024 / 1024 / 1024) + all
        all = c1

    print (int(all), 'GB')

def check_display(self):#显卡信息
    colItems = self.objSWbemServices.ExecQuery("Select * from Win32_DisplayConfiguration")
    display_list = []
    for objItem in colItems:
        Display_list.append(objItem.DeviceName)
    display_info = ';'.join(Display_list)
    return display_info

def main():
    hardware = open('c:\systeminfo.txt','w')
    hardware.write(get_os_info())
    hardware.write(cpu_mem())
    hardware.write('硬盘使用情况')
    for haredisk in disk():
        hardware.write(haredisk)
    haredisk.close()
#    print check_display()
    #cpu_men()
    #disk()
main()
#    print(platform.system())
#    print(platform.release())
#    print(platform.platform())
#    print(platform.machine())
