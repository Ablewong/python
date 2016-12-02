__author__ = 'able'
# ! /usr/bin/env python
# -*- coding: uft-8 -*-

import os
import os.path


def walk_dir(dir, fileinfo, topdown=True):
    for root, dirs, files in os.walk(dir, topdown):
        for name in files:
            pathN = 'c:\computer\\%s'%(name)
            catinfo = open(pathN, 'r')
            hardwarename = ['Operating System', 'System Manufacturer', 'Processor', 'OS Memory', 'Chip type',
                            'Monitor Model', 'Monitor Id', '    Model:']
            fileinfo.write(os.path.join(name) + '\n')
            for readline in catinfo:
                for hard in hardwarename:
                    if hard in readline:
#                        if readline.strip().startswith("Model: WDC"):
#                            print(readline)
#                            continue
                        fileinfo.write(readline)
                        print(readline)
            fileinfo.write(os.path.join(root, name) + '\n' + '\n')


#        for name in dirs:
#            print (os.path.join(name))
#            fileinfo.write(' ' + os.path.join(root,name) + '\n')
def main():
    dir = input('please input the path:')
    fileinfo = open("list.txt", 'w')
    walk_dir(dir, fileinfo)
    fileinfo.close()


main()


