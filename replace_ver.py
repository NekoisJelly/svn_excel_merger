# coding=utf-8

import re
import os
import sys
import subprocess


def replace_file(filename, ver):
    f_bak = filename + ".bak"
    rf = open(filename, "r", encoding='utf-8')
    wf = open(f_bak, "w")
    rfl = rf.readlines()
    rf.close()
    for line in rfl:
        if "<string>Excel Merge" in line:
            wf.write("   <string>Excel Merge v" + ver + "</string>\n")
        elif "#define AppVersion" in line:
            wf.write("#define AppVersion \"" + ver + "\"\n")
        else:
            wf.write(line)
    wf.close()
    os.remove(filename)
    os.rename(f_bak, filename)


def main():
    # if len(sys.argv) == 2:
    #     ver = sys.argv[1]
    # else:
    #     raise "usage: python replace_ver.py 2017.02"
    ver = '2017.01'
    log_result = subprocess.Popen("svn log -l1 -q --username ak47159754@vip.qq.com svn://gitee.com/Shelc/ExcelMerger",
                                  stderr=subprocess.STDOUT, stdout=subprocess.PIPE, shell=True).communicate()[0]
    log_result = log_result.decode('gbk')
    print(log_result)
    changelist = re.compile('r[0-9]+').findall(log_result)
    ver = ver + "(" + str(changelist[0]) + ")"

    # modify main.ui, setup.iss
    replace_file("main.ui", ver)

main()