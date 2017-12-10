# coding=utf-8

import re
import os
import sys
import datetime
import shutil
import multiprocessing

import win32com.client as win32
import pythoncom
import pywintypes
import xlrd
from PyQt4 import QtCore
from PyQt4 import QtGui
from PyQt4.QtCore import QObject
from PyQt4.QtCore import SIGNAL
from PyQt4.QtGui import QIcon
from ui import Ui_Dialog
from svnoperator import SvnOperator

EVENT_ERROR = "$EVENT-ERROR$"
EVENT_FINISHED = "$EVENT-FINISHED$"
trunk_url = 'svn://gitee.com/Shelc/ExcelMerger'
trunk_sub = '/trunk_xlsdir'
branch_dir = "C:\\Users\\Admin\\Desktop\\test"
branch_sub = '/20160717_xlsdir'
all_changes = {}
ignore_changes = {}
cache_data = {}
g_mp = None
error_msg = []
Application_Excel_Version = None
svn_optr = None
g_finished_ok = False
g_branch_first_ver = 0


def to_int(n):
    try:
        if n is None:
            return None
        if isinstance(n, int):
            return n
        if isinstance(n, float):
            return int(n)
        if isinstance(n, str):
            n = n.replace('\'', '')
            return int(float(n))

        return int(float(str(n)))
    except ValueError as e:
        print("to_int Error:" + str(e))
    return None


# class QMultiThread(QtCore.QThread):
#     update_ui = QtCore.pyqtSignal(object)
#
#     def __init__(self, keys):
#         QtCore.QThread.__init__(self)
#         self.keys = keys
#
#     def run(self):
#         multi_thread(self.keys)


class QMultiProcess(QtCore.QProcess):
    update_ui = QtCore.pyqtSignal(object)

    def __init__(self, keys):
        QtCore.QProcess.__init__(self)
        self.keys = keys

    def run(self):
        multi_process(self.keys)


def log(s):
    assert isinstance(s, str)
    if g_mp:
        g_mp.update_ui.emit(s)


def log_error(s):
    assert isinstance(s, str)
    if g_mp:
        g_mp.update_ui.emit(EVENT_ERROR + s)


# make sure 's' is utf-8 string
def log_ui(s):
    global error_msg

    # button click notify
    if s == EVENT_FINISHED:
        ui.pushButton_go.setEnabled(True)

        # show error massage if error_msg is not empty
        if len(error_msg) > 0:
            msg = ""
            for l in error_msg:
                msg += l + "\n"
            # QtGui.QMessageBox.information(Dialog, u'Excel Merge Tool', QString.fromUtf8(msg))
            wfile = open("issue.log", "w")
            wfile.write(msg)
            wfile.close()
            os.startfile("issue.log")
            error_msg = []
        return

    # for error msg
    if s.startswith(EVENT_ERROR):
        s = s[len(EVENT_ERROR):]
        error_msg.append(s)

    print(s)
    ui.textEdit_status.append("[" + str(datetime.datetime.now()) + "]" + s)
    scroll = ui.textEdit_status.verticalScrollBar()
    if scroll:
        scroll.setValue(scroll.maximum())


def get_abspath():
    try:
        root_dir = os.path.dirname(os.path.abspath(__file__))
    except NameError:  # We are the main py2exe script, not a module
        import sys
        root_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
    return root_dir


def get_suitable_excel_version():
    """获取excel版本"""
    def anonymous_func(ver1, ver2):
        result = False
        try:
            # http://stackoverflow.com/questions/5964805/implement-com-interface-type-library-in-python
            # 确认win32com缓存模块正常创建了
            win32.gencache.EnsureModule('{00020813-0000-0000-C000-000000000046}', 0, 1, int(ver1))
            excel = win32.Dispatch('Excel.Application.' + str(ver2))
            # 如果当前有打开的excel工作薄，直接返回
            if len(excel.Workbooks) > 0:
                return None
            result = True
        except pywintypes.com_error as e:
            None
        finally:
            None
        return result

    if anonymous_func(6, 12):
        return 6, 12
    if anonymous_func(7, 14):
        return 7, 14
    if anonymous_func(8, 15):
        return 8, 15
    if anonymous_func(9, 16):
        return 9, 16
    return None


def get_branch_full_file_name(filename):
    if not filename.startswith(trunk_sub):
        return None

    # make sure dir is exist
    sub_real = (filename[len(trunk_sub):len(filename)]).replace("/", "\\")
    folders = sub_real.split("\\")
    if len(folders) > 2:
        dirs = folders[1:len(folders)-1]
        cur_dir = branch_dir + branch_sub
        for d in dirs:
            cur_dir += "\\" + d
            if not os.path.exists(cur_dir):
                os.makedirs(cur_dir.decode('utf8'))
    return (branch_dir + branch_sub + filename[len(trunk_sub):len(filename)]).replace("/", "\\")


def get_temp_file_name(ext):
    # full path prefix
    prefix = get_abspath() + "\\"

    if not os.path.exists((prefix + "temp/t" + ext)):
        return prefix + "temp/t" + ext

    num = 1
    while os.path.exists((prefix + "temp/t" + str(num) + ext).decode('utf8')):
        num += 1
    return prefix + "temp/t" + str(num) + ext


def download_trunk_url_file(ver, filename):
    _, ext = os.path.splitext(filename)
    temp_file = get_temp_file_name(ext)
    svn_optr.download_url_file(ver, filename, temp_file)
    return temp_file


def read_excel_xlrd(f):
    excel = xlrd.open_workbook(f)
    table = excel.sheets()[0]
    row = table.nrows
    col = table.ncols

    # if this is an empty excel
    if row == 0 or col == 0:
        return 0, {}

    while table.cell(0, col-1).value == u"" or table.cell(0, col-1).value is None:
        col -= 1
        if col == 0:
            return None

    result = {}
    for r in range(1, row):
        local_row = []
        for c in range(0, col):
            if table.cell(r, c).ctype == xlrd.XL_CELL_NUMBER and int(table.cell(r, c).value) == table.cell(r, c).value:
                local_row.append(table.cell(r, c).value)
            elif xlrd.XL_CELL_DATE == table.cell(r, c).ctype:
                showval = xlrd.xldate_as_tuple(table.cell(r, c).value, excel.datemode)
                # 若年月日都为零，以格式(12:00:00)处理, 否则以日期格式(2016-12-12)处理
                if showval[0] == 0 and showval[1] == 0 and showval[2] == 0:
                    cell_value = "\'%d:%02d:%02d" % (showval[3], showval[4], showval[5])
                    cell_value = cell_value.decode('utf-8')
                    local_row.append(cell_value)
                else:
                    cell_value = "%4d-%02d-%02d" % (showval[0], showval[1], showval[2])
                    cell_value = cell_value.decode('utf-8')
                    local_row.append(cell_value)
            elif xlrd.XL_CELL_TEXT == table.cell(r, c).ctype:
                if table.cell(r, c).value == u"" or table.cell(r, c).value == "" or table.cell(r, c).value is None:
                    local_row.append(None)
                else:
                    local_row.append("\'" + table.cell(r, c).value)
            elif xlrd.XL_CELL_EMPTY == table.cell(r, c).ctype or xlrd.XL_CELL_BLANK == table.cell(r, c).ctype:
                local_row.append(None)
            else:
                if table.cell(r, c).value == u"" or table.cell(r, c).value == "" or table.cell(r, c).value is None:
                    local_row.append(None)
                else:
                    local_row.append(str(table.cell(r, c).value))

        rowid = to_int(local_row[0])
        if rowid is not None:
            local_row[0] = rowid
            result[rowid] = local_row

    return col, result


def get_trunk_url_file_data(ver, filename, before=False):
    global cache_data
    global g_branch_first_ver

    if before:
        ver = svn_optr.get_file_ver_before(g_branch_first_ver, ver, filename)
    key = str(ver) + filename
    if key in cache_data:
        return cache_data[key]

    tf = download_trunk_url_file(ver, filename)
    cache_data[key] = read_excel_xlrd(tf)

    try:
        os.remove(tf)
    except WindowsError as e:
        log(str(e))
    return cache_data[key]


def all_changes_of_one_file(filename, new_file, old_file):
    global all_changes
    global ignore_changes

    # NOT SUPPORT IF COL WERE CHANGED.
    if new_file[0] != old_file[0]:
        log_error("ERROR COL CHANGED:" + filename)
        ignore_changes[filename] = 1
        # 删除这个文件对应的修改
        if filename in all_changes:
            all_changes.pop(filename)
        return

    new_file_data = new_file[1]
    old_file_data = old_file[1]
    # both in new_file and in old_file
    both_in = dict([(i, new_file_data[i]) for i in filter(lambda k:k in new_file_data, old_file_data.keys())])

    # add rows
    add_result = list([new_file_data[i] for i in filter(lambda k: k not in old_file_data, new_file_data.keys())])

    # delete rows
    delete_result = list([i for i in filter(lambda k: k not in new_file_data, old_file_data.keys())])

    # delete rows
    delete_result_old = list([old_file_data[i] for i in filter(lambda k: k not in new_file_data, old_file_data.keys())])

    modify_result = []
    modify_result_old = []
    for k in both_in:
        if both_in[k] != old_file_data[k]:
            modify_result.append(both_in[k])
            modify_result_old.append(old_file_data[k])

    merge_one_file_diff(filename, add_result, modify_result, modify_result_old, delete_result, delete_result_old)


def merge_one_file_diff(filename, add_result, modify_result, modify_result_old, delete_result, delete_result_old):
    global all_changes
    global ignore_changes

    # if nothing changed
    if len(add_result) == 0 and len(modify_result) == 0 and len(delete_result) == 0:
        return

    # if nothing changed before for this file.
    if filename not in all_changes:
        all_changes[filename] = (add_result, modify_result, modify_result_old, delete_result, delete_result_old)
        return

    old_add_result = all_changes[filename][0]
    old_modify_result = all_changes[filename][1]
    old_modify_result_old = all_changes[filename][2]
    old_delete_result = all_changes[filename][3]
    old_delete_result_old = all_changes[filename][4]

    # check old data col and new data col
    if len(add_result) > 0 and len(old_add_result) > 0 and len(add_result[0]) != len(old_add_result[0]):
        log_error("ERROR COL CHANGED:" + filename)
        ignore_changes[filename] = 1
        # 删除这个文件对应的修改
        if filename in all_changes:
            all_changes.pop(filename)
        return

    # check old data col and new data col
    if len(modify_result) > 0 and len(old_modify_result) > 0 and len(modify_result[0]) != len(old_modify_result[0]):
        log_error("ERROR COL CHANGED:" + filename)
        ignore_changes[filename] = 1
        # 删除这个文件对应的修改
        if filename in all_changes:
            all_changes.pop(filename)
        return

    # add 逻辑处理(共四个步骤)
    # ＊＊＊＊＊＊＊＊ ［Start］Add 逻辑 ＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊
    # 1.若add_result数据已经存在于old_add_result中，则报冲突，同时丢弃此行的处理
    conflict_id_list = []
    for l in add_result:
        for x in old_add_result:
            if l[0] == x[0]:
                conflict_id_list.append(l[0])
    for rowid in conflict_id_list:
        for l in add_result:
            if l[0] == rowid:
                log_error("[add1 conflict]:" + filename + " [id]:" + str(rowid))
                add_result.remove(l)
                break
    # 2.若add_result数据已经存在于old_modify_result中，则报冲突，同时丢弃此行的处理
    conflict_id_list = []
    for l in add_result:
        for x in old_modify_result:
            if l[0] == x[0]:
                conflict_id_list.append(l[0])
    for rowid in conflict_id_list:
        for l in add_result:
            if l[0] == rowid:
                log_error("[add2 conflict]:" + filename + " [id]:" + str(rowid))
                add_result.remove(l)
                break
    # 3.若add_result数据已经存在于old_delete_result中，则依据old_delete_result_old中的数据转换为modify或者直接丢弃
    delete_directly = []
    convert_modify = []
    for row in old_delete_result_old:
        for l in add_result:
            if l[0] == row[0]:
                # 数据没有变化，相当于恢复以前的数据了
                # 直接丢弃（old_delete_result, old_delete_result_old，add_result 对应的数据）
                if l == row:
                    delete_directly.append(l[0])
                # 数据变化了，需要转换为modify类型
                else:
                    convert_modify.append(l[0])
    # 直接丢弃
    for l in delete_directly:
        for x in add_result:
            if l == x[0]:
                add_result.remove(x)
                break
        for x in old_delete_result:
            if l == x:
                old_delete_result.remove(x)
                break
        for x in old_delete_result_old:
            if l == x[0]:
                old_delete_result_old.remove(x)
                break
    # 转换为modify类型
    for l in convert_modify:
        for x in add_result:
            if l == x[0]:
                old_modify_result.append(x) # 可以直接添加，第2步已经处理过可能的冲突情况
                add_result.remove(x)
                break
        for x in old_delete_result_old:
            if l == x[0]:
                old_modify_result_old.append(x) # 可以直接添加，第2步已经处理过可能的冲突情况
                old_delete_result_old.remove(x)
                break
        for x in old_delete_result:
            if l == x:
                old_delete_result.remove(x)
                break
    # 4.剩下的add_result直接插入old_add_result中
    for l in add_result:
        old_add_result.append(l)
    add_result = []
    # ＊＊＊＊＊＊＊＊ ［End］Add 逻辑 ＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊

    # modify 逻辑处理（共四个步骤）
    # ＊＊＊＊＊＊＊＊ ［Start］Modify 逻辑 ＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊
    # 1.若modify_result数据已经存在于old_add_result中，则删除本条数据(包括modify_result_old)，并更新old_add_result数据即可
    conflict_row_list = []
    for l in modify_result:
        for x in old_add_result:
            if l[0] == x[0]:
                conflict_row_list.append(l)
    for row in conflict_row_list:
        conflict_data1 = None
        conflict_data2 = None
        for l in modify_result:
            if l[0] == row[0]:
                modify_result.remove(l)
                break
        for l in modify_result_old:
            if l[0] == row[0]:
                conflict_data1 = l
                modify_result_old.remove(l)
                break
        for l in old_add_result:
            if l[0] == row[0]:
                conflict_data2 = l
                break
        # 冲突了呢
        if conflict_data1 != conflict_data2:
            log_error("[modify1 conflict]:" + filename + " [id]:" + str(row[0]))
        else:
            for l in old_add_result:
                if l[0] == row[0]:
                    for i in range(1, len(l)):
                        l[i] = row[i]
                    break
    # 2.若modify_result数据已经存在于old_modify_result中
    # 则删除本条数据(包括modify_result_old)，并更新old_modify_result数据即可
    conflict_row_list = []
    for l in modify_result:
        for x in old_modify_result:
            if l[0] == x[0]:
                conflict_row_list.append(l)
    for row in conflict_row_list:
        conflict_data1 = None
        conflict_data2 = None
        for l in modify_result:
            if l[0] == row[0]:
                modify_result.remove(l)
                break
        for l in modify_result_old:
            if l[0] == row[0]:
                conflict_data1 = l
                modify_result_old.remove(l)
                break
        for l in old_modify_result:
            if l[0] == row[0]:
                conflict_data2 = l
                break
        # 冲突了呢
        if conflict_data1 != conflict_data2:
            log_error("[modify2 conflict]:" + filename + " [id]:" + str(row[0]))
        else:
            for l in old_modify_result:
                if l[0] == row[0]:
                    for i in range(1, len(l)):
                        l[i] = row[i]
                    break
    # 3.若modify_result数据已经存在于old_delete_result中，则报冲突，同时丢弃此行的处理
    conflict_id_list = []
    for l in modify_result:
        for x in old_delete_result:
            if l[0] == x:
                conflict_id_list.append(l[0])
    for rowid in conflict_id_list:
        for l in modify_result:
            if l[0] == rowid:
                log_error("[modify3 conflict]:" + filename + " [id]:" + str(rowid))
                modify_result.remove(l)
                break
    # 4.剩下的数据直接插入old_modify_result 和 old_modify_result_old 中
    for l in modify_result:
        old_modify_result.append(l)
    modify_result = []
    for l in modify_result_old:
        old_modify_result_old.append(l)
    modify_result_old = []
    # ＊＊＊＊＊＊＊＊ ［End］Modify 逻辑 ＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊

    # delete 逻辑处理(共四个步骤)
    # ＊＊＊＊＊＊＊＊ ［Start］Delete 逻辑 ＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊
    # 1.若delete_result数据已经存在于old_add_result中，
    conflict_id_list = []
    for l in delete_result:
        for x in old_add_result:
            if x[0] == l:
                conflict_id_list.append(l)
                break
    for rowid in conflict_id_list:
        for x in delete_result:
            if rowid == x:
                delete_result.remove(x)
                break
        conflict_data1 = None
        conflict_data2 = None
        for x in delete_result_old:
            if rowid == x[0]:
                conflict_data1 = x
                delete_result_old.remove(x)
                break
        for x in old_add_result:
            if rowid == x[0]:
                conflict_data2 = x
                break
        # 冲突了呢
        if conflict_data1 != conflict_data2:
            log_error("[delete1 conflict]:" + filename + " [id]:" + str(rowid))
        # 从old_add_result中移除
        else:
            for x in old_add_result:
                if x[0] == rowid:
                    old_add_result.remove(x)
                    break
    # 2.若delete_result数据已经存在于old_modify_result中，则转换为delete类型。
    # old_delete_result_old 数据也要转换. 删除 old_modify_result, old_modify_result_old)
    conflict_id_list = []
    for l in delete_result:
        for x in old_modify_result:
            if l == x[0]:
                conflict_id_list.append(l)
    for rowid in conflict_id_list:
        for x in delete_result:
            if x == rowid:
                delete_result.remove(x)
                break
        conflict_data1 = None
        conflict_data2 = None
        for x in delete_result_old:
            if x[0] == rowid:
                conflict_data1 = x
                delete_result_old.remove(x)
                break
        for x in old_modify_result:
            if x[0] == rowid:
                conflict_data2 = x
                break
        if conflict_data1 != conflict_data2:
            log_error("[delete2 conflict]:" + filename + " [id]:" + str(rowid))
        else:
            # a. delete from old_modify_result
            for x in old_modify_result:
                if x[0] == rowid:
                    old_modify_result.remove(x)
                    break
            # b. delete from old_modify_result_old, update old_delete_result_old
            for x in old_modify_result_old:
                if x[0] == rowid:
                    old_delete_result_old.append(x)
                    old_modify_result_old.remove(x)
                    break
            # c. add old_delete_result
            old_delete_result.append(rowid)
    # 3.若delete_result数据已经存在于old_delete_result中，则报冲突，同时丢弃此行的处理
    conflict_id_list = []
    for l in delete_result:
        for x in old_delete_result:
            if l == x:
                conflict_id_list.append(l)
    for rowid in conflict_id_list:
        log_error("[delete3 conflict]:" + filename + " [id]:" + str(l))
        for x in delete_result:
            if x == rowid:
                delete_result.remove(x)
                break
        for x in delete_result_old:
            if x[0] == rowid:
                delete_result_old.remove(x)
                break
    # 4.直接插入old_delete_result 和 old_delete_result_old 中
    for l in delete_result:
        old_delete_result.append(l)
    delete_result = []
    for l in delete_result_old:
        old_delete_result_old.append(l)
    delete_result_old = []
    # ＊＊＊＊＊＊＊＊ ［End］Delete 逻辑 ＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊

    # 判断一下，若没有任何修改了，则不再记录这个文件的变化
    if len(old_add_result) == 0 and len(old_modify_result) == 0 and len(old_delete_result) == 0:
        if filename in all_changes:
            all_changes.pop(filename)


def issue_one_file_change(ver, filename, change_type):
    global ignore_changes
    global all_changes

    full_branch_file = filename
    try:
        file_flag = "M"  # 'M' for type 0
        if change_type == 1:
            file_flag = "A"
        if change_type == 2:
            file_flag = "D"

        log("----正在获取单个文件差异")
        log("----r:" + str(ver) + " " + filename + " [" + file_flag + ']')

        full_branch_file = get_branch_full_file_name(filename)
        if full_branch_file is None:
            log_error("---- IGNORE NOT AFFECTED PATH:" + filename)
            return

        # for delete file
        if change_type == 2:
            if os.path.exists(full_branch_file):
                delete_log = svn_optr.delete_local_file(full_branch_file)
                log(delete_log)
            else:
                log_error("----FileNotExist:" + full_branch_file)

            # 删除这个文件对应的修改
            if filename in all_changes:
                all_changes.pop(filename)
            return

        # for add file, just copy to branch
        if change_type == 1:
            if os.path.exists(full_branch_file.decode('utf8')):
                log_error("----FileAlreadyExist:" + full_branch_file)
                ignore_changes[filename] = 1
                return

            srcfile = download_trunk_url_file(ver, filename)
            shutil.move(srcfile.decode('utf8'), full_branch_file.decode('utf8'))

            add_log = svn_optr.add_local_file(full_branch_file)
            log(add_log)
            return

        # for type modify, conflict
        assert change_type == 0
        if not os.path.exists(full_branch_file):
            log_error("[M][文件不存在]:" + full_branch_file)
            return

        # 添加删除文件，不受ignore_changes影响
        if filename in ignore_changes:
            log("----文件冲突(或列变化)，跳过此文件！")
            return

        new_file_data = get_trunk_url_file_data(ver, filename)
        old_file_data = get_trunk_url_file_data(ver, filename, True)
        if new_file_data is None or old_file_data is None:
            ignore_changes[filename] = 1
            log_error("[Invalid File]:" + filename)
            return

        all_changes_of_one_file(filename, new_file_data, old_file_data)
    except (WindowsError, IOError) as e:
        log_error("[" + file_flag + "][获取差异失败]:" + full_branch_file)
        log_error(str(e))


def issue_all_commits(commit_list):
    log("\n*************** 正在获取所有差异 **************************")
    for oc in commit_list:
        ver = oc[0]
        log("\n正在获取单次提交差异r:" + str(ver))
        log("****[LOG]:" + oc[1])
        for f in oc[2]:
            filename = f[0]
            modify_type = f[1]
            issue_one_file_change(ver, filename, modify_type)


def pre_merge_branch_file(filename, records_all):
    records_add = records_all[0]
    records_modify = records_all[1]
    records_modify_old = records_all[2]
    records_delete = records_all[3]

    f = get_branch_full_file_name(filename)
    branch_data = read_excel_xlrd(f)[1]

    # for add conflict
    for b in branch_data:
        for o in records_add:
            if o[0] == branch_data[b][0]:
                if o != branch_data[b]:
                    log_error("添加行冲突，分支文件中已经存在且内容不一致")
                    log_error("[FILE]:" + f + " [id]:" + str(o[0]))
                records_add.remove(o)
                break

    # for modify conflict
    modify_need_remove = []
    for o in records_modify:
        can_find_row = False
        for b in branch_data:
            if o[0] == branch_data[b][0]:
                can_find_row = True
                break
        if not can_find_row:
            log_error("修改行冲突，分支文件中待修改行不存在")
            log_error("[FILE]:" + f + " [id]:" + str(o[0]))
            modify_need_remove.append(o)

    for o in modify_need_remove:
        records_modify.remove(o)

    def get_records_modify_item_by_id(rowid):
        for x in records_modify:
            if x[0] == rowid:
                return x
        return None

    # for modify_old conflict
    for b in branch_data:
        for o in records_modify_old:
            if o[0] == branch_data[b][0] and o != branch_data[b]:
                if branch_data[b] != get_records_modify_item_by_id(o[0]):
                    log_error("修改行冲突，修改前后内容都不一致")
                    log_error("[FILE]:" + f + " [id]:" + str(o[0]))

                # remove modify item
                for x in records_modify:
                    if x[0] == o[0]:
                        records_modify.remove(x)
                        break
                break

    # for delete conflict
    delete_need_remove = []
    for o in records_delete:
        can_find_row = False
        for b in branch_data:
            if o == branch_data[b][0]:
                can_find_row = True
                break
        if not can_find_row:
            log_error("删除行冲突，分支文件中待删除行不存在")
            log_error("[FILE]:" + f + " [id]:" + str(o))
            delete_need_remove.append(o)

    for o in delete_need_remove:
        records_delete.remove(o)

    # modify merge to add
    for l in records_modify:
        records_add.append(l)


def merge_branch_file(filename, records_all):
    records_add = records_all[0]
    # records_modify = records_all[1]
    # records_modify_old = records_all[2]
    records_delete = records_all[3]

    f = get_branch_full_file_name(filename)
    if len(records_add) == 0 and len(records_delete) == 0:
        log("[Nothing Change] " + f)
        return

    log("[M] " + f)
    excel = None
    wb = None
    try:
        # http://stackoverflow.com/questions/5964805/implement-com-interface-type-library-in-python
        win32.gencache.EnsureModule('{00020813-0000-0000-C000-000000000046}', 0, 1, int(Application_Excel_Version[0]))
        excel = win32.Dispatch('Excel.Application.' + str(Application_Excel_Version[1]))
        excel.DisplayAlerts = False
        excel.Visible = 0

        wb = excel.Workbooks.Open(f)
        ws = wb.Worksheets(1)

        used = ws.UsedRange
        row_count = used.Row + used.Rows.Count - 1
        col_count = used.Column + used.Columns.Count - 1

        # for debug
        # assert row_count == ws.Range('A65536').End(win32.constants.xlUp).Row
        # for debug

        # xlrd bug??? max col is less than win32com max col
        if len(records_add) > 0:
            while True:
                cell_value = ws.Cells(1, col_count).Value
                if col_count > len(records_add[0]) and cell_value is None:
                    col_count -= 1
                else:
                    break

        if len(records_add) > 0 and (col_count > len(records_add[0])):
            log_error("[COL WARNING, Please Check This File By Yourself]:" + f)
            col_count = len(records_add[0])

        # read all row id
        xls_row_ids = []
        for r in range(2, row_count + 1):
            xls_row_ids.append(ws.Cells(r, 1).Value)

        # for record delete
        row_index = 1
        row_need_delete = []
        for r in xls_row_ids:
            row_index += 1
            if r in records_delete:
                ws.Rows(row_index).Delete()
                row_index -= 1
                row_need_delete.append(r)

        if row_count != (used.Row + used.Rows.Count - 1) + len(row_need_delete):
            # 某些文件row delete后, (used.Row + used.Rows.Count - 1)的值并不发生变化, 特殊处理一下。
            row_count -= len(row_need_delete)
        else:
            assert used.Row + used.Rows.Count - 1 == row_count - len(row_need_delete)
            row_count = used.Row + used.Rows.Count - 1

        # delete old
        for v in row_need_delete:
            xls_row_ids.remove(v)

        # because of title, so add 1
        assert row_count == len(xls_row_ids) + 1

        def get_excel_col_name(col):
            col = int(col)
            assert col >= 1
            if col <= 26:
                return chr(ord('A') - 1 + col)
            if col % 26 == 0:
                return get_excel_col_name(col / 26 - 1) + 'Z'
            else:
                return get_excel_col_name(col / 26) + get_excel_col_name(col % 26)

        add_row_count = 0
        for record in records_add:
            is_exist_record = False
            for r in range(2, row_count + 1):
                if to_int(xls_row_ids[r-2]) == record[0]:
                    is_exist_record = True
                    ws.Range("A"+str(r),  get_excel_col_name(len(record))+str(r)).Value = record
                    # for c in range(1, col_count + 1):
                    #    ws.Cells(r, c).Value = record[c - 1]
                    break
            # insert new a row
            if not is_exist_record:
                add_row_count += 1
                local_row = row_count + add_row_count
                ws.Range("A" + str(local_row), get_excel_col_name(len(record)) + str(local_row)).Value = record
                # for c in range(1, col_count + 1):
                #    ws.Cells(row_count + add_row_count, c).Value = record[c - 1]
        wb.SaveAs(f)
    except pywintypes.com_error as e:
        log_error("[修改Excel文件失败]:" + f)
    finally:
        if wb:
            wb.Close()
        #if excel:
        #    excel.Application.Quit()


def clean_file_empty_rows(filename):
    log("[C] " + filename)
    excel = None
    wb = None
    try:
        # http://stackoverflow.com/questions/5964805/implement-com-interface-type-library-in-python
        win32.gencache.EnsureModule('{00020813-0000-0000-C000-000000000046}', 0, 1, int(Application_Excel_Version[0]))
        excel = win32.Dispatch('Excel.Application.' + str(Application_Excel_Version[1]))
        excel.DisplayAlerts = False
        excel.Visible = 0

        wb = excel.Workbooks.Open(filename)
        ws = wb.Worksheets(1)

        used = ws.UsedRange
        row_count = used.Row + used.Rows.Count - 1

        # for debug
        # assert row_count == ws.Range('A65536').End(win32.constants.xlUp).Row
        # for debug

        empty_rows = []
        for r in range(1, row_count + 1):
            if (ws.Cells(r, 1).Value is None) or (ws.Cells(r, 1).Value == ""):
                empty_rows.append(r)
        if len(empty_rows) > 200:
            log_error("[Too Many Empty Rows]:" + str(len(empty_rows)) + " " + filename)
            return

        empty_rows = sorted(empty_rows, reverse=True)
        delete_index = 1
        row_count = len(empty_rows)
        for d in empty_rows:
            ws.Rows(d).Delete()
            log("Deleting " + str(delete_index) + " Count:" + str(row_count))
            delete_index += 1

        # if some row deleted
        if delete_index > 1:
            wb.SaveAs(filename)
    except pywintypes.com_error as e:
        log_error("[修改Excel文件失败]:" + filename)
    finally:
        if wb:
            wb.Close()
        #if excel:
        #    excel.Application.Quit()


def main(keys):
    global all_changes
    global ignore_changes
    global g_finished_ok
    global g_branch_first_ver

    # 清理当前目录表格文件空行
    if len(keys) == 0:
        results = []
        results2 = []
        local_dir = get_abspath()
        for root, dis, files in os.walk(local_dir):
            if root == local_dir:
                for f in files:
                    results.append(f)
            else:
                for f in files:
                    results2.append(root[len(local_dir) + 1:] + "\\" + f)

        results = results + results2
        for filename in results:
            if filename.endswith(".xls") or filename.endswith(".xlsm") or filename.endswith(".xlsx"):
                clean_file_empty_rows((local_dir + "\\" + filename))

        log("\n----------------------------------------------------------------------------------")
        log("二师兄，开饭啦！！！")
        log(EVENT_FINISHED)
        g_finished_ok = True
        return

    # clean temp
    try:
        tmp_dir = (get_abspath() + "/temp/")
        if os.path.isdir(tmp_dir):
            for f in os.listdir(tmp_dir):
                os.remove(tmp_dir + f)
        else:
            os.mkdir(tmp_dir)
    except WindowsError as e:
        log("main(): os.remove failed." + str(e))

    # clean data
    all_changes = {}
    ignore_changes = {}

    log("正在获取分支切出时间点（预计需要10秒）：")
    dt = datetime.datetime.now()
    if g_branch_first_ver is None:
        g_branch_first_ver = svn_optr.get_repository_oldest_ver(trunk_url + branch_sub)
    log("branch first ver=" + str(g_branch_first_ver))
    log("---- spent:" + str(datetime.datetime.now() - dt)[:7])
    commits = []
    for v in keys:
        log("正在根据关键字请求提交信息（预计需要10秒）：" + v)
        dt = datetime.datetime.now()
        commits_by_key = svn_optr.get_commits_by_key(g_branch_first_ver, v)
        log("---- spent:" + str(datetime.datetime.now() - dt)[:7])
        for oc in commits_by_key:
            ver = oc[0]
            log_text = oc[1]
            file_list = oc[2]
            log("ver:" + str(ver) + " " + log_text)
            for f in file_list:
                file_name = f[0]
                file_type = f[1]
                file_type_str = "[M]"
                if file_type == 1:
                    file_type_str = "[A]"
                if file_type == 2:
                    file_type_str = "[D]"
                log("    " + file_name + file_type_str)
        commits += commits_by_key

    commits.sort(key=lambda x: x[0])
    issue_all_commits(commits)

    log("\n\n-----------------------------------------")
    log("输出所有差异列表（不包含Add Delete操作）:")
    for k in all_changes:
        log("\n" + k)
        for v in all_changes[k][0]:
            one_row = "A [" + str(len(v)) + "]["
            for t in v:
                if isinstance(t, str):
                    one_row += t + ','
                else:
                    one_row += str(t) + ','
            one_row += ']'
            log(one_row.replace("\r", "").replace("\n", ""))
        for v in all_changes[k][1]:
            one_row = "M [" + str(len(v)) + "]["
            for t in v:
                if isinstance(t, str):
                    one_row += t + ','
                else:
                    one_row += str(t) + ','
            one_row += ']'
            log(one_row.replace("\r", "").replace("\n", ""))
        for v in all_changes[k][3]:
            one_row = "D [" + str(v) + ']'
            log(one_row.replace("\r", "").replace("\n", ""))

    log("\n-----------------------------------------")
    log("需要忽略的文件（文件冲突）:")
    if len(ignore_changes) == 0:
        log("[无]")
    else:
        for k in ignore_changes:
            log(k + " ")

    log("\n-----------------------------------------")
    log("正在Merge所有文件差异...")
    dt = datetime.datetime.now()
    for k in all_changes:
        pre_merge_branch_file(k, all_changes[k])
        merge_branch_file(k, all_changes[k])
    log("Write Excel Spend:" + str(datetime.datetime.now() - dt)[:7])

    log("\n----------------------------------------------------------------------------------")
    log("正在做最后的检查...")
    cur_branch = branch_dir + branch_sub
    cur_branch = cur_branch.replace("\\", "/")
    log("branch dir:" + cur_branch)
    modify_files = svn_optr.get_local_dir_modify_files(cur_branch)
    for f in modify_files:
        f = f.replace("\\", "/")
        log("Checking: " + f)
        d_new = read_excel_xlrd(f)
        d_svn = get_trunk_url_file_data(0, f.replace(branch_dir.replace("\\", "/"), ""))
        if d_new[0] == d_svn[0]:
            new_file_data = d_new[1]
            old_file_data = d_svn[1]
            if len(new_file_data) == len(old_file_data):
                # both in new_file and in old_file
                both_in = dict([(i, new_file_data[i]) for i in filter(lambda k:k in new_file_data, old_file_data.keys())])
                # add rows
                add_result = list(
                    [new_file_data[i] for i in filter(lambda k: k not in old_file_data, new_file_data.keys())])
                # delete rows
                delete_result = list([i for i in filter(lambda k: k not in new_file_data, old_file_data.keys())])
                if len(add_result) == 0 and len(delete_result) == 0 and len(both_in) == len(new_file_data):
                    all_equal = True
                    for k in both_in:
                        if both_in[k] != old_file_data[k]:
                            all_equal = False
                            break
                    if all_equal:
                        log("svn revert -R " + f)
                        svn_optr.get_revert_local_file(f)
                        svn_optr.update_local_file(f)

    log("\n----------------------------------------------------------------------------------")
    log("二师兄，开饭啦！！！")
    log(EVENT_FINISHED)
    g_finished_ok = True


def format_configs(lines):
    """格式化读取的配置文件"""
    global trunk_url
    global trunk_sub
    global branch_dir
    global branch_sub

    def remove_last_if(s):
        if s.endswith("/") or s.endswith("\\"):
            s = s[:-1]
        return s.replace("\\", "/")

    def add_first_if(s):
        if (not s.startswith("/")) and (not s.startswith("\\")):
            s = "/" + s
        return s.replace("\\", "/")

    trunk_url = remove_last_if(lines[0])
    trunk_sub = add_first_if(lines[1])
    trunk_sub = remove_last_if(trunk_sub)
    branch_dir = remove_last_if(lines[2])
    # make sure "D:" is upper
    branch_dir = branch_dir[0:2].upper() + branch_dir[2:]
    branch_sub = add_first_if(lines[3])
    branch_sub = remove_last_if(branch_sub)

    ui.lineEdit_trunk.setText(trunk_url)
    ui.lineEdit_trunk_postfix.setText(trunk_sub)
    ui.lineEdit_branch.setText(branch_dir)
    ui.lineEdit_branch_postfix.setText(branch_sub)


def save_config():
    lines = dict()
    lines[0] = str(ui.lineEdit_trunk.text())
    lines[1] = str(ui.lineEdit_trunk_postfix.text())
    lines[2] = str(ui.lineEdit_branch.text())
    lines[3] = str(ui.lineEdit_branch_postfix.text())
    format_configs(lines)
    if os.path.isfile("config.txt"):
        os.remove("config.txt")
    wf = open("config.txt", 'w')
    wf.write(trunk_url + '\n')
    wf.write(trunk_sub + '\n')
    wf.write(branch_dir + '\n')
    wf.write(branch_sub + '\n')
    wf.close()


# def multi_thread(keys):
#     global g_mp
#     global g_finished_ok
#
#     pythoncom.CoInitialize()
#     dt = datetime.datetime.now()
#     try:
#         main(keys)
#     finally:
#         if not g_finished_ok:
#             log("oh, SOMETHING WRONG!")
#             log(EVENT_FINISHED)
#     log("Time spent:" + str(datetime.datetime.now() - dt)[:7])
#
#     # when end thread
#     pythoncom.CoUninitialize()
#     # g_mp.quit()
#     # g_mp = None


def multi_process(keys):
    global g_mp
    global g_finished_ok

    pythoncom.CoInitialize()
    dt = datetime.datetime.now()
    try:
        main(keys)
    finally:
        if not g_finished_ok:
            log("oh, SOMETHING WRONG!")
            log(EVENT_FINISHED)
    log("Time spent:" + str(datetime.datetime.now() - dt)[:7])

    # when end thread
    pythoncom.CoUninitialize()
    # g_mp.quit()
    # g_mp = None


def get_ui_branch_version():
    str_version = str(ui.lineEdit_startver.text())
    int_version = to_int(str_version)
    if int_version is None or int_version < 0:
        return None
    return int_version


def just_do_it():
    global svn_optr
    global g_mp
    global g_branch_first_ver
    save_config()

    g_branch_first_ver = get_ui_branch_version()
    # 声明svnoperator的对象
    svn_optr = SvnOperator(trunk_url, trunk_sub)
    # 获取UI中单号输入框内的信息
    input_text = str(ui.textEdit.toPlainText())
    # 正则提取单号部分
    keys = re.compile('[A-Z]+-[0-9]+').findall(input_text)

    if len(keys) == 0:
        QtGui.QMessageBox.information(Dialog, u'Excel Merge Tool', u"没有检测到单号输入呢~")
        return
    keys = list(set(keys))
    ui.textEdit_status.setText(u"已提取单号:")
    # 将提取的单号循环输出到log文本框
    for k in keys:
        ui.textEdit_status.append(k)
    ui.textEdit_status.append('--------------------------------------------')

    g_mp = QMultiProcess(keys)
    g_mp.update_ui.connect(log_ui)
    g_mp.start()
    ui.pushButton_go.setEnabled(False)


def clean_xls():
    global g_mp

    ui.textEdit_status.setText(u"准备清理...")

    g_mp = QMultiThread([])
    g_mp.update_ui.connect(log_ui)
    g_mp.start()
    ui.pushButton_clean.setEnabled(False)
    ui.pushButton_go.setEnabled(False)
    ui.pushButton_save.setEnabled(False)


def init():
    global Application_Excel_Version

    pythoncom.CoInitialize()
    Application_Excel_Version = get_suitable_excel_version()
    if Application_Excel_Version is None:
        QtGui.QMessageBox.information(Dialog, u'Excel Merge Tool', u"1. 请关闭当前已打开Excel文档(必要时打开任务管理器杀掉Excel进程)。\n2. 请确认已安装Office(且版本>=Office 2007)！")
        sys.exit(1)
        return

    # 若目录下存在配置文件，则读取配置到lines并格式化
    if os.path.isfile("config.txt"):
        rf = open("config.txt", 'r')
        line = rf.readline()
        line_num = 0
        lines = dict()
        while line:
            line = line.replace("\t", "")
            line = line.replace("\r", "")
            line = line.replace("\n", "")
            line = line.replace("\\", "/")
            lines[line_num] = line
            line_num += 1
            line = rf.readline()
        rf.close()
        format_configs(lines)
    else:
        ui.lineEdit_trunk.setText(trunk_url)
        ui.lineEdit_trunk_postfix.setText(trunk_sub)
        ui.lineEdit_branch.setText(branch_dir)
        ui.lineEdit_branch_postfix.setText(branch_sub)

    # for debug
    ui.textEdit.setText("SDXL-3342【11.02 配置】日常配置")

    # buttons
    QObject.connect(ui.pushButton_save, SIGNAL("clicked()"), save_config)
    QObject.connect(ui.pushButton_go, SIGNAL("clicked()"), just_do_it)
    QObject.connect(ui.pushButton_clean, SIGNAL("clicked()"), clean_xls)


def un_init():
    if Application_Excel_Version:
        win32.gencache.EnsureModule('{00020813-0000-0000-C000-000000000046}', 0, 1, int(Application_Excel_Version[0]))
        excel = win32.Dispatch('Excel.Application.' + str(Application_Excel_Version[1]))
        if excel:
            excel.Application.Quit()
    pythoncom.CoUninitialize()


if __name__ == "__main__":
    app = QtGui.QApplication(sys.argv)
    Dialog = QtGui.QDialog()
    # add minimize button
    Dialog.setWindowFlags(Dialog.windowFlags()
                          | QtCore.Qt.WindowMinimizeButtonHint
                          | QtCore.Qt.WindowSystemMenuHint)
    Dialog.setWindowIcon(QIcon("favicon.ico"))
    ui = Ui_Dialog()
    ui.setupUi(Dialog)
    init()
    Dialog.show()
    exit_code = app.exec_()
    un_init()
    sys.exit(exit_code)
