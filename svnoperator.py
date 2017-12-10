# coding:utf-8
import re
import subprocess


class SvnOperator(object):
    __svn_trunk = ""
    __svn_sub = ""

    def __init__(self, trunk, sub):
        self.svn_trunk = trunk
        self.svn_sub = sub

    def get_commits_by_key(self, branch_first_ver, key):
        svn_param = "svn log -v -r" + \
                    str(branch_first_ver) + ":HEAD --stop-on-copy --search=" + key + " \"" + \
                    self.svn_trunk + self.svn_sub + "\""
        svn_logs = subprocess.Popen(
            svn_param,
            stderr=subprocess.STDOUT,
            stdout=subprocess.PIPE,
            shell=True).communicate()[0].decode('gbk')
        logs = re.compile('r([\d]+)([\s\S]*?)\r\n-----').findall(svn_logs)

        commit_list = []
        for l in logs:
            one_commit = self.__process_one_commit(self, int(l[0]), l[1])
            if one_commit is not None:
                commit_list.append(one_commit)

        return commit_list

    @staticmethod
    def get_repository_oldest_ver(rep_url):
        logs = "svn log --stop-on-copy -q -r0:HEAD -l1 \"" + rep_url + "\""
        log_result = subprocess.Popen(
            logs,
            stderr=subprocess.STDOUT,
            stdout=subprocess.PIPE, shell=True).communicate()[0].decode('gbk')
        ver_list = re.compile('r([\d]+)').findall(log_result)
        if len(ver_list) > 0:
            return int(ver_list[len(ver_list) - 1])
        return 0

    def get_file_ver_before(self, first_ver, ver, filename):
        svn_logs = subprocess.Popen(
            "svn info \"" + self.svn_trunk + filename + "\"@" + str(ver - 1),
            stderr=subprocess.STDOUT,
            stdout=subprocess.PIPE, shell=True).communicate()[0]
        svn_logs = svn_logs.decode('gbk')

        if "svn: E" in svn_logs and ("No such revision" in svn_logs or "not found" in svn_logs):
            print("[WARNING] No Before Version! r:" + ver.encode() + ":" + filename.encode())
            return ver - 1

        logs = re.compile('([\d]+)\r\n').findall(svn_logs)
        if len(logs) > 1:
            if int(logs[len(logs) - 1]) == ver:
                return ver - 1
            if first_ver > int(logs[len(logs) - 1]):
                return ver - 1
            return int(logs[len(logs) - 1])

        print("[WARNING] No Before Version! r:" + str(ver) + ":" + filename)
        return ver - 1

    @staticmethod
    def get_local_dir_modify_files(local_rep_dir):
        log_result = subprocess.Popen(
            "svn st -q \"" + local_rep_dir + "\"",
            stderr=subprocess.STDOUT,
            stdout=subprocess.PIPE, shell=True).communicate()[0]
        log_result = log_result.decode('gbk')
        changelist = re.compile('[M][ ]*([a-zA-Z]:[\s\S]*?)\r\n').findall(log_result)
        return changelist

    @staticmethod
    def update_local_repository(local_rep_dir):
        subprocess.Popen(
            "svn up \"" + local_rep_dir.decode('utf8').encode('gbk') + "\"",
            stderr=subprocess.STDOUT,
            stdout=subprocess.PIPE, shell=True).communicate()

    @staticmethod
    def is_local_repository_dirty(local_rep_dir):
        logs = "svn st -q \"" + local_rep_dir.decode('utf8').encode('gbk') + "\""
        log_result = subprocess.Popen(
            logs,
            stderr=subprocess.STDOUT,
            stdout=subprocess.PIPE, shell=True).communicate()[0]
        if log_result == "":
            return False
        return True

    @staticmethod
    def get_revert_local_file(local_rep_dir):
        subprocess.Popen(
            "svn revert -R \"" + local_rep_dir.decode('utf8').encode('gbk') + "\"",
            stderr=subprocess.STDOUT,
            stdout=subprocess.PIPE, shell=True).communicate()

    @staticmethod
    def update_local_file(local_rep_dir):
        subprocess.Popen(
            "svn up \"" + local_rep_dir.decode('utf8').encode('gbk') + "\"",
            stderr=subprocess.STDOUT,
            stdout=subprocess.PIPE, shell=True).communicate()
            
    def download_url_file(self, ver, filename, save_file):
        which_ver = str(ver)
        if ver == 0:
            which_ver = "HEAD"

        subprocess.Popen(
            "svn cat \"" + self.svn_trunk + filename + "\"@" +
            which_ver + " > \"" + save_file + "\"",
            stderr=subprocess.STDOUT,
            stdout=subprocess.PIPE, shell=True).communicate()

    @staticmethod
    def add_local_file(local_file):
        ret_string = subprocess.Popen(
            "svn add --parents \"" + local_file.decode('utf8').encode('gbk') + "\"",
            stderr=subprocess.STDOUT,
            stdout=subprocess.PIPE, shell=True).communicate()[0]
        return ret_string.decode('gbk').encode('utf-8')

    @staticmethod
    def delete_local_file(local_file):
        ret_string = subprocess.Popen(
            "svn delete --force \"" + local_file.decode('utf8').encode('gbk') + "\"",
            stderr=subprocess.STDOUT,
            stdout=subprocess.PIPE, shell=True).communicate()[0]
        return ret_string.decode('gbk').encode('utf-8')

    # *************************** private method *************************
    @staticmethod
    def __process_one_commit(self, ver, commit):
        modifies = re.compile('(   [AMD] /[\s\S]*?\r\n)').findall(commit)
        log_list = re.compile('\r\n\r\n([\s\S]*)?').findall(commit)

        # get log message
        log_text = ""
        for l in log_list:
            assert isinstance(l, str)
            log_text += l

        log_text = log_text.replace("\r", "")
        log_text = log_text.replace("\n", "")
        log_text = log_text.replace("\t", "")
        log_text = log_text.replace("\'", "\'\'")
        log_text = log_text.replace("\"", "\"\"")

        changelist = []
        for l in modifies:
            mt = self.__check_modify_file_type(l)
            if mt is not None:
                changelist.append(mt)

        if len(changelist) > 0:
            return ver, log_text, changelist
        return None

    @staticmethod
    def __check_modify_file_type(pre_file):
        if pre_file.startswith("   M "):
            modify_type = 0  # for modify
        elif pre_file.startswith("   A "):
            modify_type = 1  # for add
        elif pre_file.startswith("   D "):
            modify_type = 2  # for delete
        else:
            assert False, "Unknown svn operator type:" + pre_file
            return None

        file_name = re.compile('   [AMD] ([\s\S]*?)\r\n').findall(pre_file)[0]
        if file_name.endswith((".xls", ".xlsx", ".xlsm")):
            return file_name, modify_type

        # (+)bugfix
        # 改变的路径:
        #    A /trunk_xlsdir/1.xlsx (从 /trunk_xlsdir/Test-add_dir.xlsx:9)
        # (-)bugfix
        if file_name.endswith(")"):
            file_names = re.compile('([\s\S]*?) \(').findall(file_name)
            if len(file_names) > 0:
                file_name = file_names[0]
            if file_name.endswith((".xls", ".xlsx", ".xlsm")):
                return file_name.decode('gbk').encode('utf-8'), modify_type
        return None
