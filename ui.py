# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'main.ui'
#
# Created: Tue Oct 31 11:13:24 2017
#      by: PyQt4 UI code generator 4.11
#
# WARNING! All changes made in this file will be lost!

from PyQt4 import QtCore, QtGui

try:
    _fromUtf8 = QtCore.QString.fromUtf8
except AttributeError:
    def _fromUtf8(s):
        return s

try:
    _encoding = QtGui.QApplication.UnicodeUTF8
    def _translate(context, text, disambig):
        return QtGui.QApplication.translate(context, text, disambig, _encoding)
except AttributeError:
    def _translate(context, text, disambig):
        return QtGui.QApplication.translate(context, text, disambig)

class Ui_Dialog(object):
    def setupUi(self, Dialog):
        Dialog.setObjectName(_fromUtf8("Dialog"))
        Dialog.resize(680, 530)
        Dialog.setMinimumSize(QtCore.QSize(680, 530))
        Dialog.setMaximumSize(QtCore.QSize(680, 530))
        self.lineEdit_trunk = QtGui.QLineEdit(Dialog)
        self.lineEdit_trunk.setGeometry(QtCore.QRect(109, 12, 291, 20))
        self.lineEdit_trunk.setObjectName(_fromUtf8("lineEdit_trunk"))
        self.lineEdit_branch = QtGui.QLineEdit(Dialog)
        self.lineEdit_branch.setGeometry(QtCore.QRect(109, 45, 291, 20))
        self.lineEdit_branch.setObjectName(_fromUtf8("lineEdit_branch"))
        self.lineEdit_trunk_postfix = QtGui.QLineEdit(Dialog)
        self.lineEdit_trunk_postfix.setGeometry(QtCore.QRect(441, 12, 231, 20))
        self.lineEdit_trunk_postfix.setObjectName(_fromUtf8("lineEdit_trunk_postfix"))
        self.label = QtGui.QLabel(Dialog)
        self.label.setGeometry(QtCore.QRect(20, 16, 78, 12))
        self.label.setObjectName(_fromUtf8("label"))
        self.label_2 = QtGui.QLabel(Dialog)
        self.label_2.setGeometry(QtCore.QRect(410, 16, 31, 16))
        self.label_2.setObjectName(_fromUtf8("label_2"))
        self.label_3 = QtGui.QLabel(Dialog)
        self.label_3.setGeometry(QtCore.QRect(20, 49, 84, 12))
        self.label_3.setObjectName(_fromUtf8("label_3"))
        self.label_4 = QtGui.QLabel(Dialog)
        self.label_4.setGeometry(QtCore.QRect(410, 49, 31, 16))
        self.label_4.setObjectName(_fromUtf8("label_4"))
        self.lineEdit_branch_postfix = QtGui.QLineEdit(Dialog)
        self.lineEdit_branch_postfix.setGeometry(QtCore.QRect(441, 45, 231, 20))
        self.lineEdit_branch_postfix.setObjectName(_fromUtf8("lineEdit_branch_postfix"))
        self.textEdit = QtGui.QTextEdit(Dialog)
        self.textEdit.setGeometry(QtCore.QRect(20, 98, 551, 81))
        self.textEdit.setObjectName(_fromUtf8("textEdit"))
        self.pushButton_go = QtGui.QPushButton(Dialog)
        self.pushButton_go.setGeometry(QtCore.QRect(590, 136, 75, 41))
        self.pushButton_go.setObjectName(_fromUtf8("pushButton_go"))
        self.pushButton_save = QtGui.QPushButton(Dialog)
        self.pushButton_save.setGeometry(QtCore.QRect(590, 98, 75, 24))
        self.pushButton_save.setObjectName(_fromUtf8("pushButton_save"))
        self.textEdit_status = QtGui.QTextEdit(Dialog)
        self.textEdit_status.setGeometry(QtCore.QRect(20, 200, 641, 311))
        self.textEdit_status.setReadOnly(True)
        self.textEdit_status.setObjectName(_fromUtf8("textEdit_status"))
        self.label_5 = QtGui.QLabel(Dialog)
        self.label_5.setGeometry(QtCore.QRect(22, 78, 101, 16))
        self.label_5.setObjectName(_fromUtf8("label_5"))
        self.label_6 = QtGui.QLabel(Dialog)
        self.label_6.setGeometry(QtCore.QRect(20, 182, 41, 16))
        self.label_6.setObjectName(_fromUtf8("label_6"))
        self.pushButton_clean = QtGui.QPushButton(Dialog)
        self.pushButton_clean.setGeometry(QtCore.QRect(574, 69, 91, 24))
        self.pushButton_clean.setObjectName(_fromUtf8("pushButton_clean"))
        self.lineEdit_startver = QtGui.QLineEdit(Dialog)
        self.lineEdit_startver.setGeometry(QtCore.QRect(499, 74, 71, 20))
        self.lineEdit_startver.setObjectName(_fromUtf8("lineEdit_startver"))
        self.label_7 = QtGui.QLabel(Dialog)
        self.label_7.setGeometry(QtCore.QRect(427, 74, 71, 20))
        self.label_7.setObjectName(_fromUtf8("label_7"))

        self.retranslateUi(Dialog)
        QtCore.QMetaObject.connectSlotsByName(Dialog)
        Dialog.setTabOrder(self.textEdit, self.pushButton_go)
        Dialog.setTabOrder(self.pushButton_go, self.lineEdit_trunk)
        Dialog.setTabOrder(self.lineEdit_trunk, self.lineEdit_trunk_postfix)
        Dialog.setTabOrder(self.lineEdit_trunk_postfix, self.lineEdit_branch)
        Dialog.setTabOrder(self.lineEdit_branch, self.lineEdit_branch_postfix)

    def retranslateUi(self, Dialog):
        Dialog.setWindowTitle(_translate("Dialog", "Excel Merge v2017.01(r299)", None))
        self.label.setText(_translate("Dialog", "主干svn根目录", None))
        self.label_2.setText(_translate("Dialog", "后缀", None))
        self.label_3.setText(_translate("Dialog", "分支本地根目录", None))
        self.label_4.setText(_translate("Dialog", "后缀", None))
        self.pushButton_go.setText(_translate("Dialog", "执行", None))
        self.pushButton_save.setText(_translate("Dialog", "保存配置", None))
        self.label_5.setText(_translate("Dialog", "输入待处理单号", None))
        self.label_6.setText(_translate("Dialog", "状态", None))
        self.pushButton_clean.setText(_translate("Dialog", "清理表格空行", None))
        self.label_7.setText(_translate("Dialog", "(海外版)ver", None))


if __name__ == "__main__":
    import sys
    app = QtGui.QApplication(sys.argv)
    Dialog = QtGui.QDialog()
    ui = Ui_Dialog()
    ui.setupUi(Dialog)
    Dialog.show()
    sys.exit(app.exec_())

