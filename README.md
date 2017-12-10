# svn_excel_merger
svn下excel合分支工具

HOME: http://git.oschina.net/kylescript/ExcelMerger


###depends
python 3.4.4(32bit) + Windows 10 x64 + PyQt4-4.11-gpl-Py3.4-Qt4.8.6-x32.exe

pip3 list:
altgraph (0.14)
future (0.16.0)
macholib (1.8)
pefile (2017.9.3)
pip (7.1.2)
pyinstaller (3.3)
pypiwin32 (219)
pywin32 (221)
setuptools (18.2)
xlrd (1.0.0)

###原理描述
1. 将excel中每一行当成一个最小单元，以每行第一列id作为索引
2. 本工具不支持文件列发生变化（增加或者减少），会直接报错误，跳过此文件
3. 根据工作流分为提取差异阶段，合入分支阶段

