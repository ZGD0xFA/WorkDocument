import xlrd
import os


xlsfile = "主机列表.xls"
accesslog = "access.log"


class Solution:
    def __init__(self):
        pass

    def work(self):
        # 解析配置文件
        ret = self.analyzexlsfile()
        if not isinstance(ret, dict()):
            print("analyzexlsfile is false")
            return
        print(ret)

        retaccess = self.analyzeaccessfile()

    # function: analyzexlsfile
    # description: 解析主机管理
    #
    def analyzexlsfile(self):
        if not os.path.exists(xlsfile):
            print('Error: make sure {0} is exist!'.format(xlsfile))
            return False
        data = xlrd.open_workbook(xlsfile)
        table = data.sheet_by_index(0)
        row = table.row_values(0, start_colx=0, end_colx=None)
        cols = dict()
        for i in range(0, len(row)):
            if row[i] == "外网IP":
                cols[row[i]] = table.col_values(i, start_rowx=1)
            if row[i] == "操作系统":
                cols[row[i]] = table.col_values(i, start_rowx=1)
        if "外网IP" not in cols and "操作系统" not in cols:
            print("Content is not exist!")
            return False
        return cols


    def analyzeaccessfile(self):
        if not os.path.exists(xlsfile):
            print('Error: make sure {0} is exist!'.format(accesslog))
            return False


if __name__ == "__main__":
    option = Solution()
    option.analyzexlsfile()
    pass
