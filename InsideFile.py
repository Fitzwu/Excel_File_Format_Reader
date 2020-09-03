import os
import json
import re

m_path = os.path.dirname(__file__)
m_jsonFile_path = os.path.join(m_path,"DailyReport.json")
try:
    jsonFile = open(m_jsonFile_path,'r' ,encoding='utf-8')
    pass
except:
    print("\n 请确认配置文件 DailyReport.json 存在")
    pass
m_jsonctx = json.load(jsonFile) #is a dict


def ctxget(ctx, name):
    return ctx.get(name)

# 解析 json 文件中 range 
def ParseRange(range_string):
    rg = re.findall(r'\d+',range_string)
    if not len(rg) == 2:
        print("范围输入有错!\n")
        return None
    return rg

class OwnSheet:
    def __init__(self, range_list, SheetName = ""):
        self.range_list = range_list
        self.SheetName = SheetName
        


class config:
    __sheet = None

    def loadconf(self, level):
        num = m_jsonctx['levels']
        if level > num:
            print("配置文件设置错误！\n")
            return None
        level_name = "level_{}".format(level)
        mctx = ctxget(m_jsonctx, level_name)
        return mctx

    def makeSheetList(self):
        self.SheetNum = ctxget(self.newctx,"SheetNum")
        for i in range(self.SheetNum):
            sheet_Flag = "sheet_{}".format(i + 1)
            sheet_ctx = ctxget(self.newctx, sheet_Flag)
            sheetName = ctxget(sheet_ctx, "SheetName")
            str_range = ctxget(sheet_ctx, "range")

            range_list = []
            
            if not str_range == None:
                num_range = str_range.split("+")
                num = len(num_range)
                for i in range(num):
                    strr = num_range[i]

                    range_list.append(ParseRange(strr))
                    pass
            else:
                range_list = None

            sheet = OwnSheet(range_list, sheetName)
            self.__sheet.append(sheet)


    def __init__(self, level = 1):
        self.__conf_file = ''
        self.level = level
        self.newctx = self.loadconf(self.level)
        self.__sheet = []
        if self.newctx == None:
            print("配置文件设置错误！\n")
            return None
        else:
            self.makeSheetList()

    def getOwnsheet(self):
        return self.__sheet
        pass

    

class InsideFile:
    name = ''
    level = 1 # 1表示是第一层级 2表示是第二层级 以此类推
    __path = ''
    conf:config
    def __init__(self, file, level, path):
        self.name = file
        self.level = level
        self.__path = path
        self.conf = config(self.level)

    def File_Path(self):
        return self.__path