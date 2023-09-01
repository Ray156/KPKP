# %%
# from xlwings import App,Book
import xlwings as xw
from difflib import get_close_matches as gcm
import numpy as np


# %%
##
## 这里可以修改data2.xlsx的数据区域，不可修改列范围，仅支持修改行范围
## 默认数据范围为A12:NG23
## 12~23表示从中关村分行到通州分行的行号
## A到NG表示整个表格的列范围
ROW1 = 12 # 起始行号（中关村分行）
ROW2 = 23 # 结束行号（通州分行）
DATA_RANGE = f'A{ROW1}:NG{ROW2}'

PARTMENTS_A = ["中关村分行","东四支行","月坛支行","朝阳支行","长安支行","丰台支行","城建支行",
             "铁道支行","宣武支行","安华支行","分行营业部","通州分行"]
DATA = "data/"
PATH1 = "rule.xlsx"
PATH2 = "task2.xlsx"
MAX_ROW = 120
START_ROW = 3
A = ['经营效益（18）','发展转型（26）','风险合规（40）','社会责任（16）']
NUM = 12

OUTPUT_PATH = '各分支行得分情况表.xlsx'

# %%
# rule = pd.read_excel('rule.xlsx',sheet_name=None)
app = xw.App(visible=False, add_book=True)
app.display_alerts = False

# rulebook = None
try:
    rulebook.close()
except:
    rulebook = app.books.open(DATA + PATH1)

# %%
# level string description
# 1.xxx
# 1.1xxx
# 1.1.1xxx
# (1)xxx 括号可能是中文括号，也可能是英文括号
# a.xxx
# b.xxx
# (2)xxx
# 没有序号但不为None的行，自动归为上一级的子项

# 输出python代码

# 定义一个类，表示一个层级结构
class Level:
    def __init__(self, name, level=0,parent=None):
        self.name = name # 层级的名称
        self.level = level # 层级的深度
        self.parent = parent # 层级的父节点
        self.children = [] # 层级的子节点列表
        self.score = np.zeros(len(PARTMENTS_A),dtype=object) # 层级的分数
        self.index = 0 # 层级的序号
        self.calculateScore = None # 层级的分数计算函数

    def add_child(self, child):
        # 添加一个子节点
        self.children.append(child)
    
    def search_by_name(self, name):
        # 搜索树中的节点，返回第一个名称匹配的节点
        # if self.name == name:
        temp = gcm(name,[self.name],1,cutoff=0.7)
        # print(self.name,temp)
        if temp!=[]:
            return self
        for child in self.children:
            result = child.search_by_name(name)
            if result is not None:
                return result
        

    def print_level(self, indent=0):
        # 打印层级的名称和子节点，缩进表示层级深度
        print(" " * indent + self.name)
        for child in self.children:
            child.print_level(indent + 4)
        
    def getLeaf(self):
        # 获取所有叶节点
        if self.children == []:
            return [self]
        else:
            returnList = []
            for child in self.children:
                returnList.extend(child.getLeaf())
            return returnList

# 从字符串描述中解析出层级结构，输入为字符串列表，输出为根节点
def parse_level(description,roots=[Level('root',0)]):

    # 按行遍历字符串描述
    lastLevel = 0
    level = 0
    topItemNo = 0
    root = roots[0]
    returnList = []
    index = START_ROW
    for line in description:
        # 去掉行首和行尾的空格
        line = line.strip()
        # 如果行为空，跳过
        if not line:
            continue

        # 基础层级深度为root的深度
        level = root.level
        
        if line[0].isdigit():
            strs = line.split('.')
            # print(strs)
            if not strs[1][0].isdigit():
                level += 1
                if strs[0]=='1':
                    # 创建一个根节点
                    root = roots[topItemNo]
                    returnList.append(root)
                    # 创建一个当前节点，初始为根节点
                    current = root
                    # 创建一个栈，用于存储父节点
                    stack = [root]
                    topItemNo += 1
            else:
                level += line.count('.')+1
        elif line[0].isalpha():
            level += 5
        elif line[0] in "(（":
            level += 4
        else:
            level += 6

        # 创建一个层级结构
        current = Level(line,level)
        current.index = index
        # print(line,level,lastLevel,sep='\t')
        # for i in stack:
        #     print(i.name,end='\t')
        # print()
        # print()
        if level > lastLevel:
            current.parent = stack[-1]
            stack[-1].add_child(current)
            stack.append(current)

        elif level == lastLevel:

            stack.pop()
            # print(a.name)
            current.parent = stack[-1]
            stack[-1].add_child(current)
            stack.append(current)
        else:

            while True:
                lastLevel = stack[-1].level
                if level <= lastLevel:
                    a = stack.pop()
                else:   
                    break
            # print(a.name)
            current.parent = stack[-1]
            stack[-1].add_child(current)
            stack.append(current)
        lastLevel = level
        index += 1

    # 返回根节点
    return returnList



# %%
A = ['经营效益（18）','发展转型（26）','风险合规（40）','社会责任（16）']
# rule：sheet1 主卡
itemName = list(rulebook.sheets[0].range(f'B{START_ROW}:B{MAX_ROW}').value)
bcd = list(filter(None,itemName))
root = Level('总分',-1)
主卡  = parse_level(bcd,[Level(i,0) for i in A])
root.children = 主卡

itemName = list(rulebook.sheets[1].range(f'A{START_ROW}:A{MAX_ROW}').value)
bcd = list(filter(None,itemName))
# 对公业务转型 = root.search_by_name('对公业务转型')
root1 = Level('对公业务转型',0)
对公附卡  = parse_level(bcd,[root1])
# root1.children = 对公附卡

itemName = (rulebook.sheets[2].range(f'A{START_ROW}:A{MAX_ROW}').value)
bcd = list(filter(None,itemName))
# 个人业务转型 = root.search_by_name('个人业务转型')
root2 = Level('个人业务转型',0)
个人附卡  = parse_level(bcd,[root2])
# root2.children = 个人附卡

root.print_level()
root1.print_level()
root2.print_level()
# root.print_level()

# %%

# 获得root的所有叶子节点 

leaf = root.getLeaf()
leaf1 = root1.getLeaf()
leaf2 = root2.getLeaf()
# print([(i.name,i.index) for i in leaf])
count = 0
for i in leaf:
    print(count,i.name,i.index)
    count += 1

# %%
# 预设多种规则，每种规则对应一个函数，函数的输入为一个叶节点对应的数据list，输出为该叶节点的分数
# 例如：规则1：当某二级行增长额大于组内中位值时，30×（某二级行增长额－组内中位增长额）/（组内最大增长额－组内中位增长额）；当某二级行增长额小于组内中位值时，30×（某二级行增长额－组内中位增长额）/（组内中位增长额－组内最小增长额）。
def formula1(data,w=30):
  # 计算列表中的中位值、最大值、最小值
  data = np.array(data)
  if data.all() == 0:
      return np.zeros_like(data)
  median = data.mean()
  max_value = data.max()
  min_value = data.min()
  print(median,max_value,min_value)
  # data > median, = 30 * (growth - median) / (max_value - median)
  # data < median, = 30 * (growth - median) / (median - min_value)
  divisor = np.zeros_like(data)
  divisor[data >= median] = max_value - median
  divisor[data < median] = median - min_value
  score_list = w * (data - median) / divisor

  # print(score_list)
  return score_list

# # test =  -1,673.78 , 2,592.99 , -488.87 , -626.16 , -665.19 , -807.59 , -38.49 , -4,178.16 , 862.74 , 24.76 , 888.86 , -27.99
# test_data = [-1673.78, 2592.99, -488.87, -626.16, -665.19, -
#              807.59, -38.49, -4178.16, 862.74, 24.76, 888.86, -27.99]
# print(formula1(test_data))

# 规则2，适用于如A.3


def formula2(data):
  # 二级行当年成本收入比较上年提高的，该指标得0分。
  # 1.绝对值:比上年变动小于0时，得1分，
  data = np.array(data)
  score_list = np.where(data > 0, 0, 1).astype('float')

  # 2.变动值:将负变动值从小到大排名，排名第一的得1分，每下降一名减0.02分。
  rank = np.argsort(data[data < 0])
  score_list[data < 0] += (1 - rank * 0.02)
  return score_list





# %%
# dataxw = None
try:
    dataxw.close()
except:
    pass
dataxw = app.books.open(DATA + PATH2)
# xw1  = app.books.open('task2.xlsx')
# xw2  = app.books.open("000-2023年二季度二级分支行KPI完成情况（A组数据） - 经营效益+对公副卡——新行员2.xlsx")
# s1 = dataxw.sheets[0].range('A7:zz7').value

# 取A10:NG23转置后的数据
data = dataxw.sheets[0]

print(f'选定数据区域：{DATA_RANGE}\n')
data = data[DATA_RANGE].options(transpose=True).value
# any str to 0
data = np.array(data)
# data = np.append(data, np.array(['a']), axis=-1)
data1 = data[1:].view()
# print(data1.astype(str)==data1)
data1[data1.astype(str)==data1] = 0
# Ray156[np.where(Ray156.astype(str) == Ray156)] = 0
# data = np.where(data.astype(str) != data, data, 0)
# data = np.array(data)



print(data)


# %% 
# 经营效益
# 0 1.1经济增加值增长 4
# rule = rulebook.sheets[0].range(f'A4:H4').value
# inData1 = [55407.84 ,10512.25,20736.28,18369.57,-2836.72,-5106.09,3399.44,11204.49,20191.23,-9338.28,20638.48,-1380.51]
# inData2 = [111.72,20.98,53.34,37.92,-11.17,-13.43,13.84,426.68,139.90,-33.59,20.54,-4.17]
leafNo = 0
inData1 = np.array(data[3])
inData2 = np.array(data[4])
benchmark = 100
weight = 4
leaf[leafNo].score = (benchmark+(formula1(inData1)+formula1(inData2))/2) * weight / 100
print(f"输入数据({leaf[leafNo].name}):",inData1,inData2,sep='\n')
print(leafNo,leaf[leafNo].score,end='\n\n')

# 1 1.2人均经济增加值 5
# inData1 = [55407.84 ,10512.25,20736.28,18369.57,-2836.72,-5106.09,3399.44,11204.49,20191.23,-9338.28,20638.48,-1380.51]
leafNo = 1
inData1 = np.array(data[6])
benchmark = 100
weight = 2
leaf[leafNo].score = (benchmark+formula1(inData1)) * weight / 100
print(f"输入数据({leaf[leafNo].name}):",inData1,sep='\n')
print(leafNo,leaf[leafNo].score,end='\n\n')

# 2 2.1拨备前利润增长 7
leafNo = 2
inData1 = np.array(data[10])
inData2 = np.array(data[11])
# inData1 = [55407.84 ,10512.25,20736.28,18369.57,-2836.72,-5106.09,3399.44,11204.49,20191.23,-9338.28,20638.48,-1380.51]
# inData2 = [111.72,20.98,53.34,37.92,-11.17,-13.43,13.84,426.68,139.90,-33.59,20.54,-4.17]
benchmark = 100
weight = 4
leaf[leafNo].score = (benchmark+(formula1(inData1,50)+formula1(inData2,50))/2) * weight / 100
print(f"输入数据({leaf[leafNo].name}):",inData1,inData2,sep='\n')
print(leafNo,leaf[leafNo].score,end='\n\n')


# 3 2.2人均拨备前利润增长 8
leafNo = 3
inData1 = np.array(data[13])
# inData1 = [55407.84 ,10512.25,20736.28,18369.57,-2836.72,-5106.09,3399.44,11204.49,20191.23,-9338.28,20638.48,-1380.51]
benchmark = 100
weight = 1
leaf[leafNo].score = (benchmark+formula1(inData1)) * weight / 100
print(f"输入数据({leaf[leafNo].name}):",inData1,sep='\n')
print(leafNo,leaf[leafNo].score,end='\n\n')

# 4 3.成本收入比 9
leafNo = 4
inData1 = np.array(data[16])
benchmark = 0
weight = 2
leaf[leafNo].score = formula2(inData1)
print(f"输入数据({leaf[leafNo].name}):",inData1)
print(leafNo,leaf[leafNo].score,end='\n\n')


# 5 对公手续费及佣金净收入 12
leafNo = 5
inData1 = np.array(data[26])
inData2 = np.array(data[27])
inData3 = np.array(data[28])
# inData1 = [55407.84 ,10512.25,20736.28,18369.57,-2836.72,-5106.09,3399.44,11204.49,20191.23,-9338.28,20638.48,-1380.51]
benchmark = 100
weight = 1
leaf[leafNo].score = (benchmark+(formula1(inData2)+formula1(inData3))/2) * weight / 100
print(f"输入数据({leaf[leafNo].name}):",inData1,sep='\n')
inData1 = np.array(data[22])
weight = 0.35
leaf[leafNo].score += 100* (inData1) * weight / 100
print(f"输入数据({leaf[leafNo].name}):",inData1,sep='\n')
print(leafNo,leaf[leafNo].score,end='\n\n')

# 6 个人手续费及佣金净收入 13
# 32	33	34
leafNo = 6
inData1 = np.array(data[32])
inData2 = np.array(data[33])
inData3 = np.array(data[34])
# inData1 = [55407.84 ,10512.25,20736.28,18369.57,-2836.72,-5106.09,3399.44,11204.49,20191.23,-9338.28,20638.48,-1380.51]
benchmark = 100
weight = 1
leaf[leafNo].score = (benchmark+(formula1(inData2)+formula1(inData3))/2) * weight / 100
print(f"输入数据({leaf[leafNo].name}):",inData1,sep='\n')
inData1 = np.array(data[22])
weight = 0.35
leaf[leafNo].score += 100* (inData1) * weight / 100
print(f"输入数据({leaf[leafNo].name}):",inData1,sep='\n')
print(leafNo,leaf[leafNo].score,end='\n\n')

# 7 手续费及佣金支出 14
leafNo = 7
inData1 = np.array(data[35])
inData2 = np.array(data[36])
leaf[leafNo].score = np.where(inData1 < inData2, -0.2, 0)
print(f"输入数据({leaf[leafNo].name}):", inData1, inData2, sep='\n')
print(leaf[leafNo].score, end='\n\n')

# 8 净收入预测偏离度 15
leafNo = 8
inData1 = np.array(data[37])
leaf[leafNo].score = np.array(inData1)
print(f"输入数据({leaf[leafNo].name}):",inData1,sep='\n')
print(leafNo,leaf[leafNo].score,end='\n\n')

# 9 4.2资金资管交易性收入 16
leafNo = 9
inData1 = np.array(data[40])
inData2 = np.array(data[41])
benchmark = 100
weight = 0.3
leaf[leafNo].score = (benchmark+(formula1(inData1)+formula1(inData2))/2) * weight / 100
print(f"输入数据({leaf[leafNo].name}):",inData1,inData2,sep='\n')
print(leafNo,leaf[leafNo].score,end='\n\n')

# 10 5.1一般性存款付息率 18
leafNo = 10
inData1 = np.array(data[42])
# leaf[leafNo].score = np.where(inData1 >= 1, 0, 0.1)
print(f"输入数据({leaf[leafNo].name}):", inData1, sep='\n')
print(leaf[leafNo].score, end='\n\n')

# 11 5.2贷款收益率 19
leafNo = 11
inData2 = np.array(data[44])
# leaf[leafNo].score = np.where(inData1 >= 1, 0, 0.1)
print(f"输入数据({leaf[leafNo].name}):", inData1, inData2, sep='\n')
print(leaf[leafNo].score, end='\n\n')

# %% 
# 发展转型
# 12 1.对公业务转型 20
# 13 2.个人业务转型 21
# 14 3.1母子公司资产业务联动 23
leafNo = 14
inData1 = np.array(data[210])
inData2 = np.array(data[211])
benchmark = 100
weight = 0.5
leaf[leafNo].score = (benchmark+(formula1(inData1)+formula1(inData2))/2) * weight / 100
print(f"输入数据({leaf[leafNo].name}):",inData1,inData2,sep='\n')
print(leafNo,leaf[leafNo].score,end='\n\n')

# 15 3.2.1代发工资金额 25
leafNo = 15
inData1 = np.array(data[214])
inData2 = np.array(data[215])
benchmark = 100
weight = 0.5
leaf[leafNo].score = (benchmark+(formula1(inData1)+formula1(inData2))/2) * weight / 100
print(f"输入数据({leaf[leafNo].name}):",inData1,inData2,sep='\n')
print(leafNo,leaf[leafNo].score,end='\n\n')

# 16 3.2.2代发个人客户AUM净增 26
leafNo = 16
inData1 = np.array(data[218])
inData2 = np.array(data[219])
benchmark = 100
weight = 0.5
leaf[leafNo].score = (benchmark+(formula1(inData1)+formula1(inData2))/2) * weight / 100
print(f"输入数据({leaf[leafNo].name}):",inData1,inData2,sep='\n')
print(leafNo,leaf[leafNo].score,end='\n\n')

# 17 3.3价值商户 27
leafNo = 17
inData1 = np.array(data[222])
inData2 = np.array(data[223])
benchmark = 100
weight = 0.5
leaf[leafNo].score = (benchmark+(formula1(inData1)*0.6+formula1(inData2)*0.4)) * weight / 100
print(f"输入数据({leaf[leafNo].name}):",inData1,inData2,sep='\n')
print(leafNo,leaf[leafNo].score,end='\n\n')

# 18 3.4.1养老金融规模 29
leafNo = 18
inData1 = np.array(data[226])
inData2 = np.array(data[227])
weight = 0.5
leaf[leafNo].score = ((formula1(inData1,weight)+formula1(inData2,weight))/2)
print(f"输入数据({leaf[leafNo].name}):",inData1,inData2,sep='\n')
print(leafNo,leaf[leafNo].score,end='\n\n')

# 19 3.4.2贵金属及大宗商品重点业务规模 30
leafNo = 19
inData1 = np.array(data[231])
inData2 = np.array(data[232])
# 计划完成率不超过500%
inData2 = np.where(inData2>5,5,inData2)
weight = 0.5
leaf[leafNo].score = ((formula1(inData1,weight)*0.8+formula1(inData2,weight)*0.2))
print(f"输入数据({leaf[leafNo].name}):",inData1,inData2,sep='\n')
print(leafNo,leaf[leafNo].score,end='\n\n')

# 20 3.4.3公募基金销售规模托管产品占比 31
leafNo = 20
inData1 = np.array(data[234])
leaf[leafNo].score = inData1
print(f"输入数据({leaf[leafNo].name}):",inData1,sep='\n')
print(leafNo,leaf[leafNo].score,end='\n\n')

# 21 3.4.4对客资金交易有效客户 32
leafNo = 21
inData1 = np.array(data[237])
inData2 = np.array(data[238])
weight = 0.5
leaf[leafNo].score = ((formula1(inData1,weight)*0.7+formula1(inData2,weight))*0.3)
print(f"输入数据({leaf[leafNo].name}):",inData1,inData2,sep='\n')
print(leafNo,leaf[leafNo].score,end='\n\n')

# 22 4.1重点区域行 34
leafNo = 22
leaf[leafNo].score = np.zeros_like(inData1,dtype='float')
print(' X ',leaf[leafNo].name,' 待定')

# 23 4.2叮咚系统异地联动 35
leafNo = 23
inData1 = np.array(data[242])
inData2 = np.array(data[243])
# 计算开户数得分
conditions1 = [
    inData1 >= 1,
    inData1 >= 0.75,
    inData1 >= 0.5,
    inData1 >= 0.25,
    True
]
choices1 = [0.15, 0.1, 0, -0.1, -0.15]
score1 = np.select(conditions1, choices1)

# 计算日均存款调入数得分
conditions2 = [
    inData2 >= 10000,
    inData2 >= 1000,
    True
]
choices2 = [0.05, 0, -0.05]
score2 = np.select(conditions2, choices2)
leaf[leafNo].score = score1 + score2
print(f"输入数据({leaf[leafNo].name}):", inData1, inData2, sep='\n')
print(leaf[leafNo].score, end='\n\n')

# 24 4.3科创股债联动业务 36
leafNo = 24
inData1 = np.array(data[248])
leaf[leafNo].score = inData1
print(f"输入数据({leaf[leafNo].name}):", inData1, sep='\n')
print(leaf[leafNo].score, end='\n\n')

# 25 4.4对公全量客户新增 37
# 不该出现基准分，但是该项基准分为100
leafNo = 25
inData1 = np.array(data[251])
inData2 = np.array(data[252])
# weight = 0.3
leaf[leafNo].score = ((formula1(inData1)+formula1(inData2))/2)/100
print(f"输入数据({leaf[leafNo].name}):",inData1,inData2,sep='\n')
print("# 不该出现基准分，但是该项基准分为100")
print(leafNo,leaf[leafNo].score,end='\n\n')

# 26 4.5对公房地产信贷增长 38
# 不该出现基准分，但是该项基准分为100
leafNo = 26
inData1 = np.array(data[255])
inData2 = np.array(data[256])

inData3 = np.array(data[257])
leaf[leafNo].score = (formula1(inData1)*0.7+formula1(inData2)*0.3)/100 + formula1(inData3,0.2)
print(f"输入数据({leaf[leafNo].name}):",inData1,inData2,sep='\n')
print("# 不该出现基准分，但是该项基准分为100")

print(leafNo,leaf[leafNo].score,end='\n\n')

# 27 4.6精细化管理 39
leafNo = 27
# inData1 = np.array(data[260])
leaf[leafNo].score = np.zeros_like(inData1,dtype='float')
print("# 精细化管理无数据")
print(' X ',leaf[leafNo].name,' 待定')

# 28 对公钱包 41
leafNo = 28
inData1 = np.array(data[261])
leaf[leafNo].score = inData1
print(f"输入数据({leaf[leafNo].name}):", inData1, sep='\n')
print(leaf[leafNo].score, end='\n\n')

# 29 个人钱包 42
leafNo = 29
inData1 = np.array(data[264])
leaf[leafNo].score = inData1
print(f"输入数据({leaf[leafNo].name}):", inData1, sep='\n')
print(leaf[leafNo].score, end='\n\n')

# 30 数币商户 43
leafNo = 30
inData1 = np.array(data[268])
leaf[leafNo].score = inData1
print(f"输入数据({leaf[leafNo].name}):", inData1, sep='\n')
print(leaf[leafNo].score, end='\n\n')

# 31 创新场景建设 44
leafNo = 31
inData1 = np.array(data[269])
leaf[leafNo].score = inData1
print(f"输入数据({leaf[leafNo].name}):", inData1, sep='\n')
print(leaf[leafNo].score, end='\n\n')


# %% 
# 风险合规
# 32 1.全面风险管理 45
leafNo = 32
inData1 = np.array(data[271])
weight = 25
leaf[leafNo].score = inData1 * weight / 100
print(f"输入数据({leaf[leafNo].name}):", inData1, sep='\n')
print(leaf[leafNo].score, end='\n\n')

# 33 2.1内控合规案防 47
leafNo = 33
inData1 = np.array(data[272])
weight = 12
leaf[leafNo].score = inData1 * weight / 100
leaf[leafNo].score = inData1
print(f"输入数据({leaf[leafNo].name}):", inData1, sep='\n')
print(leaf[leafNo].score, end='\n\n')

# 34 2.2涉外业务合规评价 48
leafNo = 34
inData1 = np.array(data[273])
weight = 1
leaf[leafNo].score = inData1 * weight / 100
leaf[leafNo].score = inData1
print(f"输入数据({leaf[leafNo].name}):", inData1, sep='\n')
print(leaf[leafNo].score, end='\n\n')

# 35 2.3数据治理 49
leafNo = 35
inData1 = np.array(data[274])
weight = 25
leaf[leafNo].score = inData1 * weight / 100
leaf[leafNo].score = inData1
print(f"输入数据({leaf[leafNo].name}):", inData1, sep='\n')
print(leaf[leafNo].score, end='\n\n')


# %%
# 社会责任
# 36 (1)贷款新增 53
leafNo = 36
inData1 = np.array(data[278])
# inData1 为负数时，得0分
weight = 1.5
benchmark = 100
leaf[leafNo].score = np.where(inData1<0,0,(benchmark+formula1(inData1)) * weight / 100)
print(f"输入数据({leaf[leafNo].name}):",inData1,sep='\n')
print(leafNo,leaf[leafNo].score,end='\n\n')

# 37 (2)贷款新增计划完成率 54
leafNo = 37
inData1 = np.array(data[279])
inData2 = np.array(data[280])
# inData1 为负数时，得0分
weight = 1.5
benchmark = 100
leaf[leafNo].score = np.where(inData1<0,0,(benchmark+formula1(inData1)) * weight / 100) * inData2
print(f"输入数据({leaf[leafNo].name}):",inData1,inData2,sep='\n')
print(leafNo,leaf[leafNo].score,end='\n\n')

# 38 a.客户新增 57
leafNo = 38
inData1 = np.array(data[285])
# inData1 为负数时，得0分
weight = 2.5
benchmark = 100
leaf[leafNo].score = np.where(inData1<0,0,(benchmark+formula1(inData1)) * weight / 100)
print(f"输入数据({leaf[leafNo].name}):",inData1,sep='\n')
print(leafNo,leaf[leafNo].score,end='\n\n')

# 39 b.客户新增完成率 58
leafNo = 39
inData1 = np.array(data[286])
# inData1 为负数时，得0分；限定区间为0~1.3
weight = 0.5
leaf[leafNo].score = np.where(inData1 < 0, 0, np.where(inData1 > 1.3, 1.3, inData1)) * weight
print(f"输入数据({leaf[leafNo].name}):", inData1, sep='\n')
print(leaf[leafNo].score, end='\n\n')

# 40 c.客户留存率 59
leafNo = 40
inData1 = np.array(data[290])
inData2 = np.array(data[291])
weight = 0.5
benchmark = 100
leaf[leafNo].score = (benchmark+formula1(inData1)*0.6+formula1(inData2)*0.4) * weight / 100
print(f"输入数据({leaf[leafNo].name}):",inData1,inData2,sep='\n')
print(leafNo,leaf[leafNo].score,end='\n\n')

# 41 (2)首贷户占比 60
# 完成率都是0，所以得分都为0
leafNo = 41
inData1 = np.array(data[292])
inData2 = np.array(data[293])
# inData1 限定区间为0~1,大于1得满分
weight = 0.5
benchmark = 100
leaf[leafNo].score = np.where(
    inData2 <= 0, 0, np.where(inData2 >= 1, 1.3, 1+formula1(inData1,0.3))) * weight
print(f"输入数据({leaf[leafNo].name}):", inData1, sep='\n')
print(leaf[leafNo].score, end='\n\n')


# 42 （1）小微企业信用贷款增长 62
leafNo = 42
inData1 = np.array(data[294])
if inData1.all() == 0:
    leaf[leafNo].score = 0
else:
    mean = np.mean(inData1)
    maxv = np.max(inData1)
    score = np.zeros(inData1.shape)
    score[inData1 < mean] = 0.25 * inData1[inData1 < mean] / mean
    score[inData1 > mean] = 0.25 + 0.25 * (inData1[inData1 > mean] - mean) / (maxv - mean)
    score[inData1 < 0] = -0.5
    leaf[leafNo].score = score
print(f"输入数据({leaf[leafNo].name}):",inData1,sep='\n')
print(leafNo,leaf[leafNo].score,end='\n\n')

# 43 （2）小微企业中长期贷款 63
leafNo = 43
inData1 = np.array(data[295])
inData2 = np.array(data[296])
if inData1.all() == 0:
    leaf[leafNo].score = np.zeros_like(inData1, dtype=float)
else:
    max_growth = np.max(inData1)
    score = np.zeros(inData1.shape)
    score[inData1 <= 0] = -0.25
    score[inData1 > 0] = 0.125 * inData1[inData1 >= 0] / max_growth + 0.125 * np.minimum(1, inData2[inData2 >= 0] / 0.005)
    leaf[leafNo].score = score
print(f"输入数据({leaf[leafNo].name}):",inData1,inData2,sep='\n')
print(leafNo,leaf[leafNo].score,end='\n\n')

# 44 （3）小微企业制造业贷款 64
leafNo = 44
inData1 = np.array(data[297])
inData2 = np.array(data[298])
if inData1.all() == 0:
    leaf[leafNo].score = np.zeros_like(inData1, dtype=float)
else:
    max_growth = np.max(inData1)
    score = np.zeros(inData1.shape)
    score[inData1 < 0] = -0.25
    score[inData1 >= 0] = 0.125 * inData1[inData1 >= 0] / max_growth + 0.125 * np.minimum(1, inData2[inData2 >= 0] / 0.005)
    leaf[leafNo].score = score
print(f"输入数据({leaf[leafNo].name}):",inData1,inData2,sep='\n')
print(leafNo,leaf[leafNo].score,end='\n\n')

# 45 （4）小微企业当年续贷累放 65
# 上年同期	当年累放	占比提升
leafNo = 45
inData1 = np.array(data[299])
inData2 = np.array(data[300])
inData3 = np.array(data[301])
score = np.zeros(inData1.shape)
score[(inData3 < 0) & (inData2 > inData1)] = 0.125
score[inData2 <= 0] = -0.25
score[inData3 >= 0] = 0.25
leaf[leafNo].score = score
print(f"输入数据({leaf[leafNo].name}):",inData1,inData2,inData3,sep='\n')
print(leafNo,leaf[leafNo].score,end='\n\n')


# 46 （5）小微企业银税互动贷款 66
leafNo = 46
inData1 = np.array(data[302])
inData2 = np.array(data[303])
leaf[leafNo].score = np.zeros_like(inData1,dtype=float)
if inData1.any() != 0:
    leaf[leafNo].score = np.where(inData1 < 0, -0.25, 0.125 * (inData1/inData1.max()))
if inData2.any() != 0:
    leaf[leafNo].score += np.where(inData1 < 0, 0, 0.125 * (inData2/inData2.max()))

print(f"输入数据({leaf[leafNo].name}):",inData1,inData2,sep='\n')
print(leafNo,leaf[leafNo].score,end='\n\n')

# 47 (1)对公涉农贷款 69
# 307	308	309
leafNo = 47
inData1 = np.array(data[307])
inData2 = np.array(data[308])
inData3 = np.array(data[309])
weight = 1

leaf[leafNo].score = np.where(inData1 < 0, 0,
                              1 + formula1(inData1,0.35) + formula1(inData2,0.35)  + inData3*0.3)
print(f"输入数据({leaf[leafNo].name}):",inData1,inData2,inData3,sep='\n')
print(leafNo,leaf[leafNo].score,end='\n\n')

# 48 (2)农户生产经营贷款 70
# 312	313
leafNo = 48
inData1 = np.array(data[312])
inData2 = np.array(data[313])
weight = 1
leaf[leafNo].score = 1 + (formula1(inData1,weight)*0.5+formula1(inData2,weight)*0.5)
leaf[leafNo].score = np.where(inData1 < 0, 0.2, leaf[leafNo].score)
print(f"输入数据({leaf[leafNo].name}):",inData1,inData2,sep='\n')
print(leafNo,leaf[leafNo].score,end='\n\n')

# 49 1.2.2农户生产经营贷款客户 71
# 316	317
leafNo = 49
inData1 = np.array(data[316])
inData2 = np.array(data[317])
weight = 1
leaf[leafNo].score = 1+ (formula1(inData1,weight)*0.5+formula1(inData2,weight)*0.5)
leaf[leafNo].score = np.where(inData1 < 0, 0.2, leaf[leafNo].score)
print(f"输入数据({leaf[leafNo].name}):",inData1,inData2,sep='\n')
print(leafNo,leaf[leafNo].score,end='\n\n')

# 50 (1)达标服务点数 73
# 仅通州分行有数据，故都赋基准分
leafNo = 50
inData1 = data[319]
weight = 0.5
benchmark = 100
leaf[leafNo].score = (benchmark+formula1(inData1)) * weight / 100
print(f"输入数据({leaf[leafNo].name}):",inData1,sep='\n')
print(leafNo,leaf[leafNo].score,end='\n\n')

# 51 (2)计划完成率 74
leafNo = 51
inData1 = np.array(data[320])
score = np.where(inData1<0.6,0,inData1)
weight = 0.5
leaf[leafNo].score = np.where(inData1 > 0.6, score * weight / 100, 0)
print(f"输入数据({leaf[leafNo].name}):",inData1,sep='\n')
print(leafNo,leaf[leafNo].score,end='\n\n')

# 52 1.2.4乡村振兴业务综合评价 75
leafNo = 52
inData1 = data[321]
leaf[leafNo].score = np.array(inData1)
print(f"输入数据({leaf[leafNo].name}):",inData1,sep='\n')
print(leafNo,leaf[leafNo].score,end='\n\n')

# 53 2.1.1制造业含贴贷款增长 78
leafNo = 53
inData1 = data[325]
inData2 = data[326]
inData3 = data[327]
benchmark = 100
weight = 0.8
leaf[leafNo].score = (benchmark+0.3*formula1(inData1)+0.2*formula1(inData2)+0.5*inData3) * weight / 100
print(f"输入数据({leaf[leafNo].name}):",inData1,inData2,inData3,sep='\n')
print(leafNo,leaf[leafNo].score,end='\n\n')

# 54 2.1.2制造业中长期贷款增长 79
leafNo = 54
inData1 = data[331]
inData2 = data[332]
inData3 = data[333]
benchmark = 100
weight = 1.2
leaf[leafNo].score = (benchmark+0.3*formula1(inData1)+0.2*formula1(inData2)+0.5*inData3) * weight / 100
print(f"输入数据({leaf[leafNo].name}):",inData1,inData2,inData3,sep='\n')
print(leafNo,leaf[leafNo].score,end='\n\n')

# 55 2.2民营企业支持 80
leafNo = 55
inData1 = data[336]
inData2 = data[337]
benchmark = 100
weight = 1
leaf[leafNo].score = (benchmark+(formula1(inData1)+formula1(inData2))/2) * weight / 100
print(f"输入数据({leaf[leafNo].name}):",inData1,inData2,sep='\n')
print(leafNo,leaf[leafNo].score,end='\n\n')

# 56 2.3民营企业客户 81
leafNo = 56
inData1 = np.array(data[340])
median = inData1.mean()
max_value = inData1.max()
min_value = inData1.min()
# divisor = np.zeros_like(inData1)
# score = np.where(inData1<0,-0.25,0)

leaf[leafNo].score = np.where(inData1 < 0, -0.25,  
                        np.where(inData1 < median,
                                (0.125 + 0.125 * (inData1 - median)/(median - min_value)),
                                (0.125 + 0.125 * (inData1 - median)/(max_value - median))))
print(f"输入数据({leaf[leafNo].name}):",inData1,sep='\n')
print(leafNo,leaf[leafNo].score,end='\n\n')


# 57 战略新兴贷款新增 83
# 未考虑加分情况
leafNo = 57
inData1 = data[343]
inData2 = data[344]
# inData3 = data[345]
benchmark = 100
weight = 1
leaf[leafNo].score = (benchmark+0.7*formula1(inData1)+0.3*formula1(inData2)) * weight / 100
print(f"输入数据({leaf[leafNo].name}):",inData1,inData2,inData3,sep='\n')
print(leafNo,leaf[leafNo].score,end='\n\n')

# 58 3.2.1转化客户数 85
leafNo = 58
inData1 = data[347]
leaf[leafNo].score = np.array(inData1)
print(f"输入数据({leaf[leafNo].name}):",inData1,sep='\n')
print(leafNo,leaf[leafNo].score,end='\n\n')


# 59 3.2.2推荐给建信住房并签约的项目数 86
leafNo = 59
inData1 = data[349]
leaf[leafNo].score = np.array(inData1)
print(f"输入数据({leaf[leafNo].name}):",inData1,sep='\n')
print(leafNo,leaf[leafNo].score,end='\n\n')

# 60 3.2.3保障性租赁住房支持 87
leafNo = 60
inData1 = data[351]
leaf[leafNo].score = np.array(inData1)
print(f"输入数据({leaf[leafNo].name}):",inData1,sep='\n')
print(leafNo,leaf[leafNo].score,end='\n\n')

# 61 3.2.4合规协同 88
leafNo = 61
inData1 = np.array(data[352])
leaf[leafNo].score = inData1
print(f"输入数据({leaf[leafNo].name}):",inData1,sep='\n')
print(leafNo,leaf[leafNo].score,end='\n\n')

# 62 4.1绿色贷款占比 91
leafNo = 62
inData1 = data[356]
leaf[leafNo].score = np.array(inData1)
print(f"输入数据({leaf[leafNo].name}):",inData1,sep='\n')
print(leafNo,leaf[leafNo].score,end='\n\n')


# 63 4.2绿色贷款余额增长贡献 92
leafNo = 63
inData1 = data[358]
leaf[leafNo].score = np.array(inData1)
print(f"输入数据({leaf[leafNo].name}):",inData1,sep='\n')
print(leafNo,leaf[leafNo].score,end='\n\n')

# 64 4.3绿色贷款新增计划执行 93
leafNo = 64
inData1 = data[360]
leaf[leafNo].score = np.array(inData1)
print(f"输入数据({leaf[leafNo].name}):",inData1,sep='\n')
print(leafNo,leaf[leafNo].score,end='\n\n')


# 65 4.4绿色贷款超额增长 94
leafNo = 65
inData1 = data[362]
leaf[leafNo].score = np.array(inData1)
print(f"输入数据({leaf[leafNo].name}):",inData1,sep='\n')
print(leafNo,leaf[leafNo].score,end='\n\n')

# 66 4.5绿色债券承销 95
leafNo = 65
leaf[leafNo].score = np.zeros_like(inData1,dtype=float)
print("提示：没有计划完成率数据")
# print(f"输入数据({leaf[leafNo].name}):",inData1,sep='\n')
print(leafNo,leaf[leafNo].score,end='\n\n')

# 67 5.1.1消保评价 98
leafNo = 67
inData1 = data[364]
leaf[leafNo].score = np.array(inData1)
print(f"输入数据({leaf[leafNo].name}):",inData1,sep='\n')
print(leafNo,leaf[leafNo].score,end='\n\n')

# 68 千百佳网点服务管理 100
leafNo = 68
inData1 = data[365]
leaf[leafNo].score = np.array(inData1)
print(f"输入数据({leaf[leafNo].name}):",inData1,sep='\n')
print(leafNo,leaf[leafNo].score,end='\n\n')

# 69 客户等候超长网点数量 101
leafNo = 69
inData1 = data[366]
leaf[leafNo].score = np.array(inData1)
print(f"输入数据({leaf[leafNo].name}):",inData1,sep='\n')
print(leafNo,leaf[leafNo].score,end='\n\n')

# 70 超长客户占比 102
leafNo = 70
inData1 = data[367]
leaf[leafNo].score = np.array(inData1)
print(f"输入数据({leaf[leafNo].name}):",inData1,sep='\n')
print(leafNo,leaf[leafNo].score,end='\n\n')

# 71 网点服务声誉风险 103
leafNo = 71
inData1 = data[368]
leaf[leafNo].score = -1*np.array(inData1)
print(f"输入数据({leaf[leafNo].name}):",inData1,sep='\n')
print(leafNo,leaf[leafNo].score,end='\n\n')

# 72 5.2信访工作 104
leafNo = 72
inData1 = data[370]
leaf[leafNo].score = -1*np.array(inData1)
print(f"输入数据({leaf[leafNo].name}):",inData1,sep='\n')
print(leafNo,leaf[leafNo].score,end='\n\n')


# %%
count = 0
for i in leaf1:
    print('#',count,i.name,i.index)
    count += 1


# %%
# 对公副卡root1
# 0 1.1对公加权有效客户 4
leaf1No = 0
inData1 = np.array(data[48])
inData2 = np.array(data[49])
benchmark = 100
weight = 2
leaf1[leaf1No].score = (benchmark + (formula1(inData1)+formula1(inData2))/2 )* weight / 100
print(f"输入数据({leaf1[leaf1No].name}):",inData1,inData2,sep='\n')
print(leaf1[leaf1No].score,end='\n\n')

# 1 1.2当年新开结算账户 5
# 计划完成率规则不明确，计算可能有误
leaf1No = 1
inData1 = np.array(data[53])
inData2 = np.array(data[54])
inData3 = np.array(data[55])
benchmark = 100
weight = 0.7
leaf1[leaf1No].score = (benchmark + (formula1(inData1)*0.6+formula1(inData2)*0.4))* weight / 100
leaf1[leaf1No].score += formula1(inData3,0.3)
print("X   计划完成率规则不明确，计算可能有误")
print(f"输入数据({leaf1[leaf1No].name}):",inData1,inData2,sep='\n')
print(leaf1[leaf1No].score,end='\n\n')

# 2 2.1对公一般性存款日均 7
# 增长额、增长率权重与计分规则冲突，以计分规则优先
leaf1No = 2
inData1 = np.array(data[59])
inData2 = np.array(data[60])
inData3 = np.array(data[61])
benchmark = 100
weight = 4
leaf1[leaf1No].score = (benchmark + (formula1(inData1,50)+formula1(inData2,50))/2 )* weight / 100
# 计划完成率得分计算
score = np.zeros(inData3.shape)
score[inData3 < 0] = 0
score[(0 <= inData3) & (inData3 <= 1.3)] = inData3[(0 <= inData3) & (inData3 <= 1.3)] * 100
score[(1.3 < inData3) & (inData3 <= 2)] = 130 + (inData3[(1.3 < inData3) & (inData3 <= 2)] - 1.3) * 20
score[inData3 > 2] = 145 + (inData3[inData3 > 2] - 2) * 2
leaf1[leaf1No].score += score/100
print(f"输入数据({leaf1[leaf1No].name}):",inData1,inData2,inData3,sep='\n')
print(leaf1[leaf1No].score,end='\n\n')

# 3 2.2对公集团全量资金 8
leaf1No = 3
inData1 = np.array(data[70])
inData2 = np.array(data[71])
benchmark = 100
weight = 0.5
leaf1[leaf1No].score = (benchmark + (formula1(inData1,50)+formula1(inData2,50))/2 )* weight / 100
print(f"输入数据({leaf1[leaf1No].name}):",inData1,inData2,sep='\n')
print(leaf1[leaf1No].score,end='\n\n')

# 4 2.3社会化平台存款沉淀 9
leaf1No = 4
inData1 = np.array(data[98])
inData2 = np.array(data[99])
inData2 = np.where(inData2>1.5,1.5,inData2)
benchmark = 100
weight = 0.5
leaf1[leaf1No].score = (benchmark + (formula1(inData1)*0.6+formula1(inData2,50)*0.4) )* weight / 100
print(f"输入数据({leaf1[leaf1No].name}):",inData1,inData2,sep='\n')
print(leaf1[leaf1No].score,end='\n\n')

# 5 2.4同业活期存款 10
leaf1No = 5
inData1 = np.array(data[101])
inData2 = np.array(data[102])
inData3 = np.array(data[103])
leaf1[leaf1No].score = formula1(inData1,0.25)*0.3 + formula1(inData2,0.25)*0.35+formula1(inData3,0.25)*0.35
print(f"输入数据({leaf1[leaf1No].name}):",inData1,inData2,inData3,sep='\n')
print(leaf1[leaf1No].score,end='\n\n')

# 6 3.1.1现金管理客户净增 13
# 加减分与计分区间冲突，以加减分为准
leaf1No = 6
inData1 = np.array(data[106])
inData2 = np.array(data[107])
weight = 0.3
benchmark = 100
leaf1[leaf1No].score = (benchmark + (formula1(inData1,50)*0.5+formula1(inData2,50)*0.5) )* weight / 100
print(f"输入数据({leaf1[leaf1No].name}):",inData1,inData2,sep='\n')
print(leaf1[leaf1No].score,end='\n\n')

# 7 3.1.2债券承销业务大中型客户增长 14
# 加减分与计分区间冲突，以加减分为准
leaf1No = 7
inData1 = np.array(data[108])
weight = 0.1
benchmark = 100
leaf1[leaf1No].score = (benchmark + formula1(inData1,50) )* weight / 100
print(f"输入数据({leaf1[leaf1No].name}):",inData1,sep='\n')
print(leaf1[leaf1No].score,end='\n\n')

# 8 3.1.3财务顾问及投行平台签约客户 15
# 考核方式与计分规则冲突，以计分规则为准

leaf1No = 8
inData1 = np.array(data[110])
weight = 0.1
benchmark = 100
leaf1[leaf1No].score = (benchmark + formula1(inData1) )* weight / 100
print(f"输入数据({leaf1[leaf1No].name}):",inData1,sep='\n')
print(leaf1[leaf1No].score,end='\n\n')

# 9 3.2.1国际业务金融总量 17
# 计分规则与计分区间冲突，以计分规则为准
leaf1No = 9
inData1 = np.array(data[113])
weight = 0.3
leaf1[leaf1No].score = inData1 * weight
print(f"输入数据({leaf1[leaf1No].name}):",inData1,sep='\n')
print(leaf1[leaf1No].score,end='\n\n')

# 10 3.2.2跨境人民币结算量 18
# 计分规则与计分区间冲突，以计分规则为准
leaf1No = 10
inData1 = np.array(data[116])
weight = 0.2
leaf1[leaf1No].score = inData1 * weight
print(f"输入数据({leaf1[leaf1No].name}):",inData1,sep='\n')
print(leaf1[leaf1No].score,end='\n\n')

# 11 3.3科技信贷 19
leaf1No = 11
inData1 = np.array(data[122])
inData2 = np.array(data[123])
weight = 0.5
benchmark = 100
leaf1[leaf1No].score = (benchmark + (formula1(inData1,50)*0.5+formula1(inData2,50)*0.5) )* weight / 100
print(f"输入数据({leaf1[leaf1No].name}):",inData1,inData2,sep='\n')
print(leaf1[leaf1No].score,end='\n\n')

# 12 3.4中型客户增长 20
leaf1No = 12
inData1 = np.array(data[127])
inData2 = np.array(data[128])
weight = 0.25
benchmark = 100
leaf1[leaf1No].score = (benchmark + (formula1(inData1) * 0.5 + formula1(inData2,50)*0.5)) * weight / 100
# leaf1[leaf1No].score += inData2 * weight
print(f"输入数据({leaf1[leaf1No].name}):",inData1,inData2,sep='\n')
print(leaf1[leaf1No].score,end='\n\n')

# 13 3.5基础设施贷款新增 21
leaf1No = 13
inData1 = np.array(data[131])
inData2 = np.array(data[132])
weight = 0.5
benchmark = 100
leaf1[leaf1No].score = (benchmark + (formula1(inData1)*0.5+formula1(inData2)*0.5) )* weight / 100
print(f"输入数据({leaf1[leaf1No].name}):",inData1,inData2,sep='\n')
print(leaf1[leaf1No].score,end='\n\n')

# 14 3.6.1供应链“建行e贷”融资总额新增 23
leaf1No = 14
inData1 = np.array(data[135])
inData2 = np.array(data[136])
weight = 0.25
benchmark = 100
leaf1[leaf1No].score = (benchmark + (formula1(inData1)*0.7+formula1(inData2)*0.3) )* weight / 100
print(f"输入数据({leaf1[leaf1No].name}):",inData1,inData2,sep='\n')
print(leaf1[leaf1No].score,end='\n\n')

# 15 3.6.2供应链“建行e贷”普惠贷款余额新增 24
leaf1No = 15
inData1 = np.array(data[139])
inData2 = np.array(data[140])
weight = 0.25
benchmark = 100
leaf1[leaf1No].score = (benchmark + (formula1(inData1)*0.7+formula1(inData2)*0.3) )* weight / 100
print(f"输入数据({leaf1[leaf1No].name}):",inData1,inData2,sep='\n')
print(leaf1[leaf1No].score,end='\n\n')


# %%
count = 0
for i in leaf2:
    print('#',count,i.name,i.index)
    count += 1

# %%
# 个人副卡root2
# 0 1.1个人加权有效客户增长 4
# 144	145	146
leaf2No = 0
inData1 = np.array(data[144])
inData2 = np.array(data[145])
inData3 = np.array(data[146])
weight = 1
benchmark = 100
leaf2[leaf2No].score = (benchmark + (formula1(inData1)*0.5+formula1(inData2)*0.5) )* weight / 100
leaf2[leaf2No].score += inData3 * weight
print(f"输入数据({leaf2[leaf2No].name}):",inData1,inData2,inData3,sep='\n')
print(leaf2[leaf2No].score,end='\n\n')

# 1 1.2加权财富管理客户新增 5
leaf2No = 1
inData1 = np.array(data[149])
inData2 = np.array(data[150])
weight = 1
benchmark = 100
leaf2[leaf2No].score = (benchmark + (formula1(inData1)*0.5+formula1(inData2)*0.5) )* weight / 100
print(f"输入数据({leaf2[leaf2No].name}):",inData1,inData2,sep='\n')
print(leaf2[leaf2No].score,end='\n\n')

# 2 2.1个人存款日均新增 7
leaf2No = 2
inData1 = np.array(data[154])
inData2 = np.array(data[155])
inData3 = np.array(data[156])
weight = 3
benchmark = 100
leaf2[leaf2No].score = (benchmark + (formula1(inData1,50)*0.6+formula1(inData2,50)*0.4) )* weight / 100
score = np.zeros(inData3.shape)
score[inData3 < 0] = 0
score[(0 <= inData3) & (inData3 <= 1.3)] = inData3[(0 <= inData3) & (inData3 <= 1.3)] * 100
score[(1.3 < inData3) & (inData3 <= 2)] = 130 + (inData3[(1.3 < inData3) & (inData3 <= 2)] - 1.3) * 20
score[inData3 > 2] = 145 + (inData3[inData3 > 2] - 2) * 2
leaf2[leaf2No].score += score/100
print(f"输入数据({leaf2[leaf2No].name}):",inData1,inData2,sep='\n')
print(leaf2[leaf2No].score,end='\n\n')

# 3 2.2个人全量资金日均新增 8
leaf2No = 3
inData1 = np.array(data[159])
inData2 = np.array(data[160])
weight = 2
benchmark = 100
leaf2[leaf2No].score = (benchmark + (formula1(inData1,50)*0.6+formula1(inData2,50)*0.4) )* weight / 100
print(f"输入数据({leaf2[leaf2No].name}):",inData1,inData2,sep='\n')
print(leaf2[leaf2No].score,end='\n\n')

# 4 3.1.1手机银行月活晋升达标客户数 11
leaf2No = 4
inData1 = np.array(data[169])
weight = 0.25
benchmark = 100
leaf2[leaf2No].score = (benchmark + formula1(inData1) )* weight / 100
print(f"输入数据({leaf2[leaf2No].name}):",inData1,sep='\n')
print(leaf2[leaf2No].score,end='\n\n')

# 5 3.1.2手机银行月活晋升达标客户增长率 12
leaf2No = 5
inData1 = np.array(data[170])
weight = 0.25
benchmark = 100
leaf2[leaf2No].score = (benchmark + formula1(inData1) )* weight / 100
print(f"输入数据({leaf2[leaf2No].name}):",inData1,sep='\n')
print(leaf2[leaf2No].score,end='\n\n')

# 6 3.2.1建行生活权益客户数 14
leaf2No = 6
inData1 = np.array(data[171])
weight = 0.25
benchmark = 100
leaf2[leaf2No].score = (benchmark + formula1(inData1) )* weight / 100
print(f"输入数据({leaf2[leaf2No].name}):",inData1,sep='\n')
print(leaf2[leaf2No].score,end='\n\n')

# 7 3.2.2建行生活权益客户晋升率 15
leaf2No = 7
inData1 = np.array(data[172])
weight = 0.25
benchmark = 100
leaf2[leaf2No].score = (benchmark + formula1(inData1) )* weight / 100
print(f"输入数据({leaf2[leaf2No].name}):",inData1,sep='\n')
print(leaf2[leaf2No].score,end='\n\n')

# 8 3.2.3建行生活政府消费券项目承接加分项 16
leaf2No = 8
inData1 = np.array(data[173])
leaf2[leaf2No].score = inData1
print(f"输入数据({leaf2[leaf2No].name}):",inData1,sep='\n')
print(leaf2[leaf2No].score,end='\n\n')

# 9 3.3社保卡和养老金账户新增 17
leaf2No = 9
inData1 = np.array(data[176])
weight = 0.5
benchmark = 100
leaf2[leaf2No].score = (benchmark + formula1(inData1) )* weight / 100
print(f"输入数据({leaf2[leaf2No].name}):",inData1,sep='\n')
print(leaf2[leaf2No].score,end='\n\n')

# 10 信用卡分期收入 19
leaf2No = 10
inData1 = np.array(data[180])
inData2 = np.array(data[181])
inData3 = np.array(data[182])
weight = 0.4
benchmark = 100
leaf2[leaf2No].score = (benchmark + (formula1(inData1)*0.6+formula1(inData2)*0.4) )* weight / 100
weight = 0.1
leaf2[leaf2No].score += inData3 * 0.1
print(f"输入数据({leaf2[leaf2No].name}):",inData1,inData2,inData3,sep='\n')
print(leaf2[leaf2No].score,end='\n\n')

# 11 3.5.1信用卡消费交易额（含分期） 21
leaf2No = 11
inData1 = np.array(data[186])
inData2 = np.array(data[187])
inData3 = np.array(data[188])
weight = 0.20
benchmark = 100
leaf2[leaf2No].score = (benchmark + (formula1(inData1)*0.6+formula1(inData2)*0.4) )* weight / 100
weight = 0.05
leaf2[leaf2No].score += inData3 * 0.1
print(f"输入数据({leaf2[leaf2No].name}):",inData1,inData2,inData3,sep='\n')
print(leaf2[leaf2No].score,end='\n\n')

# 12 3.5.2信用卡贷款余额 22
leaf2No = 12
inData1 = np.array(data[192])
inData2 = np.array(data[193])
inData3 = np.array(data[194])
weight = 0.225
benchmark = 100
leaf2[leaf2No].score = (benchmark + (formula1(inData1)*0.6+formula1(inData2)*0.4) )* weight / 100
weight = 0.025
leaf2[leaf2No].score += inData3 * 0.1
print(f"输入数据({leaf2[leaf2No].name}):",inData1,inData2,inData3,sep='\n')
print(leaf2[leaf2No].score,end='\n\n')

# 13 3.6私人银行客户金融资产净增 23
leaf2No = 13
inData1 = np.array(data[197])
weight = 0.5
leaf2[leaf2No].score = inData1 * weight
print(f"输入数据({leaf2[leaf2No].name}):",inData1,sep='\n')
print(leaf2[leaf2No].score,end='\n\n')

# 14 个人消费贷款新增 26
leaf2No = 14
inData1 = np.array(data[201])
inData2 = np.array(data[202])
leaf2[leaf2No].score = (formula1(inData1,0.3)+formula1(inData2,0.2))
print(f"输入数据({leaf2[leaf2No].name}):",inData1,inData2,sep='\n')
print(leaf2[leaf2No].score,end='\n\n')

# 15 3.7.2个人住房贷款新增 27
leaf2No = 15
inData1 = np.array(data[207])
leaf2[leaf2No].score = formula1(inData1,0.5)
print(f"输入数据({leaf2[leaf2No].name}):",inData1,sep='\n')
print(leaf2[leaf2No].score,end='\n\n')


# %%
# bugfix: “债转股”不在“3.2.4合规协同”
# 将其移动到“3.服务战略新兴产业”下
# 此后求取leaf会发生变动，导致出错，所以此后不要再求leaf
# 该节点单独计分
new_parent = root.search_by_name('3.服务战略新兴产业')
old_parent = root.search_by_name('3.2.4合规协同')
child = old_parent.search_by_name('债转股')
old_parent.children = []
new_parent.children.append(child)
inData1 = np.array(data[353])
child.score = 0.25 * inData1 / inData1.max()
print(f"输入数据({leaf2[leaf2No].name}):",inData1,sep='\n')
print(leaf2[leaf2No].score,end='\n\n')

# %%

# 将root1、root2插入root
a = root.search_by_name('1.对公业务转型')
# root1
a.children = root1.children

b = root.search_by_name('2.个人业务转型')
# root2
b.children = root2.children

# %%
# 递归，自底向上根据叶子节点的得分计算父节点的得分
def cal_score(parent):
    if parent.children == []:
        return
    else:
        parent.score *= 0
        for child in parent.children:
            cal_score(child)
            try:
                parent.score += child.score
            except:
                print(child.score.dtype,parent.score.dtype)
                print(child.name,parent.name)
            # parent.score += child.score
cal_score(root)

# %%

# 前序遍历，输出树结构，存入dataframe
#columns = ['name','score','rank','score','rank','score','rank'......]
columns = ['指标名称']
for i in PARTMENTS_A:
    columns.append(i)
    columns.append('排名')
print(columns)


out = xw.Book()
sheet = out.sheets.add()
depth = 1
ranklist = []

sheet.range('1:1').value = columns
# 深度优先搜索输出
stack = [root]
while stack!=[]:
    depth +=1
    node = stack.pop()
    # print(depth,node.name)
    ranklist = - np.argsort(node.score).argsort() + len(node.score)

    if node.score.all() == 0:
        oneRow = [node.name] + [np.nan]*len(PARTMENTS_A)*2
    else:
        oneRow = [node.name] + np.column_stack((node.score,ranklist)).flatten().tolist()
    if node.name in A:
        depth += 1

    # 插入名字，【（分数，排名），（分数，排名）】
    print(depth,oneRow)

    sheet.range(f'{depth}:{depth}').value = list(oneRow)
    stack += node.children[::-1]
    # print(depth,stack)


out.save(OUTPUT_PATH)
out.close()
dataxw.close()
rulebook.close()

# %%


# %%