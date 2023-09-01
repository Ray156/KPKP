# %%
from xlwings import App,Book
# import xlwings as xw
from difflib import get_close_matches as gcm
from os import makedirs
from fuzzywuzzy import fuzz as fz
# fz.partial_ratio('丰岳支行','丰科园支行')
# %%
# 预先设置好所有机构的名称，用于后续的机构名称的统一
PARTMENT = ["营业部","丰岳支行","丰科园支行","樊羊路支行","郭公庄支行",\
            "长辛店支行","云岗支行","怡海花园支行","五里店支行","梅市口支行",\
            "长安新城支行","橙色年代支行","晓月苑支行","韩庄子支行","马官营支行",\
            "业务管理部","驻行纪检组","住房金融业务部","信用卡业务部","国际业务部","综合管理部",
            "机构业务部","公司业务部","集团客户部","普惠金融事业部","个人金融部","清机团队",]

# 单个sheet最大行数
MAX_ROW = 120
# 机构名称表头的关键词，用于匹配机构名称
TITLE_GROUP=['机构名称','网点','部门','团队名称']
# 源文件路径
SOURCE_PATH = "data/task1.xlsx"
# 输出文件路径
OUTPUT_PATH = "各部门绩效/"
# %%

app = App(visible=False, add_book=True)
app.display_alerts = False
makedirs(OUTPUT_PATH,exist_ok=True)
# app.screen_updating = True  # 是否实时刷新excel程序的显示内容
# %%

for p in PARTMENT:
    print("-正在处理：",p,".xlsx----------------")
    inbook = app.books.open(SOURCE_PATH)
    
    # 逐个sheet进行处理
    for sheetNo in range(len(inbook.sheets)):
        sheet = inbook.sheets[sheetNo]
        row1 = 1 # 表头信息所在行
 
        cell_value = None

        while row1<10:
            cell_value = sheet.range('A'+str(row1)).value

            if cell_value != None:

                matchword = gcm(cell_value,TITLE_GROUP,1,cutoff=0.7)
                if matchword != []:
                    while sheet.range(f'A{row1+1}').value == None:
                        row1 += 1
                    break
            row1 +=1

        str_list = [None] * (row1 - 1) + sheet.range(f'A{row1}',f'A{MAX_ROW}').value

        flag = True
        remain = []
        for i,j in enumerate(str_list):
            if j == None:
                if not flag and not sheet.range(i+1,i+1).value == None:

                    remain.append(i+1)
                continue

            match_score = fz.partial_ratio(p,j)
            if match_score > 80:
                remain.append(i+1)
                flag = False
            else:
                flag = True
                continue

            # matchword = gcm(p,[j],1,cutoff=0.65)
            # if matchword != []:
            #     remain.append(i+1)
            #     flag = False
            # else:
            #     flag = True
            #     continue

        lastk = MAX_ROW + 1
        remain = remain[::-1] +[row1]
        for k in remain:
            if k == lastk - 1:
                lastk = k
                continue
            sheet.range(f'A{k+1}:A{lastk-1}').api.EntireRow.Delete()
            lastk = k
        # if row1+1 == remain[0]:
        #     pass
        # else
        #     sheet.range(f'A{row1+1}:A{row2-1}').api.EntireRow.Delete()
        # 仅保留0~row1行和row2~row3行，若row2为None，则保留0~row1行, 其余行删除
        print(f"{sheetNo:<4}{p:　<10}{sheet.name:　<25}\t{remain}")

    inbook.save(f'{OUTPUT_PATH}{p}.xlsx')
    inbook.close()

# %%
