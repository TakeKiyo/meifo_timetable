#入力する部分は、ファイル名、行数、仕事数の割り振り、サポート入れる人
import pandas as pd
from numpy.random import randint

#ファイル名入力、読み込み
file_name = "worktable_template.xlsx"
table = pd.read_excel(file_name)

#n:エクセルの行数
n = 20
#欠損値の補完
table = table.fillna('-')

#m:照明とかの仕事の数の割り振り
m1 = 5
m2 = 10
m2_2 = 11
m3 = 15
m4 = 20

#サポート入れる人
s_drummer = [""]
s_bassist = [""]
s_keyboardist = [""]
s_guitarist =[""]
supporter = s_drummer + s_bassist + s_keyboardist + s_guitarist


#ドラマーのサポート
num_s_dr = len(s_drummer)
#まずランダムに割りあてる
for i in range(n):
    table.at[i,'support_dr'] = s_drummer[randint(num_s_dr)]
    # 自分のサポートになっていたらやり直す
    while table.at[i, 'Dr.'] == table.at[i, 'support_dr']:
        table.at[i, 'support_dr'] = s_drummer[randint(num_s_dr)]
#自分の出番の次にサポートに入る
for j in range(n - 1):
    if table.at[j, 'Dr.'] in s_drummer:
        table.at[j+1, 'support_dr'] = table.at[j, 'Dr.']
#-などにはサポートを割り振らない
for k in range(n):
    if table.at[k, 'Dr.'] == "-":
        table.at[k, 'support_dr'] = "-"

#ベースのサポート
num_s_ba = len(s_bassist)
#まずランダムに割りあてる
for i in range(n):
    table.at[i,'support_ba'] = s_bassist[randint(num_s_ba)]
#サポート連続にならないようにする
for h in range(1,n):
    while table.at[h,'support_ba'] == table.loc[h - 1,'support_ba']:
        table.at[h,'support_ba'] = s_bassist[randint(num_s_ba)]
# 演奏の後にサポートにならないようにする
for k in range(n - 1):
    while table.at[k, 'Ba.'] == table.at[k + 1, 'support_ba']:
        table.at[k + 1, 'support_ba'] = s_bassist[randint(num_s_ba)]
#演奏の前にサポートにならないようにする
for l in range(1,n):
    while table.at[l, 'Ba.'] == table.at[l - 1, 'support_ba']:
        table.at[l - 1, 'support_ba'] = s_bassist[randint(num_s_ba)]
#自分のサポートをしないようにする、-のところは-
for j in range(n):
    while table.at[j, 'Ba.'] == table.at[j, 'support_ba']:
        table.at[j, 'support_ba'] = s_bassist[randint(num_s_ba)]
    if table.at[j, 'Ba.'] == "-":
        table.at[j, 'support_ba'] = "-"

#キーボードのサポート
num_s_key = len(s_keyboardist)
#まずランダムに割りあてる
for i in range(n):
    table.at[i,'support_key'] = s_keyboardist[randint(num_s_key)]
#サポート連続にならないようにする
for h in range(1,n):
    while table.at[h,'support_key'] == table.loc[h - 1,'support_key']:
        table.at[h,'support_key'] = s_keyboardist[randint(num_s_key)]
#演奏の後にサポートにならないようにする
for k in range(n - 1):
    while table.at[k, 'Key.'] == table.at[k + 1, 'support_key']:
        table.at[k + 1, 'support_key'] = s_keyboardist[randint(num_s_key)]
#演奏の前にサポートにならないようにする
for l in range(1,n):
    while table.at[l, 'Key.'] == table.at[l - 1, 'support_key']:
        table.at[l - 1, 'support_key'] = s_keyboardist[randint(num_s_key)]
#自分のサポートをしないようにする、-のところは-
for j in range(n):
    while table.at[j, 'Key.'] == table.at[j, 'support_key']:
        table.at[j, 'support_key'] = s_keyboardist[randint(num_s_key)]
    if table.at[j, 'Key.'] == "-":
        table.at[j, 'support_key'] = "-"

#ギター
num_s_gt = len(s_guitarist)

#ギター1


#まずランダムに割りあてる
for i in range(n):
    table.at[i,'support_gt1'] = s_guitarist[randint(num_s_gt)]
#サポート連続にならないようにする
for h in range(1,n):
    while table.at[h,'support_gt1'] == table.loc[h - 1,'support_gt1']:
        table.at[h,'support_gt1'] = s_guitarist[randint(num_s_gt)]
#演奏の後にサポートにならないようにする
for k in range(n - 1):
    while table.at[k, 'Gt.1'] == table.at[k + 1, 'support_gt1']:
        table.at[k + 1, 'support_gt1'] = s_guitarist[randint(num_s_gt)]
#演奏の前にサポートにならないようにする
for l in range(1,n - 1):
    while table.at[l, 'Gt.1'] == table.at[l - 1, 'support_gt1']:
        table.at[l - 1, 'support_gt1'] = s_guitarist[randint(num_s_gt)]
#自分のサポートなし、ギター2とのかぶりなし、-のところは-
for j in range(n):
    while table.at[j, 'Gt.1'] == table.at[j, 'support_gt1'] or table.at[j, 'support_gt1'] == table.at[j, 'Gt.2']:
        table.at[j, 'support_gt1'] = s_guitarist[randint(9)]
    if table.at[j, 'Gt.1'] == "-":
        table.at[j, 'support_gt1'] = "-"


#ギター2

#まずランダムに割りあてる
for i in range(n):
   table.at[i,'support_gt2'] = s_guitarist[randint(num_s_gt)]
#サポート連続にならないようにする
for h in range(1,n):
    while table.at[h,'support_gt2'] == table.loc[h - 1,'support_gt2']:
        table.at[h,'support_gt2'] = s_guitarist[randint(num_s_gt)]
#演奏の次にサポートにならないようにする
for k in range(n - 1):
    while table.at[k, 'Gt.2'] == table.at[k + 1, 'support_gt2']:
        table.at[k + 1, 'support_gt2'] = s_guitarist[randint(num_s_gt)]
#演奏の前にサポートにならないようにする
for l in range(1,n - 1):
    while table.at[l, 'Gt.2'] == table.at[l - 1, 'support_gt2']:
        table.at[l - 1, 'support_gt2'] = s_guitarist[randint(num_s_gt)]
#自分のサポートなし、ギター１とのかぶりなし、ギター1のサポートとのかぶりなし、-のところは-
# 自分のサポートなし、ギター１とのかぶりなし、ギター1のサポートとのかぶりなし、-のところは-
for j in range(n):
    while table.at[j, 'Gt.2'] == table.at[j, 'support_gt2'] or table.at[j, 'Gt.1'] == table.at[j, 'support_gt2'] or table.at[j, 'support_gt1'] == table.at[j, 'support_gt2']:
        table.at[j, 'support_gt2'] = s_guitarist[randint(num_s_gt)]
    if table.at[j, 'Gt.2'] == "-":
        table.at[j, 'support_gt2'] = "-"


#Media、Cam&Time、照明

#0〜m1-1
m1_performer = []
m1_supporter = []
# 演奏する人を抽出
for i in range(m1):
    m1_performer.append(table.at[i,'Vo.'])
    m1_performer.append(table.at[i, 'Gt.1'])
    m1_performer.append(table.at[i, 'Gt.2'])
    m1_performer.append(table.at[i, 'Ba.'])
    m1_performer.append(table.at[i, 'Dr.'])
    m1_performer.append(table.at[i, 'Key.'])
    m1_performer.append(table.at[i, 'その他'])
#-を除外
while "-" in m1_performer:
    m1_performer.remove("-")
print(m1_performer)
#演奏しない人でサポートできる人を抽出
for i in range(len(supporter)):
    if supporter[i] not in m1_performer:
        m1_supporter.append(supporter[i])
print(m1_supporter)
num_m1_sup = len(m1_supporter)
#mediaを割り当て
table.ix[0:m1 - 1,'Media'] = m1_supporter[randint(num_m1_sup)]
#ビデオにかぶらないようにカメラ割り当て
table.ix[0:m1 - 1,'Cam&Time'] = m1_supporter[randint(num_m1_sup)]
while table.ix[0,'Cam&Time'] == table.ix[0,'Media']:
    table.ix[0:m1 - 1, 'Cam&Time'] = m1_supporter[randint(num_m1_sup)]
#ビデオ、カメラにかぶらないように割り当て
table.ix[0:m1 - 1,'照明'] = m1_supporter[randint(num_m1_sup)]
while table.ix[0,'照明'] == table.ix[0,'Media'] or table.ix[0,'照明'] == table.ix[0,'Cam&Time']:
    table.ix[0:m1 - 1, '照明'] = m1_supporter[randint(num_m1_sup)]

#m1〜m2-1
m2_performer = []
m2_supporter = []
#演奏する人を抽出
for i in range(m1,m2):
    m2_performer.append(table.at[i,'Vo.'])
    m2_performer.append(table.at[i, 'Gt.1'])
    m2_performer.append(table.at[i, 'Gt.2'])
    m2_performer.append(table.at[i, 'Ba.'])
    m2_performer.append(table.at[i, 'Dr.'])
    m2_performer.append(table.at[i, 'Key.'])
    m2_performer.append(table.at[i, 'その他'])
while "-" in m2_performer:
    m2_performer.remove("-")
print(m2_performer)
for i in range(len(supporter)):
    if supporter[i] not in m2_performer:
        m2_supporter.append(supporter[i])
print(m2_supporter) #仕事入れる人
num_m2_sup = len(m2_supporter)
#Mediaを連続にならないように割り当て
table.ix[m1:m2 - 1,'Media'] = m2_supporter[randint(num_m2_sup)]
while table.ix[m1,'Media'] == table.ix[0,'Media']:
    table.ix[m1:m2 - 1, 'Media'] = m2_supporter[randint(num_m2_sup)]
#カメラを決める、連続かビデオとかぶってたらもう1回
table.ix[m1:m2 - 1,'Cam&Time'] = m2_supporter[randint(num_m2_sup)]
while table.ix[m1,'Cam&Time'] == table.ix[m1,'Media'] or table.ix[m1,'Cam&Time'] == table.ix[0,'Cam&Time']:
    table.ix[m1:m2 - 1, 'Cam&Time'] = m2_supporter[randint(num_m2_sup)]
#照明を決める、連続か、ビデオ、カメラとかぶってたらもう1回
table.ix[m1:m2 - 1,'照明'] = m2_supporter[randint(num_m2_sup)]
while table.ix[m1,'照明'] == table.ix[m1,'Media'] or table.ix[m1,'照明'] == table.ix[m1,'Cam&Time'] or table.ix[m1,'照明'] == table.ix[0,'照明']:
    table.ix[m1:m2 - 1, '照明'] = m2_supporter[randint(num_m2_sup)]


#m2_2〜m3-1
m3_performer = []
m3_supporter = []
for i in range(m2_2,m3):
    m3_performer.append(table.at[i,'Vo.'])
    m3_performer.append(table.at[i, 'Gt.1'])
    m3_performer.append(table.at[i, 'Gt.2'])
    m3_performer.append(table.at[i, 'Ba.'])
    m3_performer.append(table.at[i, 'Dr.'])
    m3_performer.append(table.at[i, 'Key.'])
    m3_performer.append(table.at[i, 'その他'])
while "なし" in m3_performer:
    m3_performer.remove("なし")
print(m3_performer)
for i in range(len(supporter)):
    if supporter[i] not in m3_performer:
        m3_supporter.append(supporter[i])
print(m3_supporter) #仕事入れる人
num_m3_sup = len(m3_supporter)
#Mediaを連続しないように割り当て
table.ix[m2_2:m3 - 1,'Media'] = m3_supporter[randint(num_m3_sup)]
while table.ix[m2_2,'Media'] == table.ix[m1,'Media']:
    table.ix[m2_2:m3 - 1, 'Media'] = m3_supporter[randint(num_m3_sup)]
#カメラを決める、ビデオとかぶってたらもう1回
table.ix[m2_2:m3 - 1,'Cam&Time'] = m3_supporter[randint(num_m3_sup)]
while table.ix[m2_2,'Cam&Time'] == table.ix[m2_2,'Media'] or table.ix[m2_2,'Cam&Time'] == table.ix[m1,'Cam&Time']:
    table.ix[m2_2:m3 - 1, 'Cam&Time'] = m3_supporter[randint(num_m3_sup)]
#照明を決める、ビデオ、カメラとかぶってたらもう1回
table.ix[m2_2:m3 - 1,'照明'] = m3_supporter[randint(num_m3_sup)]
while table.ix[m2_2,'照明'] == table.ix[m2_2,'Media'] or table.ix[m2_2,'照明'] == table.ix[m2_2,'Cam&Time'] or table.ix[m2_2,'照明'] == table.ix[m1,'照明']:
    table.ix[m2_2:m3 - 1, '照明'] = m3_supporter[randint(num_m3_sup)]

m4_performer = []
m4_supporter = []
for i in range(m3,m4):
    m4_performer.append(table.at[i,'Vo.'])
    m4_performer.append(table.at[i, 'Gt.1'])
    m4_performer.append(table.at[i, 'Gt.2'])
    m4_performer.append(table.at[i, 'Ba.'])
    m4_performer.append(table.at[i, 'Dr.'])
    m4_performer.append(table.at[i, 'Key.'])
    m4_performer.append(table.at[i, 'その他'])
while "なし" in m4_performer:
    m4_performer.remove("なし")
print(m4_performer)
for i in range(len(supporter)):
    if supporter[i] not in m4_performer:
        m4_supporter.append(supporter[i])
print(m4_supporter) #仕事入れる人
num_m4_sup = len(m4_supporter)
#Mediaの仕事を連続にならないように割り当て
table.ix[m3:m4 - 1,'Media'] = m4_supporter[randint(num_m4_sup)]
while table.ix[m3,'Media'] == table.ix[m2_2,'Media']:
    table.ix[m3:m4 - 1, 'Media'] = m4_supporter[randint(num_m4_sup)]
#カメラを決める、ビデオとかぶってたらもう1回
table.ix[m3:m4 - 1,'Cam&Time']= m4_supporter[randint(num_m4_sup)]
while table.ix[m3,'Cam&Time'] == table.ix[m3,'Media'] or table.ix[m3,'Cam&Time'] == table.ix[m2_2,'Cam&Time']:
    table.ix[m3:m4 - 1, 'Cam&Time'] = m4_supporter[randint(num_m4_sup)]
#照明を決める、ビデオ、カメラとかぶってたらもう1回
table.ix[m3:m4 - 1,'照明'] = m4_supporter[randint(num_m4_sup)]
while table.ix[m3,'照明'] == table.ix[m3,'Media'] or table.ix[m3,'照明'] == table.ix[m3,'Cam&Time'] or table.ix[m3,'照明'] == table.ix[m2_2,'照明']:
    table.ix[m3:m4 - 1, '照明'] = m4_supporter[randint(num_m4_sup)]


table.to_excel('worktable.xlsx')
