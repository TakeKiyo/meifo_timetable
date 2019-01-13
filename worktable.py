import pandas as pd
from numpy.random import randint

print('エクセルのファイル名を入力')
filename = str(input())
filename = filename + '.xlsx'
print(filename)
table = pd.read_excel(filename)
table = table.fillna('-')
n = len(table)
print(table)
#とりあえず、サポートを全て4バンドとした

support_drummer = []
support_bassist = []
support_keyboardist = []
support_guitarist = []

print('------サポートに入れる人の名前を入力------')

print('ドラマーの人数を入力')
num_s_dr = int(input())
print('ドラムのサポートに入れる人の名前を'+str(num_s_dr)+'人入力してください。')
for i in range(num_s_dr):
    support_drummer.append(str(input()))

print('ベースの人数を入力')
num_s_bass = int(input())
print('ベースサポートに入れる人の名前を'+str(num_s_bass)+'人入力してください。')
for i in range(num_s_bass):
    support_bassist.append(str(input()))

print('キーボードの人数を入力')
num_s_keyboard = int(input())
print('キーボードサポートに入れる人の名前を'+str(num_s_keyboard)+'人入力してください。')
for i in range(num_s_keyboard):
    support_keyboardist.append(str(input()))

print('ギターの人数を入力')
num_s_guitar = int(input())
print('ギターサポートに入れる人の名前を'+str(num_s_guitar)+'人入力してください。')
for i in range(num_s_guitar):
    support_guitarist.append(str(input()))

all_supporter = support_drummer + support_bassist + support_keyboardist + support_guitarist
print(all_supporter)



#ドラムのサポートをここから割り振り
#まずランダムに割り振り
for i in range(n):
    table.at[i,'ドラムサポート'] = support_drummer[randint(num_s_dr)]
    #自分のサポートになっていたらやり直し
    while table.at[i,'Dr.'] == table.at[i,'ドラムサポート']:
        table.at[i, 'ドラムサポート'] = support_drummer[randint(num_s_dr)]
    #自分の出番の次は自分がサポートをする
    for j in range(n-1):
        if table.at[j,'Dr.'] in support_drummer:
            table.at[j+1,'ドラムサポート'] = table.at[j,'Dr.']
    for k in range(n):
        if table.at[k, 'Dr.'] == "-":
            table.at[k, 'ドラムサポート'] = "-"

#ベースとキーボードの割り振り方は同じなので関数にした
def ba_key_table(num_s,name,support_column,support_list):
    #まずは割り振り
    for i in range(n):
        table.at[i,support_column] = support_list[randint(num_s)]
    #サポートの連続をなるべく避ける
    for h in range(1,n):
        while table.at[h,support_column] == table.at[h-1,support_column]:
            table.at[h,support_column] = support_list[randint(num_s)]
    #演奏後のサポートになるべくならないようにする
    for k in range(n-1):
        while table.at[k,name] == table.at[k+1,support_column]:
            table.at[k+1,support_column] = support_list[randint(num_s)]
    #演奏の前のサポートは避ける
    for l in range(1,n):
        while table.at[l,name] == table.at[l-1,support_column]:
            table.at[l-1,support_column] = support_list[randint(num_s)]
    #自分のサポートはしない、-のサポートは-
    for j in range(n):
        while table.at[j,name] == table.at[j,support_column]:
            table.at[j,support_column] = support_list[randint(num_s)]
        if table.at[j,name] == '-':
            table.at[j,support_column] = '-'
    return table

ba_key_table(num_s_bass,'Ba.','ベースサポート',support_bassist)
ba_key_table(num_s_keyboard,'Key.','キーボサポート',support_keyboardist)

#ギターサポート1の割り振り
#まずランダムに割りあてる
for i in range(n):
    table.at[i,'ギターサポート1'] = support_guitarist[randint(num_s_guitar)]
#サポート連続にならないようにする
for h in range(1,n):
    while table.at[h,'ギターサポート1'] == table.at[h - 1,'ギターサポート1']:
        table.at[h,'ギターサポート1'] = support_guitarist[randint(num_s_guitar)]
#演奏の後にサポートにならないようにする
for k in range(n - 1):
    while table.at[k, 'Gt.1'] == table.at[k + 1, 'ギターサポート1']:
        table.at[k + 1, 'ギターサポート1'] = support_guitarist[randint(num_s_guitar)]
#演奏の前にサポートにならないようにする
for l in range(1,n - 1):
    while table.at[l, 'Gt.1'] == table.at[l - 1, 'ギターサポート1']:
        table.at[l - 1, 'ギターサポート1'] = support_guitarist[randint(num_s_guitar)]
#自分のサポートなし、ギター2とのかぶりなし、-のところは-
for j in range(n):
    while table.at[j, 'Gt.1'] == table.at[j, 'ギターサポート1'] or table.at[j, 'ギターサポート1'] == table.at[j, 'Gt.2']:
        table.at[j, 'ギターサポート1'] = support_guitarist[randint(num_s_guitar)]
    if table.at[j, 'Gt.1'] == "-":
        table.at[j, 'ギターサポート1'] = "-"

#ギターサポート2の割り振り
# まずランダムに割りあてる
for i in range(n):
    table.at[i, 'ギターサポート2'] = support_guitarist[randint(num_s_guitar)]
# サポート連続にならないようにする
for h in range(1, n):
    while table.at[h, 'ギターサポート2'] == table.at[h - 1, 'ギターサポート1']:
        table.at[h, 'ギターサポート2'] = support_guitarist[randint(num_s_guitar)]
# 演奏の後にサポートにならないようにする
for k in range(n - 1):
    while table.at[k, 'Gt.2'] == table.at[k + 1, 'ギターサポート2']:
        table.at[k + 1, 'ギターサポート2'] = support_guitarist[randint(num_s_guitar)]
# 演奏の前にサポートにならないようにする
for l in range(1, n - 1):
    while table.at[l, 'Gt.2'] == table.at[l - 1, 'ギターサポート2']:
        table.at[l - 1, 'ギターサポート2'] = support_guitarist[randint(num_s_guitar)]
# 自分のサポートなし、ギター１とのかぶりなし、ギター1のサポートとのかぶりなし、-のところは-
for j in range(n):
    while table.at[j, 'Gt.2'] == table.at[j, 'ギターサポート2'] or table.at[j, 'ギターサポート2'] == table.at[j, 'Gt.1'] or table.at[j,'ギターサポート1'] == table.at[j,'ギターサポート2']:
        table.at[j, 'ギターサポート2'] = support_guitarist[randint(num_s_guitar)]
    if table.at[j, 'Gt.1'] == "-":
        table.at[j, 'ギターサポート2'] = "-"

print(table)

print('分割数を入力')
m = int(input())
#例として4
index_num = 0
def media_camera_light(m):
    performer = []
    supporter = []
    for i in range(m):
        performer.append(table.at[index_num+i,'Vo.'])
        performer.append(table.at[index_num+i, 'Gt.1'])
        performer.append(table.at[index_num+i, 'Gt.2'])
        performer.append(table.at[index_num+i, 'Ba.'])
        performer.append(table.at[index_num+i, 'Dr.'])
        performer.append(table.at[index_num+i, 'Key.'])
        performer.append(table.at[index_num+i, 'その他'])
    while '-' in performer:
        performer.remove('-')
    for j in range(len(all_supporter)):
        if all_supporter[j] not in performer:
            supporter.append(all_supporter[j])

    num_of_supporters = len(supporter)

    # mediaを割り当て
    table.ix[index_num:index_num+m-1, 'Media'] = supporter[randint(num_of_supporters)]
    # ビデオにかぶらないようにカメラ割り当て
    table.ix[index_num:index_num+m-1, 'Cam&Time'] = supporter[randint(num_of_supporters)]
    while table.at[index_num, 'Cam&Time'] == table.at[index_num, 'Media']:
        table.ix[index_num:index_num + m-1, 'Cam&Time'] = supporter[randint(num_of_supporters)]
    # ビデオ、カメラにかぶらないように割り当て
    table.ix[index_num:index_num+m-1, '照明'] = supporter[randint(num_of_supporters)]
    while table.at[index_num, '照明'] == table.at[index_num, 'Media'] or table.at[index_num, '照明'] == table.at[index_num, 'Cam&Time']:
        table.ix[index_num:index_num + m-1, '照明'] = supporter[randint(num_of_supporters)]

    return table

for kaisuu in range(n//m):
    media_camera_light(m)
    index_num += m

media_camera_light(n-index_num)
print(table)
table.index = table.pop('時間')

table.to_excel('仕事表.xlsx')