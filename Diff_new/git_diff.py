path_file_A = "1.text"
path_file_B = "2.text"
path_patch_file = "12.text"
relations = dict()

def drop_nums(start_position, finis_position, min, max):
    global relations
    for i in range(start_position, finis_position + 1):
        list = relations.get(i)
        if list is None:
            continue
        for num in list:
            if num >= min and num <= max:
                relations.update({i: [num]})
                drop_from_rel(i, num)
                break

def drop_from_rel(index, num):
    global relations
    for i in range(len(relations)):
        if i + 1 == index:
            continue
        list = relations.get(i + 1)
        if num in list:
            list.remove(num)
            relations.update({i + 1: list})


with open(path_file_A, encoding="utf-8") as f_A:
    file_A = f_A.readlines()

with open(path_file_B, encoding="utf-8") as f_B:
    file_B = f_B.readlines()


#Сформировали список совпадающих строк в relations
for i in range(len(file_A)):
    line_A = file_A[i].rstrip()

    list = []
    for k in range(len(file_B)):
        line_B = file_B[k].rstrip()
        if line_A == line_B:
            list.append(k + 1)
        relations.update({i + 1: list})

#Находим якоря
was_find = True
while was_find:
    was_find = False
    anchor = 0
    for i in range(len(relations)):
        list = relations.get(i + 1)
        if len(list) == 1: # это якорь
            if anchor == 0 or anchor == i:
                anchor = i + 1
            else:
                min = relations.get(anchor)[0]
                max = list[0]
                drop_nums(anchor + 1, i, min, max)
                was_find = True
                break

        if i == len(relations) - 1 and 0 < anchor < i - 1:
            min = relations.get(anchor)[0]
            max = len(file_B)
            drop_from_rel(anchor + 1, i, min, max)
            was_find = True
            break
