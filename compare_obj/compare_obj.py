import random


class MyClass():
    def __init__(self):
        self.dict = {}


def create_2_cls_obj_list(list1, list2):
    obj_list1 = []
    obj_list2 = []

    for dict1 in list1:
        obj = MyClass()
        obj.dict = dict1
        obj_list1.append(obj)

    for dict2 in list2:
        obj = MyClass()
        obj.dict = dict2
        obj_list2.append(obj)

    return obj_list1, obj_list2


def compare_obj(list1, list2):
    obj_list1, obj_list2 = create_2_cls_obj_list(list1, list2)
    all_compare_list = create_all_compare_list(obj_list1, obj_list2)
    all_compare_list = sorted(all_compare_list, key=lambda group: group[2], reverse=True)

    all_list_len = len(all_compare_list)
    align_list1 = []
    align_list2 = []
    for i in range(all_list_len):
        group = all_compare_list[i]
        obj1 = group[0]
        obj2 = group[1]
        degree = group[2]
        if degree > 0 and obj1 not in align_list1 and obj2 not in align_list2:
            align_list1.append(obj1)
            align_list2.append(obj2)

            obj_list1.remove(obj1)
            obj_list2.remove(obj2)
    # 剩下list1中未匹配到的
    for obj1 in obj_list1:
        align_list1.append(obj1)
        align_list2.append(None)
    # 剩下list2中未匹配到的
    for obj2 in obj_list2:
        align_list1.append(None)
        align_list2.append(obj2)
    return align_list1, align_list2


def create_all_compare_list(obj_list1, obj_list2):
    all_compare_list = []
    for obj1 in obj_list1:
        for obj2 in obj_list2:
            compare_group = compare_objects_degree(obj1, obj2)
            all_compare_list.append(compare_group)
    return all_compare_list


def compare_objects_degree(obj1, obj2):
    obj1_keys = obj1.dict.keys()
    obj2_keys = obj2.dict.keys()

    same_keys = set(obj1_keys) & set(obj2_keys)
    same_name_count = len(same_keys)
    same_value_count = 0
    for key in same_keys:
        if obj1.dict[key] == obj2.dict[key]:
            same_value_count += 1
    if same_name_count == same_value_count != 0:
        return [obj1, obj2, 1000]
    return [obj1, obj2, same_value_count]


def out_single_file(fb, obj):
    # 输出丢失或新增的对象
    for key, value in obj.dict.items():
        fb.write(f"  |--{key}={value}\n")
    fb.write("\n\n")


def out_diff_file(fb, obj1, obj2):
    # 输出未丢失或新增的的对象对比
    keys1 = list(obj1.dict.keys())
    keys2 = list(obj2.dict.keys())
    for key in keys2:
        if key not in keys1:
            keys1.append(key)
    keys = keys1
    for key in keys:
        if key in obj1.dict.keys() and key in obj2.dict.keys():
            fb.write(f"  |--{key}={obj1.dict[key]}".ljust(40, ' ') + f"  |--{key}={obj2.dict[key]}\n")
        elif key in obj1.dict.keys():
            fb.write(f"  |--{key}={obj1.dict[key]}".ljust(40, ' ') + "  |-- =\n")
        else:
            fb.write("  |-- = ".ljust(40, ' ') + f"  |--{key}={obj2.dict[key]}\n")
    fb.write("\n\n")


def out_all_files(list1, list2):
    lost_obj_file = "result/丢失的对象.txt"
    add_obj_file = "result/新增的对象.txt"
    diff_obj_file = "result/对象对比.txt"
    obj_list1, obj_list2 = compare_obj(list1, list2)

    len_objs = len(obj_list1)

    with open(lost_obj_file, 'w') as lost_file:
        with open(add_obj_file, 'w') as add_file:
            with open(diff_obj_file, 'w') as diff_file:
                for i in range(len_objs):
                    obj1 = obj_list1[i]
                    obj2 = obj_list2[i]
                    if obj1 is None:
                        out_single_file(add_file, obj2)
                    if obj2 is None:
                        out_single_file(lost_file, obj1)
                    if obj1 and obj2:
                        out_diff_file(diff_file, obj1, obj2)


if __name__ == '__main__':
    dict_list1 = [
        {
            "id": 1,
            "name": "name1",
            "age": 11,
            "gender": "woman",
            "color": "blue",
        },
        {
            "id": 2,
            "name": "name2",
            "age": 5,
            "gender": "woman",
            "color": "pink",
        },
        {
            "id": 3,
            "name": "name3",
            "age": 55,
            "gender": "man",
            "color": "black",
        },
        {
            "id": 4,
            "name": "name4",
            "age": 2,
            "gender": "man",
            "color": "orange",
        },
    ]

    dict_list2 = [
        {
            "id": 1,
            "name": "name111",
            "age": 16,
            "gender": "woman",
            "color": "blue:pink:ooo",
        },
        {
            "id": 6,
            "name": "name6",
            "age": 6,
            "gender": "",
            "color": "pnk",
        },
        {
            "id": 3,
            "name": "name3",
            "age": 20,
            "gender": "man",
            "color": "black",
            "area": "America",
        },
        {
            "id": 4,
            "name": "name4",
            # "gender": "man",
            "color": "",
        },
        {
            "id": 20,
            "name": "name20",
            "gender": "man",
            "color": "green",
            "music": "sunset......",
        },
    ]
    out_all_files(dict_list1, dict_list2)
