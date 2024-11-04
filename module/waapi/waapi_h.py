"""获取对象信息"""


def find_obj(obj_list):
    if not obj_list:
        obj_sub_id = ""
        obj_sub_path = ""
    else:
        obj_sub_id = obj_list[0]['id']
        obj_sub_path = obj_list[0]['path']
    return obj_list, obj_sub_id, obj_sub_path


"""WAQL查询"""


# 通过type查询
def waql_by_type(waql_type, obj_id_or_path):
    return {'waql': ' "%s" select descendants where type = "%s" ' % (obj_id_or_path, waql_type)}


# 通过类型查询所有
def waql_from_type(waql_type):
    return {'waql': f'from type  {waql_type} '}


# 通过ID查询对象
def waql_from_id(waql_id):
    return {'waql': 'from object  "%s" ' % waql_id}
