"""一级系统检测"""
# def check_first_system_name(name):
#     if word_list[0] == "Amb":
#         return name.replace("Amb_", "")
#     elif word_list[0] == "Cg":
#         return name.replace("Cg_", "")
#     elif word_list[0] == "Char":
#         system_name = "Char"
#         # 如果检测通过
#         if check_by_char(name.replace(system_name + "_", "")):
#             create_wwise_content(cell_sound.value, system_name)
#     elif word_list[0] == "Imp":
#         return name.replace("Imp_", "")
#     elif word_list[0] == "Mon":
#         system_name = "Mon"
#         # 如果检测通过
#         if check_by_char(name.replace(system_name + "_", "")):
#             create_wwise_content(cell_sound.value, system_name)
#     elif word_list[0] == "Mus":
#         return name.replace("Mus_", "")
#     elif word_list[0] == "Sys":
#         check_by_sys(name.replace("Sys_", ""))
#     elif word_list[0] == "VO":
#         return name.replace("VO_", "")
#     else:
#         print(name + "：一级系统名称（如Amb、Char）等有误，请检查拼写")
#         return name

"""获取随机容器的父级：需要弃用"""
#
#
# # 数据获取
#
# def find_rnd_parent(path, rnd_name):
#     # 所有Actor-Mixer获取
#     mixer_list, mixer_id, _ = find_obj(
#         {'waql': ' "%s" select descendants,this where type = "ActorMixer" ' % path})
#     # pprint(mixer_container_list)
#     # [{'id': '{446650AE-60B4-40A0-8B14-3470DBBB83B0}', 'name': 'Amb', 'notes': ''},
#     #  {'id': '{E0A2685C-387A-4180-BB45-6E6C1559E7F8}', 'name': 'CG', 'notes': ''}]
#     # 存储符合的rnd的父级，最后找出最长最符合的
#     mixer_len = 0
#     rnd_parent_id = {}
#     rnd_parent_path = ""
#     rnd_parent_name = ""
#     for mixer_dict in mixer_list:
#         if mixer_dict['name'] in rnd_name:
#             if len(mixer_dict['name']) > mixer_len:
#                 mixer_len = len(mixer_dict['name'])
#                 rnd_parent_id = mixer_dict['id']
#     return rnd_parent_id

"""已废弃：按照json配置检查"""
# def check_by_json_by_re(system_dict):
#     pattern = (
#         # 一级结构
#             "(?P<the1>" + "^" +
#             word_list[0] +
#             ")" + "_" +
#
#             # module
#             "(?P<the2>" +
#             system_dict['module'] +
#             ")" + "_" +
#
#             # name
#             "(?P<the3>" +
#             system_dict['name'] +
#             ")" + "_"
#
#             # property
#                   "(?P<the4>" +
#             system_dict['property'] +
#             ")" + "_*"
#
#             # 其他
#                   "(?P<the5>" +
#             ".*" +
#             ")"
#     )
#     result = re.search(pattern, name)
#     if result == None:
#         print((name + "：中间有字段匹配错误，请检查是否为结构顺序或拼写错误"))
#     else:
#         data_list = re.finditer(pattern, name)
#         for item in data_list:
#             item_dict = item.groupdict()
#             # pprint(item_dict)
#             # {'the1': 'Char',
#             #  'the2': 'Skill',
#             #  'the3': 'C01',
#             #  'the4': 'Focus',
#             #  'the5': 'Ready2'}
#
#             # 字符串长度检查
#             check_by_str_length(item_dict['the5'], system_dict['length'])

"""怪物类型检查"""
# def check_by_Mon(name):
#     # if "Mob_" in name:
#     #     name = name.replace("Mob_", "")
#     #     check_by_mob(name, system_name)
#     # elif "Boss_" in name:
#     #     name = name.replace("Boss_", "")
#     #     check_by_boss(name, system_name)
#     # else:
#     #     print_error("：Boss、Mob拼写有误或漏写，请检查")

"""Boss和Mob检查"""
# def check_by_Boss(name, system_name):
#     module_dict = js_dict['Mon']['Boss']['module']
#     flag = 0
#     for module in module_dict:
#         # 模块层查询
#         if module + "_" in name:
#             name = name.replace(module + "_", "")
#             # Boss名层查询
#             name = check_by_re(js_dict['Mon']['Boss']['name'] + "_", name)
#             if name != None:
#                 # 技能层查询
#                 name = check_by_re(module_dict[module]['property'] + "_*", name)
#                 # 长度限制查询
#                 if name != None:
#                     check_by_str_length(name, js_dict['Mon']['Boss']['length'])
#
#             flag = 1
#             break
#
#     if flag == 0:
#         print_error_by_module(module_dict, system_name)
#
#
# """Mob类型检查"""
#
#
# def check_by_Mob(name, system_name):
#     module_dict = js_dict['Mon']['Mob']['module']
#     flag = 0
#     for module in module_dict:
#         # 模块层查询
#         if module + "_" in name:
#             name = name.replace(module + "_", "")
#             # 小怪名层查询
#             name = check_by_re(js_dict['Mon']['Mob']['name'] + "_", name)
#             if name != None:
#                 # 技能层查询
#                 name = check_by_re(module_dict[module]['property'] + "_*", name)
#                 # 长度限制查询
#                 if name != None:
#                     check_by_str_length(name, js_dict['Mon']['Mob']['length'])
#             flag = 1
#             break
#
#     if flag == 0:
#         print_error_by_module(module_dict, system_name)

"""将事件写入json"""
# js_event_info_path = r"F:\pppppy\SP\module\ue\ue_ID表生成\js_event_info.json"
# event_list, event_id, _ = find_obj(
#     {'waql': ' "%s" select descendants where type = "Event" ' % wwise_dict['Event_Root']})
# js_event_info = json.dumps(event_list, indent=4)
# with open(js_event_info_path, 'w') as writeJs:
#     writeJs.write(js_event_info)

"""事件创建ID"""
# # 保存已有事件的id
# id_list = []
# for event_dict in event_list:
#     data = re.search(r'(?<=id:)\d+|$', event_dict['notes'])
#     if data.group():
#         id_list.append(int(data.group()))
#
# for event_dict in event_list:
#     # 获取值域范围
#     get_module_range(event_dict['name'])
#     if (range_min != 0) and (range_max != 0) and range_min and range_max:
#         # id创建
#         tmp = range_min
#         while tmp in id_list:
#             tmp = tmp + 1
#         args = {
#             'object': event_dict['id'],
#             'value': "id:" + str(tmp) + "|"
#         }
#         client.call("ak.wwise.core.object.setNotes", args)

"""获取ID范围"""
# def get_module_range(event_name):
#     global range_min, range_max
#     flag = 0
#     for module_name in js_dict:
#         if str(module_name) in str(event_name):
#             range_min = js_dict[module_name].get('min')
#             range_max = js_dict[module_name].get('max')
#             flag = 1
#             return
#     if flag == 0:
#         range_min = 0
#         range_max = 0
#         print_error("找不到" + event_name + "的系统名，无法设置ID")

"""从媒体资源导入"""
# args_import = {
#             # createNew
#             # useExisting：会增加一个新媒体文件但旧的不会删除
#             # replaceExisting:会销毁Sound，上面做的设置都无了
#             "importOperation": "replaceExisting",
#             "default": {
#                 "importLanguage": "SFX"
#             },
#             "imports": [
#                 {
#                     "audioFile": source_path,
#                     "objectPath": rnd_path + '\\<Sound SFX>' + media_name,
#                     "originalsSubFolder": word_list[0]
#                     #                                                         名为Test 0的顺序容器            名为My SFX 0 的音效
#                     # "objectPath": "\\Actor-Mixer Hierarchy\\Default Work Unit\\<Sequence Container>Test 0\\<Sound SFX>My SFX 0"
#                 }
#             ]
#         }
#         # 定义返回结果参数，让其只返回 Windows 平台下的信息，信息中包含 GUID 和新创建的对象名
#         opts = {
#             "platform": "Windows",
#             "return": [
#                 "id", "name"
#             ]
#         }
#         import_info = client.call("ak.wwise.core.audio.import", args_import, options=opts)
