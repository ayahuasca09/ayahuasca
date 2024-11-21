import module.cloudfeishu.cloudfeishu_h as cloudfeishu_h
import config
import os
import module.excel.excel_h as excel_h
from pprint import pprint
import re

is_pass = True

"""根目录获取"""
# 获取当前脚本的文件名及路径
script_name_with_extension = __file__.split('/')[-1]
# 去掉扩展名
script_name = script_name_with_extension.rsplit('.', 1)[0]
# 替换为files
root_path = script_name.replace("func", "files")
# print(f"当前脚本的名字是: {root_path}")

"""文件获取"""
# 规范检查表
excel_check_path = os.path.join(root_path, config.excel_check_name)
# print(wb)

"""在线表获取"""
# 规范检查表
check_name_token = config.check_name_token
cloudfeishu_h.download_cloud_sheet(check_name_token, excel_check_path)
sheet_config, wb = excel_h.excel_get_sheet(excel_check_path, "配置说明")

"""初始化配置说明"""
# 初始化空列表
first_column_data = []

# 遍历工作表的第一列,索引为表格行-1
for row in sheet_config.iter_rows(min_row=1, max_col=1, max_row=sheet_config.max_row):
    for cell in row:
        first_column_data.append(cell.value)

"""********************功能函数***********************"""
"""检测正则表达式的pattern是否正确"""


def is_valid_regex(pattern):
    try:
        re.compile(pattern)
        return True, None
    except re.error as e:
        print(f"正则表达式不合规。错误信息: {str(e)}")
        return False


"""分隔符的大小写和长度检查"""


def check_by_length_and_word(name):
    # pprint(cell_sound.value)
    word_list = name.split("_")
    if len(word_list) <= config.word_list_len:
        is_title = True
        for word in word_list:
            """判定_的每个都不能超过10个"""
            check_by_str_length(word, config.one_word_len, name)
            """判定_的每个开头必须大写"""
            if len(word) > 0:
                # 只需要检查第一个字符，因为istitle会导致检查字符串只能首字母大写
                if word[0].istitle() == False and word[0].isdigit() == False:
                    is_title = False
                    break
        if is_title == False:
            print_error(name + "：通过”_“分隔的每个单词开头都需要大写")
    else:
        print_error(name + "：通过”_“分隔的单词个数不能超过" + str(config.word_list_len))

    if word_list:
        return word_list[0]
    else:
        return None


"""字符串长度检查"""


def check_by_str_length(str1, length, name):
    if len(str1) > length:
        if "_" in str1:
            print_error(name + "：描述后缀名过长，限制字符数为" + str(length))
        else:
            print_error(name + "：通过_切分的某个单词过长，每个单词长度不应超过10个字符，需要再拆分")


"""基础规范检查"""


def check_basic(media_name, audio_unit_list, event_unit_list):
    global is_pass
    # 检查是否为空
    if media_name:
        # 检查是否中文
        if not check_is_chinese(media_name):
            check_by_length_and_word(media_name)
            check_by_wb(media_name, audio_unit_list, event_unit_list)
        else:
            is_pass = False
    else:
        is_pass = False

    return is_pass


"""正则匹配"""


def check_by_re(pattern, name, media_name):
    # 将双反斜杠替换为单反斜杠
    pattern = pattern.replace('\\\\', '\\')
    if is_valid_regex(pattern):
        result = re.search(pattern, name)
        if result == None:
            print_error(media_name + "：请检查" + pattern + "是否拼写错误或未添加到列表中")
        # else:
        #     print(name)


"""log捕获"""


def print_warning(warn_info):
    print("[warning]" + warn_info)


"""报错捕获"""


def print_error(error_info):
    global is_pass
    is_pass = False
    print("[error]" + error_info)
    return False


"""检测字符串是否含有中文"""


def check_is_chinese(string):
    for ch in string:
        if u'\u4e00' <= ch <= u'\u9fff':
            return True
    return False


"""存储需要创建的unit结构"""


def get_audio_unit_list():
    # 用于存储unit值
    audio_unit_list = []

    sheet_names = wb.sheetnames
    # 加载所有工作表
    for sheet_name in sheet_names:
        sheet = wb[sheet_name]
        """获取工作表中的系统名"""
        # 获取工作表的颜色
        if sheet.sheet_properties.tabColor:
            # upper：字符串转大写
            if sheet.sheet_properties.tabColor.rgb.upper() == first_column_data[1]:
                # print(sheet_name)
                audio_unit_list.append(sheet_name)
                """获取工作列中的系统名"""
                # 从第二行开始遍历每行
                for row in sheet.iter_rows(min_row=2):
                    row_unit_name = sheet_name
                    # 遍历每列的单元格
                    for cell in row:
                        if cell.column == 1:
                            # 检查若为中文则跳过
                            if check_is_chinese(cell.value):
                                break
                        # print(cell.value)
                        # 获取单元格的背景填充色
                        fill = cell.fill
                        if fill.fill_type == 'solid':
                            if fill.start_color.index == first_column_data[1]:
                                row_unit_name += "_" + cell.value
                                audio_unit_list.append(row_unit_name)
                                # print(row_unit_name)

    return audio_unit_list


# pprint(get_audio_unit_list())

"""Audio Unit检测"""


# begin_name：资源表的前一个名字
# sheet：规范检查表的当前sheet
# end_name：资源表的后一个名字
# cell：规范检查表单元格是否有标注，例如填色或颜色
# audio_unit_list：资源表中用来收集unit的名字

def check_audio_unit(begin_name, sheet, end_name, cell, audio_unit_list):
    unit_name = begin_name
    if begin_name:
        # 表格名称获取
        if sheet.sheet_properties.tabColor:
            # upper：字符串转大写
            if sheet.sheet_properties.tabColor.rgb.upper() == first_column_data[1]:
                if unit_name not in audio_unit_list:
                    audio_unit_list.append(unit_name)

    if end_name:
        # 表格内容获取
        fill = cell.fill
        if fill.fill_type == 'solid':
            if fill.start_color.index == first_column_data[1]:
                unit_name += "_" + end_name
                if unit_name not in audio_unit_list:
                    audio_unit_list.append(unit_name)

    return unit_name


"""Event Unit检测"""


def check_event_unit(begin_name, end_name, cell, audio_unit_list, event_unit_list):
    unit_name = begin_name
    if begin_name and end_name:
        # 表格内容获取
        font_color = cell.font.color.rgb
        if font_color == first_column_data[2]:
            unit_name += "_" + end_name
            if (unit_name not in audio_unit_list) and (unit_name not in event_unit_list):
                event_unit_list.append(unit_name)

    return unit_name


"""命名规范表检查"""


def check_by_wb(media_name, audio_unit_list, event_unit_list):
    global is_pass
    sheet_names = wb.sheetnames

    # 解析media_name的数据
    media_name_list = media_name.split('_')
    # print(media_name_list)

    sys_name_flag = 0
    begin_name = ""
    # 加载所有工作表
    for sheet_name in sheet_names:
        """media_name_list[0]"""
        if media_name_list[0] == sheet_name:
            sys_name_flag = 1

            sheet = wb[sheet_name]
            begin_name = check_audio_unit(media_name_list[0], sheet, "", None, audio_unit_list)
            # 获取列数
            max_column = sheet.max_column
            # print(max_column)
            # 遍历列
            flag = 0
            check_row = ""
            for row in sheet.iter_rows(min_row=1, max_col=1, max_row=sheet.max_row):
                for cell in row:
                    if check_is_chinese(cell.value):
                        break
                    # 检查是哪一行的规则
                    # 检查是否带*，带*需要跳过
                    elif cell.value == "*":
                        flag = 2
                        is_pass = True
                        break
                    elif cell.value == media_name_list[1]:
                        check_row = cell.row
                        begin_name = check_audio_unit(begin_name, sheet, media_name_list[1], cell, audio_unit_list)
                        flag = 1
                        break
            # print(media_name)
            # print(is_pass)
            # 遍历选中规则的行
            if flag == 1 and check_row:
                for i in range(2, sheet.max_column + 1):
                    if len(media_name_list) > i:
                        # 为*则跳过检查
                        if sheet.cell(row=check_row,
                                      column=i).value != "*":
                            is_pass = True
                            check_by_re(sheet.cell(row=check_row,
                                                   column=i).value, media_name_list[i], media_name)
                            if is_pass:
                                begin_name = check_audio_unit(begin_name, sheet, media_name_list[i],
                                                              sheet.cell(row=check_row,
                                                                         column=i), audio_unit_list)
                                check_event_unit(media_name_list[0], media_name_list[i], sheet.cell(row=check_row,
                                                                                                    column=i),
                                                 audio_unit_list, event_unit_list)
                        # else:
                        #     print(is_pass)


            elif flag == 0:
                print_error(media_name + "：" + media_name_list[1] + "未找到，请检查命名是否正确")

            break

    if sys_name_flag == 0:
        print_error(media_name + "：" + media_name_list[0] + "未找到，请检查命名是否正确")
    return is_pass

# 检查命名规范测试
# audio_unit_list = []
# event_unit_list = []
# name_list = ['Char_Skill_C01_Counter', 'Char_Skill_C01_Execution1_Stand', 'Char_Skill_C01_Execution2_Stand',
#              'Char_Skill_C01_Strafe_Exit']
# for name in name_list:
#     if check_by_wb(name, audio_unit_list, event_unit_list):
#         print(name)

#
# # unit列表长度排序
# audio_unit_list.sort(key=len)
# event_unit_list.sort(key=len)
# pprint(audio_unit_list)
# pprint(event_unit_list)
