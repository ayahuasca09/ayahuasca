import re

"""测试名称"""
name = "Char_C03_Skill_Execu1_01_Pre_Hook_LP"



"""测试结果类型"""
# 查看匹配的位置和名称
# <re.Match object; span=(0, 1), match='C'>
# result = re.search(pattern, name)
# print(result)

"""Char检查规则范例"""
# 以Char开头,命名分组以供输出(?P<正则>)
# pattern = ("(?P<first>"
#            "^Char"
#            "_"
#            # 角色名称：C01，C02……
#            "C\\d{2,4}"
#            ")"
#            "_"
#            # 所属哪个模块
#            "((Skill)|(Foley))"
#            "_"
#            # 所属哪个Actor-Mixer,*d代表出现0或多次数字
#            "((Atk\\d*)|(Hook)|(SkyAtk)|(Battle)|"
#            "(Dodge)|(Counter)|"
#            "(Execu\\d*)|"
#            "(Focus)|"
#            "(Gen)|"
#            "(Hit)|(Death)|"
#            "(Strafe)|"
#            "(Ult))"
#            "_"
#            # 最后的状态描述，不能超过20个
#            ".*"
#            )

"""正则取代_LP测试"""
pattern = r"_LP$"
# new_name = re.sub(pattern, "", name)
# print(new_name)

"""正则去除数字末尾"""
# rnd_name = re.sub(r"(_\d{2,4})$", "", name)
# print(rnd_name)

"""正则输出ID"""
aa = "id:ssd|hdabgk|7391"
data = re.search(r'(?<=id:)\d+|$', aa)
if data.group():
    print(int(data.group()))

