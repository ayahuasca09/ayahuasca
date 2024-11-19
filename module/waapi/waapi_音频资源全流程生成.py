import subprocess

# 媒体资源替换及generated soundbank
subprocess.run(['python', r"F:\pppppy\SP\module\waapi\waapi_音频资源导入自动化\waapi_音频资源导入_从媒体资源导入.py"])

# Wwise属性自动设置
subprocess.run(['python', r"F:\pppppy\SP\module\waapi\waapi_音频资源导入自动化\waapi_wwise属性自动配置.py"])

# reconcile 同步UE资产
command = (r"S:\Ver_1.0.0\Editor\Engine\Binaries\Win64\UnrealEditor.exe S:\Ver_1.0.0\Project\SilverPalace.uproject "
           r"-run=WwiseReconcileCommandlet -modes=all")
subprocess.run(command, capture_output=True, text=True, check=True, shell=True)
# command = (
#     r"S:\chen.gong_weekly\Editor\Engine\Binaries\Win64\UnrealEditor.exe S:\Ver_1.0.0\Project\SilverPalace.uproject "
#     r"-run=WwiseReconcileCommandlet -modes=all")
# subprocess.run(command, capture_output=True, text=True, check=True, shell=True)

# 生成id表
subprocess.run(['python', r"F:\pppppy\SP\module\ue\ue_ID表生成\ue_ID表生成v3.py"])

# 生成id表
subprocess.run(['python', r"F:\pppppy\SP\module\ue\ue_ID表生成\ue_ID表生成v3.py"])

# dcc audio表数据生成
subprocess.run(['python', r"F:\pppppy\SP\module\ue\ue_ID表生成\DCC表的引用路径生成.py"])

# 在ue中导入数据表
