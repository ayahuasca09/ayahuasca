import os.path
import subprocess
from . import config

wwise_console_path = config.wwise_console_path
wwise_path = config.wwise_path
wproj_path = os.path.join(wwise_path, 'SilverPalace_WwiseProject.wproj')

ue_gensoundbank_path = config.ue_gensoundbank_path
# print(ue_gensoundbank_path)
gensoundbank_output_win_path = config.gensoundbank_output_win_path
gensoundbank_output_android_path = config.gensoundbank_output_android_path
gensoundbank_output_ios_path = config.gensoundbank_output_ios_path

ue_genes_path = config.ue_genes_path
external_output_win_path = config.external_output_win_path
# external_output_android_path = 'S:\\Ver_1.0.0\\Project\\Content\\Audio\\GeneratedExternalSources\\Android'
external_output_android_path = config.external_output_android_path
# external_output_ios_path = 'S:\\Ver_1.0.0\\Project\\Content\\Audio\\GeneratedExternalSources\\iOS'
external_output_ios_path = config.external_output_ios_path

"""UE资产同步"""
ue_proj_path = config.ue_project_path
# SP工程路径
ue_sp_path = os.path.join(ue_proj_path, "SilverPalace.uproject")
ue_root_path = os.path.dirname(ue_proj_path)
# print(ue_sp_path)
ue_editor_path = config.ue_editor_path
if not config.ue_editor_path:
    ue_editor_path = os.path.join(ue_root_path, "Editor", "Engine", "Binaries", "Win64", "UnrealEditor.exe")


# print(ue_editor_path)

def run_ue_reconcile():
    command_path = ue_editor_path + " " + ue_sp_path + " " + "-run=WwiseReconcileCommandlet -modes=all" + " -nocompile"
    # print(command_path)
    subprocess.run(command_path, capture_output=True, text=True, check=True, shell=True)


# run_ue_reconcile()


"""执行exe文件"""


def run_exe(exe_path):
    try:
        # 使用 subprocess.run 来运行 .exe 文件
        result = subprocess.run([exe_path], check=True)
        print("Process finished with return code:", result.returncode)
        print(exe_path + "已执行")
    except subprocess.CalledProcessError as e:
        print("Execution failed:", e)


# run_exe(r"媒体资源替换及随机资源新增\媒体资源替换及随机资源新增打包版.exe")
# run_exe(r"Wwise属性配置\Wwise属性配置打包版.exe")
# run_exe(r"ID表生成\ID表生成打包版.exe")


"""soundbank生成"""


def gen_soundbank():
    args = [
        'generate-soundbank',
        wproj_path,
        '--soundbank-path', 'Windows', gensoundbank_output_win_path,
        '--soundbank-path', 'Android', gensoundbank_output_android_path,
        '--soundbank-path', 'iOS', gensoundbank_output_ios_path,
        '--root-output-path', ue_gensoundbank_path
    ]

    # 执行命令
    try:
        subprocess.run([wwise_console_path] + args, check=True)
        print("Command executed successfully")
    except subprocess.CalledProcessError as e:
        print(f"An error occurred: {e}")


"""ES生成"""


def gen_es(external_input_path):
    args = [
        'convert-external-source',
        wproj_path,
        '--output', 'Windows', external_output_win_path,
        '--output', 'Android', external_output_android_path,
        '--output', 'iOS', external_output_ios_path,
        '--source-file', external_input_path
    ]

    # 执行命令
    try:
        subprocess.run([wwise_console_path] + args, check=True)
        print("Command executed successfully")
    except subprocess.CalledProcessError as e:
        print(f"An error occurred: {e}")


"""AI语音生成"""


def gen_ai_language(external_input_path, language):
    # 若没有路径需要创建
    win_output = os.path.join(external_output_win_path, language, 'AILanguage')
    android_output = os.path.join(external_output_android_path, language, 'AILanguage')
    ios_output = os.path.join(external_output_ios_path, language, 'AILanguage')
    if not os.path.exists(win_output):
        os.makedirs(win_output)
    if not os.path.exists(android_output):
        os.makedirs(android_output)
    if not os.path.exists(ios_output):
        os.makedirs(ios_output)

    args = [
        'convert-external-source',
        wproj_path,
        '--output', 'Windows', os.path.join(external_output_win_path, language, 'AILanguage'),
        '--output', 'Android', os.path.join(external_output_android_path, language, 'AILanguage'),
        '--output', 'iOS', os.path.join(external_output_ios_path, language, 'AILanguage'),
        '--source-file', external_input_path
    ]

    # 执行命令
    try:
        subprocess.run([wwise_console_path] + args, check=True)
        print("Command executed successfully")
    except subprocess.CalledProcessError as e:
        print(f"An error occurred: {e}")


"""正式语音生成"""


def gen_es_language(external_input_path, language):
    # 若没有路径需要创建
    win_output = os.path.join(external_output_win_path, language)
    android_output = os.path.join(external_output_android_path, language)
    ios_output = os.path.join(external_output_ios_path, language)
    if not os.path.exists(win_output):
        os.makedirs(win_output)
    if not os.path.exists(android_output):
        os.makedirs(android_output)
    if not os.path.exists(ios_output):
        os.makedirs(ios_output)

    args = [
        'convert-external-source',
        wproj_path,
        '--output', 'Windows', os.path.join(external_output_win_path, language),
        '--output', 'Android', os.path.join(external_output_android_path, language),
        '--output', 'iOS', os.path.join(external_output_ios_path, language),
        '--source-file', external_input_path
    ]

    # 执行命令
    try:
        subprocess.run([wwise_console_path] + args, check=True)
        print("Command executed successfully")
    except subprocess.CalledProcessError as e:
        print(f"An error occurred: {e}")


"""在ue外部执行"""


def run_unreal_editor_with_script(script_path):
    print("*************UE导表开始执行**************")
    # Construct the command
    command = [
        ue_editor_path,  # Path to the Unreal Engine executable
        ue_sp_path,  # Path to your .uproject file
        "-run=pythonscript",
        f"-script={script_path}"  # Path to your Python script
    ]

    # Print the command to verify correctness
    print("Running command:", " ".join(command))

    # Execute the command
    try:
        subprocess.run(command, check=True)
        print("Script executed successfully.")
    except subprocess.CalledProcessError as e:
        print("An error occurred while executing the script:", e)
    print("*************UE导表执行完毕**************")
    print("*************若UE处于开启状态，需重启UE才能看到导表修改**************")

# # Set your paths here
# python_script_path = r"F:\pppppy\SP\module\ue\ue_h.py"  # Update this path
#
# # Run the function
# run_unreal_editor_with_script(ue_editor_path, ue_sp_path, python_script_path)
