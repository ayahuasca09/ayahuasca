import subprocess
import os
import shutil


def copy_directories(src, dst):
    # 确保目标路径存在
    if not os.path.exists(dst):
        os.makedirs(dst)

    # 遍历源路径下的所有文件夹
    for item in os.listdir(src):
        src_path = os.path.join(src, item)
        dst_path = os.path.join(dst, item)

        # 判断是否为文件夹
        if os.path.isdir(src_path):
            # 使用 copytree 函数，并设置 dirs_exist_ok=True 以替代同名资源
            shutil.copytree(src_path, dst_path, dirs_exist_ok=True)


def move_directories(src, dst):
    # 确保目标路径存在
    if not os.path.exists(dst):
        os.makedirs(dst)

    # 遍历源路径下的所有文件夹
    for dir_name in os.listdir(src):
        src_dir = os.path.join(src, dir_name)
        dst_dir = os.path.join(dst, dir_name)

        # 确认是文件夹
        if os.path.isdir(src_dir):
            # 如果目标文件夹已存在，先删除
            if os.path.exists(dst_dir):
                shutil.rmtree(dst_dir)

            # 移动文件夹
            shutil.move(src_dir, dst_dir)
            print(f"Moved {src_dir} to {dst_dir}")


output_path = r"S:\chen.gong_DCC_Audio\Audio\Tool\Auto_Sound\临时打包文件"


def package_script(script_name, output_name):
    # 定义脚本路径
    # script_name = "媒体资源替换及随机资源新增打包版.py"

    # 检查输出路径是否存在，不存在则创建
    if not os.path.exists(output_path):
        os.makedirs(output_path)

    # 构建 PyInstaller 命令
    command = [
        "pyinstaller",
        "--onedir",
        "--add-data", "comlib:dest",
        "--distpath", output_path,
        "--name", output_name,
        script_name
    ]

    try:
        # 执行命令
        subprocess.run(command, check=True)
        print(script_name + "打包成功！")
    except subprocess.CalledProcessError as e:
        print(script_name + f"打包失败：{e}")
    except Exception as e:
        print(script_name + f"发生错误：{e}")


gen_dt_id_path = r"F:\pppppy\SP\func\ID表生成打包版.py"
check_name_path = r''


def package_script_add_other_script(script_name, output_name, other_script):
    # 定义脚本路径
    # script_name = "媒体资源替换及随机资源新增打包版.py"

    # 检查输出路径是否存在，不存在则创建
    if not os.path.exists(output_path):
        os.makedirs(output_path)

    # 构建 PyInstaller 命令
    command = [
        "pyinstaller",
        "--onedir",
        "--add-data", "comlib:dest",
        "--distpath", output_path,
        "--name", output_name,
        "--hidden-import", other_script,  # 隐式导入
        script_name
    ]

    try:
        # 执行命令
        subprocess.run(command, check=True)
        print(script_name + "打包成功！")
    except subprocess.CalledProcessError as e:
        print(script_name + f"打包失败：{e}")
    except Exception as e:
        print(script_name + f"发生错误：{e}")


def package_script_add_other_script_2(script_name, output_name, other_script, other_script2):
    # 定义脚本路径
    # script_name = "媒体资源替换及随机资源新增打包版.py"

    # 检查输出路径是否存在，不存在则创建
    if not os.path.exists(output_path):
        os.makedirs(output_path)

    # 构建 PyInstaller 命令
    command = [
        "pyinstaller",
        "--onedir",
        "--add-data", "comlib:dest",
        "--distpath", output_path,
        "--name", output_name,
        "--hidden-import", other_script,
        "--hidden-import", other_script2,
        script_name
    ]

    try:
        # 执行命令
        subprocess.run(command, check=True)
        print(script_name + "打包成功！")
    except subprocess.CalledProcessError as e:
        print(script_name + f"打包失败：{e}")
    except Exception as e:
        print(script_name + f"发生错误：{e}")


if __name__ == "__main__":
    # package_script("媒体资源替换及随机资源新增打包版.py", "媒体资源替换及随机资源新增")
    # package_script("ES导入自动化打包版.py", "ES_Import")
    # package_script("ID表生成打包版v2.py", "ID表生成")
    # Switch和State有引用ID表生成的exe
    # package_script("State和Switch创建打包版.py", "State和Switch创建")
    # package_script("Wwise属性配置打包版.py", "Wwise属性配置")
    # package_script("GenSoundbank打包版.py", "GenSoundbank")
    # package_script("Amb自动化打包版.py", "全局环境音生成")
    # package_script("Mus自动化打包版.py", "Mus自动化")
    # package_script("Bus自动化打包版.py", "Bus自动化")
    package_script_add_other_script_2("媒体资源占位生成打包版.py", "媒体资源占位生成", "命名规范检查打包版.py",
                                      "媒体资源从表导入打包版.py")
    # package_script("音效语音全流程自动化打包版.py", "音效语音全流程自动化")
    # package_script("AI语音自动化打包版v2.py", "AILanguage_Import")
    copy_directories(output_path, r"S:\chen.gong_DCC_Audio\Audio\Tool\Auto_Sound")
    shutil.rmtree(output_path)
    os.mkdir(output_path)
