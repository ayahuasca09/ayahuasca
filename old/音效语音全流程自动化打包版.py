import os.path
import comlib.config as config
import comlib.exe_h as exe_h

auto_sound_path = config.auto_sound_path


def exe_path_run(dir_path, exe_name):
    gen_dt_id_path = os.path.join(auto_sound_path, dir_path, exe_name)
    exe_h.run_exe(gen_dt_id_path)


exe_path_run("媒体资源替换及随机资源新增", "媒体资源替换及随机资源新增.exe")
exe_path_run("Wwise属性配置", "Wwise属性配置.exe")
exe_h.run_ue_reconcile()
# id表生成
exe_path_run("ID表生成", "ID表生成.exe")

os.system("pause")
