import subprocess

# 定义命令和参数
command = r"D:\Wwise2022.1.16.8522\Authoring\x64\Release\bin\WwiseConsole.exe"
args = [
    'convert-external-source',
    r'S:\chen.gong_DCC_Audio\Audio\SilverPalace_WwiseProject\SilverPalace_WwiseProject.wproj',
    '--output', 'Windows', r'S:\Ver_1.0.0\Project\Content\Audio\GeneratedExternalSources\Windows',
    '--output', 'Android', r'S:\Ver_1.0.0\Project\Content\Audio\GeneratedExternalSources\Android',
    '--output', 'iOS', r'S:\Ver_1.0.0\Project\Content\Audio\GeneratedExternalSources\iOS',
    '--source-file', r'S:\chen.gong_DCC_Audio\Audio\Tool\Auto_Sound\ES_Import\ExternalSource.wsources'
]

# 执行命令
try:
    subprocess.run([command] + args, check=True)
    print("Command executed successfully")
except subprocess.CalledProcessError as e:
    print(f"An error occurred: {e}")
