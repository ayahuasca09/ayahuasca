import subprocess

# 定义命令和参数
command = r"D:\Wwise2022.1.16.8522\Authoring\x64\Release\bin\WwiseConsole.exe"
args = [
    'generate-soundbank',
    r'S:\chen.gong_DCC_Audio\Audio\SilverPalace_WwiseProject\SilverPalace_WwiseProject.wproj',
    '--soundbank-path', 'Windows', r'S:\Ver_1.0.0\Project\Content\Audio\GeneratedSoundBanks\Windows',
    '--soundbank-path', 'Android', r'S:\Ver_1.0.0\Project\Content\Audio\GeneratedSoundBanks\Android',
    '--soundbank-path', 'iOS', r'S:\Ver_1.0.0\Project\Content\Audio\GeneratedSoundBanks\iOS',
    '--root-output-path', r'S:\Ver_1.0.0\Project\Content\Audio\GeneratedSoundBanks'
]

# 执行命令
try:
    subprocess.run([command] + args, check=True)
    print("Command executed successfully")
except subprocess.CalledProcessError as e:
    print(f"An error occurred: {e}")
