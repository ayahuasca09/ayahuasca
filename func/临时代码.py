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