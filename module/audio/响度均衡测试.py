from pydub import AudioSegment


def match_target_amplitude(sound, target_dBFS):
    change_in_dBFS = target_dBFS - sound.dBFS
    return sound.apply_gain(change_in_dBFS)


def normalize_audio_to_target(input_file, output_file, target_dBFS):
    # 加载音频文件
    audio = AudioSegment.from_file(input_file)

    # 获取原始音频的属性
    channels = audio.channels
    sample_width = audio.sample_width
    frame_rate = audio.frame_rate
    bitrate = audio.frame_rate * audio.sample_width * 8 * channels  # 位速率

    # 将响度标准化为目标值
    normalized_audio = match_target_amplitude(audio, target_dBFS)

    # 导出标准化后的音频文件，保留原始属性
    normalized_audio.export(
        output_file,
        format="wav",
        bitrate=f"{bitrate}k",
        channels=channels,
        sample_width=sample_width,
        frame_rate=frame_rate
    )


# 使用示例
input_path = "input_audio.wav"
output_path = "output_audio_normalized.wav"
target_dBFS = -20.0  # 目标响度值，单位为 dBFS

normalize_audio_to_target(input_path, output_path, target_dBFS)

### 说明
"""音频属性获取：
`channels`: 获取声道数。
`sample_width`: 获取每个样本的字节数。
`frame_rate`: 获取采样率。
`bitrate`: 计算位速率。对于非 MP3 格式，可以忽略此参数。
导出时保留原始属性：
使用 `export` 方法时，指定 `channels`、`sample_width`、`frame_rate` 等参数以保持原始音频的属性。
对于格式为 MP3 的音频，使用 `bitrate` 参数指定原始位速率。
这个示例确保在标准化响度的同时保留音频的原始技术属性。"""
