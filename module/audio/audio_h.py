from pydub import AudioSegment
import numpy as np


# 检查是否为静音片段，若是则返回True
def is_silent(audio_path, silence_threshold=-50.0):
    """
    检查音频片段是否为静音。

    :param audio_segment: pydub.AudioSegment, 音频片段
    :param silence_threshold: float, 静音的分贝阈值
    :return: bool, 如果音频为静音，则返回True
    """

    # 加载音频文件
    audio_segment = AudioSegment.from_wav(audio_path)
    audio_chunks = np.array(audio_segment.get_array_of_samples())
    max_amplitude = np.max(np.abs(audio_chunks))

    # 如果最大振幅为零，则音频为完全静音
    if max_amplitude == 0:
        return True

    # 将最大振幅转换为分贝
    max_db = 20 * np.log10(max_amplitude / float(audio_segment.max_possible_amplitude))

    return max_db < silence_threshold
