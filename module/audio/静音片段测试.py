from pydub import AudioSegment
import numpy as np


def is_silent(audio_segment, silence_threshold=-50.0):
    """
    检查音频片段是否为静音。

    :param audio_segment: pydub.AudioSegment, 音频片段
    :param silence_threshold: float, 静音的分贝阈值
    :return: bool, 如果音频为静音，则返回True
    """
    audio_chunks = np.array(audio_segment.get_array_of_samples())
    max_amplitude = np.max(np.abs(audio_chunks))

    # 如果最大振幅为零，则音频为完全静音
    if max_amplitude == 0:
        return True

    # 将最大振幅转换为分贝
    max_db = 20 * np.log10(max_amplitude / float(audio_segment.max_possible_amplitude))

    return max_db < silence_threshold


def main():
    # 加载音频文件
    audio = AudioSegment.from_wav("Media_Temp.wav")

    # 检查是否为静音
    if is_silent(audio):
        print("The audio is silent.")
    else:
        print("The audio is not silent.")


if __name__ == "__main__":
    main()