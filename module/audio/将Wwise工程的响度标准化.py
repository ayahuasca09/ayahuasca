from pydub import AudioSegment
from pydub.effects import normalize
import os
from pprint import pprint

wwise_media_path = r'S:\chen.gong_DCC_Audio\Audio\SilverPalace_WwiseProject\Originals'


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


"""遍历文件目录获取文件名称和路径"""


def get_type_file_name_and_path(file_type, dir_path):
    file_dict = {}
    file_list = []
    # 遍历文件夹下的所有子文件
    # 绝对路径，子文件夹，文件名
    for root, dirs, files in os.walk(dir_path):
        # {'name': ['1111.akd', 'Creature Growls 4.akd', 'Sonic Salute_005_2_20.akd'],
        #  'path': 'S:\\chen.gong_DCC_Audio\\Audio\\SilverPalace_WwiseProject\\Originals\\SFX'}
        for file in files:
            if file_type in file:
                new_dict = {}
                file_dict[file] = os.path.join(
                    root, file)
                new_dict[file] = os.path.join(
                    root, file)
                file_list.append(new_dict)

    return file_dict, file_list


# 使用示例
# input_path = "input.wav"
# output_path = "output_audio_normalized.wav"
# target_dBFS = -30.0  # 目标响度值，单位为 dBFS
#
# normalize_audio_to_target(input_path, output_path, target_dBFS)
file_dict, file_list = get_type_file_name_and_path('.wav', wwise_media_path)
# pprint(file_dict)

target_dBFS = -40
for key in file_dict:
    # pprint(file_dict[key])
    normalize_audio_to_target(file_dict[key], file_dict[key], target_dBFS)
