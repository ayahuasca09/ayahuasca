import xml.etree.ElementTree as ET
import xml.dom.minidom as minidom
from lxml import etree

"""xml元素及数据插入"""


def create_element_from_xml_string(xml_string):
    """
    根据传入的 XML 字符串创建一个新的元素。

    :param xml_string: 包含结构信息的 XML 字符串。
    :return: 新创建的元素。
    """
    # 解析XML字符串，并返回根元素
    new_element = ET.fromstring(xml_string)
    return new_element


def elements_are_equal(e1, e2):
    """
    比较两个XML元素及其子元素是否相等。

    :param e1: 第一个元素。
    :param e2: 第二个元素。
    :return: 如果元素相等则返回True，否则返回False。
    """
    if e1.tag != e2.tag or e1.attrib != e2.attrib or (e1.text or '').strip() != (e2.text or '').strip():
        return False
    return all(elements_are_equal(c1, c2) for c1, c2 in zip(e1, e2))


def insert_into_xml(file_path, xpath, new_element):
    """
    在指定的 XML 文档中的特定位置插入新的元素。

    :param file_path: XML 文件路径。
    :param xpath: 要插入新元素的 XPath 表达式。
    :param new_element: 要插入的新的元素（xml.etree.ElementTree.Element 对象）。
    """
    tree = ET.parse(file_path)  # 解析XML文档，返回一个ElementTree对象
    root = tree.getroot()  # 获取XML文档的根元素

    # 根据XPath表达式查找要插入的父元素
    parent_element = root.find(xpath)
    if parent_element is None:
        print(f"无法找到指定的路径: {xpath}")
        return

    # 检查是否已有相同的元素
    for existing_element in parent_element.findall(new_element.tag):
        if elements_are_equal(existing_element, new_element):
            # print("相同的元素已经存在于XML中。")
            return

    # 将新元素插入到找到的父元素中
    parent_element.append(new_element)

    # 美化并保存修改后的XML文档
    save_pretty_xml(tree, file_path)
    print("XML 文件已更新。")


def save_pretty_xml(tree, file_path):
    """
    美化并保存XML文档。

    :param tree: ElementTree对象。
    :param file_path: 要保存的文件路径。
    """
    xml_str = ET.tostring(tree.getroot(), 'utf-8')  # 将XML元素转换为字节字符串
    parsed_str = minidom.parseString(xml_str)
    pretty_xml_str = "\n".join([line for line in parsed_str.toprettyxml(indent="  ").split("\n") if line.strip()])

    with open(file_path, 'w', encoding='utf-8') as f:
        f.write(pretty_xml_str)  # 将美化后的XML字符串写回文件


# 示例XML数据，表示要插入的条目的结构和属性
# xml_data = '''
# <Entry>
#     <Path Name="AbaAbaB" ID="{aaaaaaaaaaaaa}" />
#     <AudioNodeInfo Name="Mus_Loading" ID="{AD0A4F68-B1E4-457C-BB0E-94639CE826A1}"
#                    WorkUnitID="{8B94E63D-66D7-49B0-A134-D2A9CDBE7006}"
#                    Platform="Linked" />
# </Entry>
# '''
#
# # 创建新的XML元素
# new_entry = create_element_from_xml_string(xml_data)

# 将新元素插入到指定的XML文件中
# insert_into_xml('Mus_Test.xml', './/EntryList', new_entry)


"""xml元素及数据删除"""


def remove_element_from_xml(file_path, xpath):
    tree = etree.parse(file_path)
    root = tree.getroot()

    # 使用 lxml 的 XPath 功能
    elements_to_remove = root.xpath(xpath)
    if not elements_to_remove:
        print(f"没有找到与 XPath 匹配的元素: {xpath}")
        return

    for element_to_remove in elements_to_remove:
        parent = element_to_remove.getparent()
        if parent is not None:
            parent.remove(element_to_remove)

    # 写回修改后的 XML 文件
    with open(file_path, 'wb') as f:
        f.write(etree.tostring(tree, pretty_print=True, encoding='utf-8'))

    print("元素已删除并更新 XML 文件。")


# remove_element_from_xml('Mus_Test.xml', './/Entry[Path[@Name="New"]]')

"""xml数据读取和输出"""


def print_attributes(element, indent=0):
    """
        打印XML元素的属性。

        :param element: XML元素。
        :param indent: 缩进级别。
        """
    prefix = ' ' * (indent * 2)
    for attr, value in element.attrib.items():
        print(f"{prefix}  @{attr}: {value}")


def parse_element(element, indent=0, is_show_property=True):
    """
        递归解析XML元素及其子元素。

        :param element: 当前XML元素。
        :param indent: 缩进级别。
        :is_show_property: 是否显示标签属性。
        """
    prefix = ' ' * (indent * 2)
    for child in element:
        print(f"{prefix}{child.tag}")
        if is_show_property:
            print_attributes(child, indent)
        parse_element(child, indent + 1)


def parse_xml(file_path):
    """
        解析XML文件并格式化输出其结构。

        :param file_path: XML文件的路径。
        """
    # 从文件读取XML内容
    tree = ET.parse(file_path)
    root = tree.getroot()

    # 打印根元素及其属性
    print(f"Root element: {root.tag}")
    print_attributes(root)

    # 递归解析XML节点
    parse_element(root)


# 递归解析XML元素及其子元素并输出
# parse_xml('Mus_Test.xml')


"""xml已有属性更改"""


def modify_xml_attribute(file_path, element_tag, attribute_name, new_value, output_file_path):
    """
        修改XML文件中指定元素的属性值，适用于应用所有标签

        :param file_path: 输入XML文件的路径。
        :param element_tag: 要修改的元素的标签。
        :param attribute_name: 要修改的属性名称。
        :param new_value: 新的属性值。
        :param output_file_path: 保存修改后XML文件的路径。
        """
    # 从文件读取XML内容
    tree = ET.parse(file_path)
    root = tree.getroot()

    # 查找指定的元素
    for element in root.iter(element_tag):
        if attribute_name in element.attrib:
            print(f"修改前: {element_tag} - @{attribute_name}: {element.attrib[attribute_name]}")
            # 修改属性值
            element.set(attribute_name, new_value)
            print(f"修改后: {element_tag} - @{attribute_name}: {element.attrib[attribute_name]}")

    # 将修改后的XML保存到新文件
    tree.write(output_file_path, encoding='utf-8', xml_declaration=True)


# 设置元素属性值
# modify_xml_attribute('Mus_Test.xml', 'MusicSwitchContainer', 'Name',
#                      'New_Global_Name', 'Modified_Mus_Test.xml')


def update_xml_attribute(file_path, tag_name, attribute_name, attribute_value, target_attribute, target_value):
    """
    更新XML文件中指定标签的属性值。

    :param file_path: XML文件路径
    :param tag_name: 要查找的标签名称
    :param attribute_name: 用于匹配的属性名称
    :param attribute_value: 用于匹配的属性值
    :param target_attribute: 要更新的目标属性名称
    :param target_value: 要设置的新属性值
    """
    tree = ET.parse(file_path)
    root = tree.getroot()

    # 查找符合条件的标签
    for elem in root.findall(f".//{tag_name}[@{attribute_name}='{attribute_value}']"):
        elem.set(target_attribute, target_value)  # 修改属性值

    # 保存修改后的XML文件
    tree.write(file_path, encoding='utf-8', xml_declaration=True)

# 使用示例
# update_xml_attribute(
#     'music.xml',
#     'MusicSwitchContainer',
#     'Name',
#     'Mus_Map_A02',
#     'ShortID',
#     '987654321'
# )
