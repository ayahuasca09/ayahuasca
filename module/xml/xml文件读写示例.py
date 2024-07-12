# xml读取库
import xml.etree.ElementTree as et

# xml写入库
from xml.dom.minidom import Document

"""**************读xml**************"""
# 打开xml文件
tree = et.ElementTree(file=r"F:\pppppy\SP\module\xml\ExternalSources.xml")
# 获取根节点
root = tree.getroot()
# print(root)
# <Element 'ExternalSourcesList' at 0x000001A2A100AFC0>

# 获取根节点的标签
"""print(root.tag)
# ExternalSourcesList"""

# 获取根节点的属性
"""print(root.attrib)
# {'SchemaVersion': '1', 'Root': 'New_Media'}"""

# 循环遍历子节点
"""for root_children in root:
    print(root_children.tag, root_children.attrib)
    # Source {'Path': 'VO_Lorin_01.wav'}
    # Source {'Path': 'VO_Lorin_02.wav'}
    # Source {'Path': 'VO_Lorin_03.wav'}
    # Source {'Path': 'VO_Lorin_04.wav'}
    # 若还要读里面的子元素，则再作循环即可"""

# 输出某个子节点
"""print(root[2].attrib)
# {'Path': 'VO_Lorin_03.wav'}"""

# 输出子节点的子节点的文本内容
"""for root_children in root:
    for root_sub in root_children:
        print(root_sub.text)
        # 我是帅哥"""

# 索引输出子节点的子节点的文本内容
"""print(root[2][0].text)
# 我是帅哥"""

# 通过标签访问（默认访问第一个）
"""print(root.find("Source").attrib)
# {'Path': 'VO_Lorin_01.wav'}"""

# 通过标签访问所有
"""source_list = root.findall("Source")
for source in source_list:
    print(source.attrib)
    # {'Path': 'VO_Lorin_01.wav'}
    # {'Path': 'VO_Lorin_02.wav'}
    # {'Path': 'VO_Lorin_03.wav'}"""

# 标签的标签访问
"""print(root.find("Source").find("name").text)
# 我是大佬"""

# 通过标签索引的访问
"""print(root.findall("Source")[0].attrib)
# {'Path': 'VO_Lorin_01.wav'}"""

"""**************写xml**************"""
# 创建一个对象
doc = Document()

# 创建根节点
root = doc.createElement("root")
# 添加载Document对象上
doc.appendChild(root)

# 创建子节点
head = doc.createElement("head")
root.appendChild(head)

# 预览xml
"""print(doc.toxml())
# <?xml version="1.0" ?><root><head/></root>"""

# 创建文本内容
"""text1 = doc.createTextNode("1")
code = doc.createElement("code")
code.appendChild(text1)
head.appendChild(code)
print(doc.toxml())"""

# xml文件写入
"""with open('test.xml', 'w+') as f:
    # 按照格式写入
    f.write(doc.toprettyxml(encoding="UTF-8").decode("UTF-8"))
    f.close()"""
