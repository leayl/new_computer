"""
分两部分，两种方式进行
1. etree.ElementTree, 个人觉得更适合在对文档层级控制要求不那么高时使用
2. dom.minidom,
"""
import os
import xml.etree.ElementTree as ET
import xml.dom.minidom as minidom


def read_xml_ET(xml_path):
    """
    读取使用xml.etree.ElementTree读取的xml内容
    """
    # 读取文件数据，生成root根元素
    et = ET.parse(xml_path)
    root = et.getroot()
    # 或者直接使用fromstring读取文件数据生成root
    # xml_str_data = ''
    # with open(xml_path, 'r') as f:
    #     xml_str_data = f.read()
    # root = ET.fromstring(xml_str_data)

    # 获取元素标签名
    root_tag = root.tag  # 如<data></data>元素，结果为data

    # 标签可直接迭代
    for element in root:
        attribs = element.attrib  # 获取元素属性字典
        content = element[0].text  # 子级可以嵌套，使用element.text获取文本元素
        for attrib, value in attribs.items():
            print(f"{attrib}: {value}")

    # 递归遍历所有指定标签的子级元素
    for child in root.iter('neighbor'):  # 查找子级中所有的neighbor元素
        print(child.attrib)

    # 在当前元素的直接子元素中查找所有指定标签的子级元素
    for country in root.findall('country'):
        name = country.get("name")

    # 在当前元素的直接子元素中查找第一个指定标签的子级元素
    country1 = root.find("country")
    if country1:
        print(len(country1))


def modify_xml_ET(xml_path):
    """
    修改由xml.etree.ElementTree读取的xml文件
    """
    # 解析文档，获得根元素
    et = ET.parse(xml_path)
    root = et.getroot()

    # 获得要修改的元素
    country1 = root[0]
    # 修改和增加属性
    country1.set("name", f"{country1.attrib['name']}modify")
    country1.set("area", "Aisa")

    # 删除元素
    for country in root.findall("country"):
        rank = int(country.find("rank").text)
        if rank > 50:
            root.remove(country)
    # 新建元素
    new_ele = ET.Element("country")
    new_ele.set("name", "Finland")
    # 在元素内部追加并新建元素
    new_rank = ET.SubElement(new_ele, "rank", {"type": "new"})
    new_rank.text = "10"
    # 在末尾追加子元素
    root.append(new_ele)
    # 在指定位置插入新增元素
    root.insert(1, new_ele)

    # a = ET.SubElement(root, "country", {"name": "Norway"})
    # b = ET.SubElement(a, 'rank', {"type": "Scandinavia"})

    # 保存或另存为
    et.write(xml_path)


def read_xml_minidom(xml_path):
    """
    使用xml.dom.minidom解析并读取xml
    """
    # 解析xml文件并获取根节点
    dom_tree = minidom.parse(xml_path)
    root_node = dom_tree.documentElement

    # 根据标签名获取节点所有符合要求的子节点
    country_nodes = root_node.getElementsByTagName("country")
    for country_node in country_nodes:
        # 获取节点属性值
        name = country_node.getAttribute("name")
        print("name", name)
        # 获取节点的文档节点内容
        rank_node = country_node.getElementsByTagName("rank")[0]
        rank = rank_node.childNodes[0].data
        print(rank_node.tagName, rank)


def modify_xml_minidom(xml_path, modify_xml_path):
    """
    使用xml.dom.minidom操作修改读取的xml文件
    1.创建一个新元素结点createElement()
    2.创建一个文本节点createTextNode()
    3.将文本节点挂载元素结点上
    4.将元素结点挂载到其父元素上
    """
    dom_tree = minidom.parse(xml_path)
    root_node = dom_tree.documentElement

    countries = root_node.getElementsByTagName("country")
    countries[0].setAttribute("area", "Asia")

    # 新建country节点
    new_country_node = dom_tree.createElement("country")
    # 设置节点属性
    new_country_node.setAttribute("name", "Finland")

    new_rank_node = dom_tree.createElement("rank")
    # 新建文本节点
    new_rank_text_node = dom_tree.createTextNode("2")
    # 将文本节点挂载在父节点中
    new_rank_node.appendChild(new_rank_text_node)

    new_neighbor_node = dom_tree.createElement("neighbor")
    new_neighbor_text_node = dom_tree.createTextNode("Norway")
    new_neighbor_node.appendChild(new_neighbor_text_node)

    new_country_node.appendChild(new_rank_node)
    new_country_node.appendChild(new_neighbor_node)

    root_node.appendChild(new_country_node)

    with open(modify_xml_path, 'w') as f:
        # 修改后原有数据格式会变化，每行后多了一个空行
        dom_tree.writexml(f, addindent='', newl='\n', encoding='utf-8')


def create_xml_minidom(xml_path):
    """
    新建全新xml文件
    """
    # 创建文档
    domTree = minidom.Document()
    # 创建根节点
    root_node = domTree.createElement("countries")
    for i in range(5):
        # 创建根节点的子节点
        country = domTree.createElement("country")
        country.setAttribute('name', f"country{i}")

        rank = domTree.createElement('rank')
        rank_text = domTree.createTextNode(f"{i + 1}")
        rank.appendChild(rank_text)

        neighbor = domTree.createElement("neighbor")
        neighbor.setAttribute('area', f'area{i + 1}')
        neighbor_text = domTree.createTextNode(f"邻居{i+1}")
        neighbor.appendChild(neighbor_text)

        country.appendChild(rank)
        country.appendChild(neighbor)
        # 将节点挂载在根节点
        root_node.appendChild(country)
    # 将根节点挂载在文档
    domTree.appendChild(root_node)

    # 新建文件，并将文档保存至文件
    with open(xml_path, 'w') as f:
        domTree.writexml(f, newl='\n', encoding='gbk')


if __name__ == "__main__":
    xml_path = "xml_files/old_xml.xml"
    modify_xml_path = "xml_files/modify_xml.xml"
    new_xml_path = "xml_files/new_xml.xml"
    # read_xml_ET(xml_path)
    # modify_xml_ET(modify_xml_path)
    # read_xml_minidom(xml_path)
    # modify_xml_minidom(xml_path, modify_xml_path)
    create_xml_minidom(new_xml_path)
