import os
import shutil
import json
import logging
import xml.etree.ElementTree as et
import re


def get_addon_info(path: str):
    logging.info('获取配置信息')
    with open(path, encoding='utf-8') as f:
        text = f.read()
    name = re.search(r"name: '([\s\S]+?)'", text)
    version = re.search(r"version: '([\s\S]+?)'", text)
    if name and version:
        return name.group(1), version.group(1)
    else:
        raise '获取配置信息失败'


def copy_dir(old_dir: str, new_dir: str):
    logging.info('复制当前文件夹')
    parent_dir = os.path.dirname(new_dir)
    for root_path, dir_names, file_names in os.walk(parent_dir):
        for sub_dir_name in dir_names:
            if NAME in sub_dir_name:
                shutil.rmtree(parent_dir + os.path.sep + sub_dir_name)  # 删除文件夹
    shutil.copytree(
        old_dir,
        new_dir,
        ignore=shutil.ignore_patterns('.git', '.gitignore', '*.config', '*.log', THIS_NAME),
    )


def add_XML(path: str):
    logging.info('添加加载项信息')
    if not os.path.exists(path):
        with open(path, 'w') as f:
            f.write('<jsplugins>\n</jsplugins>')
    tree = et.parse(path)
    root = tree.getroot()
    nodes: list[et.Element] = list(
        filter(lambda node: node.attrib['name'] == NAME, root.findall('jsplugin'))
    )
    if nodes == []:
        root.append(
            et.Element(
                'jsplugin',
                {
                    'name': NAME,
                    'type': 'wps',
                    'url': 'https://github.com/Cubxx/wps-paper',
                    'version': VERSION,
                },
            )
        )
    else:
        nodes[0].attrib['version'] = VERSION
    tree.write(path, encoding='utf-8')


if __name__ == '__main__':
    THIS_NAME = os.path.basename(__file__)
    logging.basicConfig(
        filename=THIS_NAME.replace('.py', '.log'),
        filemode='w',
        format='%(asctime)s.%(msecs)03d %(levelname)s %(module)s: %(message)s',
        datefmt='%H:%M:%S',
        level=logging.INFO,
    )
    try:
        JSADDON_DIR = os.environ['APPDATA'] + '\\kingsoft\\wps\\jsaddons\\'
        NAME, VERSION = get_addon_info('config.js')
        if not os.path.exists(JSADDON_DIR):
            os.makedirs(JSADDON_DIR)
        new_dir_name = NAME + '_' + VERSION
        copy_dir(os.path.dirname(__file__), JSADDON_DIR + new_dir_name)
        add_XML(JSADDON_DIR + 'publish.xml')
        logging.info('安装成功, 当前文件夹可删除')
    except Exception as e:
        logging.error(e)
