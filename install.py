"""
手动安装教程
1. 在文件管理器的地址栏输入 %AppData%\kingsoft\wps\jsaddons, 按回车, 找不到文件夹就自己新建
2. 把下载的安装包解压至该文件夹下
3. 更改文件名为"加载项名_版本号"的格式, 如: "wps-paper-23.9"更改为"wps-paper_23.9"
4. 打开当前文件夹下的"jsplugins.xml"文件, 如果没有就自己新建一个
5. 在文件内添加加载项信息:
文件内容结构如下
    <jsplugins>
        <!-- 你可以在这里添加新的加载项信息 -->
    </jsplugins>
需要添加加载项信息为
    <jsplugin name="在此填写加载项名" version="在此填写版本号" type="wps" url="https://github.com/Cubxx/wps-paper"/>
"""

import os
import shutil
import json
import logging
import xml.etree.ElementTree as et


def get_config(path):
    logging.info('获取配置文件')
    with open(path, encoding='utf-8') as f:
        json_str = f.read().replace('const config = ', '')
    return json.loads(json_str)


def copy_dir_to(new_parent_dir, config):
    logging.info('复制当前文件夹')
    this_path = os.path.abspath(__file__)
    this_dir = os.path.dirname(this_path)
    new_dir_name = config['name'] + '_' + config['version']
    new_dir = new_parent_dir + os.path.sep + new_dir_name
    for dir_name in os.listdir(new_parent_dir):
        if config['name'] in dir_name:
            shutil.rmtree(new_parent_dir + os.path.sep + dir_name)
    shutil.copytree(
        this_dir, new_dir, ignore=shutil.ignore_patterns('.git', 'web.config', LOG_NAME)
    )


def add_XML(path, config):
    logging.info('添加加载项信息')
    if not os.path.exists(path):
        with open(path, 'w') as f:
            f.write('<jsplugins>\n</jsplugins>')
    else:
        tree = et.parse(path)
        root = tree.getroot()
        nodes = list(
            filter(lambda node: node.attrib['name'] == config['name'], root.findall('jsplugin'))
        )
        if nodes == []:
            root.append(
                et.Element(
                    'jsplugin',
                    {
                        'name': config['name'],
                        'type': 'wps',
                        'url': 'https://github.com/Cubxx/wps-paper',
                        'version': config['version'],
                    },
                )
            )
        else:
            nodes[0].attrib['version'] = config['version']
        tree.write(path, encoding='utf-8')


def main():
    JSADDON_DIR = os.environ['APPDATA'] + '\\kingsoft\\wps\\jsaddons'
    XML_PATH = JSADDON_DIR + os.path.sep + 'jsplugins.xml'
    CONFIG = get_config('config.js')
    if not os.path.exists(JSADDON_DIR):
        os.makedirs(JSADDON_DIR)
    add_XML(XML_PATH, CONFIG)
    copy_dir_to(JSADDON_DIR, CONFIG)


if __name__ == '__main__':
    LOG_NAME = 'install.log'
    logging.basicConfig(
        filename=LOG_NAME,
        filemode='w',
        format='%(asctime)s.%(msecs)03d %(levelname)s %(module)s: %(message)s',
        datefmt='%H:%M:%S',
        level=logging.INFO,
    )
    try:
        main()
        logging.info('安装成功, 当前文件夹可删除')
    except Exception as e:
        logging.error(e)
