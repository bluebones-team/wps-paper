from genericpath import isdir
import os
import shutil
import logging
import xml.etree.ElementTree as et
import re


def get_addon_info(path: str):
    logging.info('获取加载项名字和版本号')
    with open(path, encoding='utf-8') as f:
        text = f.read()
    name = re.search(r"name: '([\s\S]+?)'", text)
    version = re.search(r"version: '([\s\S]+?)'", text)
    if name and version:
        return name.group(1), version.group(1)
    else:
        logging.error('获取失败')
        quit()


def remove_dir(parent_dir: str):
    '''移除包含 NAME 的文件夹'''
    for name in os.listdir(parent_dir):
        if (NAME in name) and os.path.isdir(parent_dir + os.path.sep + name):
            shutil.rmtree(parent_dir + os.path.sep + name)
            logging.info('文件夹已移除')
            return
    logging.info('找不到相应文件夹，无需移除')


def copy_dir(old_dir: str, new_dir: str):
    parent_dir = os.path.dirname(new_dir)
    remove_dir(parent_dir)
    logging.info('复制文件夹')
    ignore_patterns = ['.git', THIS_NAME]
    if os.path.exists('.gitignore'):
        with open('.gitignore', encoding='utf-8') as f:
            ignore_patterns += f.read().split('\n')
    shutil.copytree(
        old_dir,
        new_dir,
        ignore=shutil.ignore_patterns(*ignore_patterns),
    )


def add_XML(XML_path: str):
    logging.info('注册加载项')
    if not os.path.exists(XML_path):
        with open(XML_path, 'w') as f:
            f.write('<jsplugins>\n</jsplugins>')
    tree = et.parse(XML_path)
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
    tree.write(XML_path, encoding='utf-8')


def remove_XML(XML_path: str):
    logging.info('注销加载项')
    if not os.path.exists(XML_path):
        logging.info('找不到注册文件，无需卸载')
        return
    tree = et.parse(XML_path)
    root = tree.getroot()
    nodes: list[et.Element] = list(
        filter(lambda node: node.attrib['name'] == NAME, root.findall('jsplugin'))
    )
    if nodes != []:
        for node in nodes:
            root.remove(node)
    tree.write(XML_path, encoding='utf-8')


def install():
    if not os.path.exists(JSADDON_DIR):
        os.makedirs(JSADDON_DIR)
    new_dir_name = NAME + '_' + VERSION
    copy_dir(os.path.dirname(__file__), JSADDON_DIR + new_dir_name)
    add_XML(XML_PATH)
    logging.info('安装成功, 当前文件夹可删除')


def uninstall():
    remove_dir(JSADDON_DIR)
    remove_XML(XML_PATH)
    logging.info('卸载成功, 当前文件夹可删除')


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
        XML_PATH = JSADDON_DIR + 'publish.xml'
        NAME, VERSION = get_addon_info('config.js')

        {
            '0': uninstall,
            '1': install,
        }[input('please input number 1 or 0\n(1:install, 0:uninstall)\n')]()
    except Exception as e:
        logging.error(e)
