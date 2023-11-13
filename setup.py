import os
import shutil
import xml.etree.ElementTree as et
import re


def get_addon_info(path: str):
    print('获取加载项名字和版本号')
    with open(path, encoding='utf-8') as f:
        text = f.read()
    name = re.search(r"name: '([\s\S]+?)'", text)
    version = re.search(r"version: '([\s\S]+?)'", text)
    if name and version:
        return name.group(1), version.group(1)
    else:
        raise '获取失败'


def copy_dir(old_dir: str, new_dir: str):
    remove_dir()
    print('安装加载项')
    ignore_patterns = ['.git', 'README.md']
    if os.path.exists('.gitignore'):
        with open('.gitignore', encoding='utf-8') as f:
            ignore_patterns += f.read().split('\n')
    shutil.copytree(
        old_dir,
        new_dir,
        ignore=shutil.ignore_patterns(*ignore_patterns),
    )


def add_XML():
    print('注册加载项')
    if not os.path.exists(XML_PATH):
        with open(XML_PATH, 'w') as f:
            f.write('<jsplugins>\n</jsplugins>')
    tree = et.parse(XML_PATH)
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
    tree.write(XML_PATH, encoding='utf-8')


def remove_dir():
    print('卸载旧加载项')
    for name in os.listdir(JSADDON_DIR):
        if (NAME in name) and os.path.isdir(JSADDON_DIR + name):
            shutil.rmtree(JSADDON_DIR + name)
            return
    print('找不到旧加载项，无需卸载')


def remove_XML():
    print('注销加载项')
    if not os.path.exists(XML_PATH):
        print('找不到加载项，无需注销')
        return
    tree = et.parse(XML_PATH)
    root = tree.getroot()
    nodes: list[et.Element] = list(
        filter(lambda node: node.attrib['name'] == NAME, root.findall('jsplugin'))
    )
    if nodes == []:
        print('找不到加载项，无需注销')
        return
    for node in nodes:
        root.remove(node)
    tree.write(XML_PATH, encoding='utf-8')


def install():
    if not os.path.exists(JSADDON_DIR):
        os.makedirs(JSADDON_DIR)
    new_dir_name = NAME + '_' + VERSION
    copy_dir(os.path.dirname(__file__), JSADDON_DIR + new_dir_name)
    add_XML()
    print('安装成功, 当前文件夹可删除')


def uninstall():
    remove_dir()
    remove_XML()
    print('卸载成功, 当前文件夹可删除')


if __name__ == '__main__':
    JSADDON_DIR = os.environ['APPDATA'] + '\\kingsoft\\wps\\jsaddons\\'
    XML_PATH = JSADDON_DIR + 'publish.xml'
    try:
        NAME, VERSION = get_addon_info('config.js')
        {
            '0': uninstall,
            '1': install,
        }[input('输入数字(1:安装,0:卸载)')]()
    except Exception as e:
        print('错误', e)
    input('按Enter键退出')
