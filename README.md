- [1. 介绍](#1-介绍)
- [2. 安装](#2-安装)
  - [2.1. 本地安装](#21-本地安装)
  - [2.2. 在线安装](#22-在线安装)
  - [2.3. 启用加载项（可选）](#23-启用加载项可选)


# 1. 介绍
这是一个主要用于**搓论文**的WPS加载项项目

每个文件夹表示一个加载项，文件夹内有说明文档

# 2. 安装
**视频教程** `https://www.bilibili.com/list/525570753?sid=3253331`

## 2.1. 本地安装
1. 从右边的 `Releases` 界面下载安装包
2. 解压到一个你喜欢的位置就行，一定要解压
3. 首次安装需要根据安装包内的**安装教程**进行操作

## 2.2. 在线安装
> 因为国内经常无法连接github，所以建议[**本地安装**](#本地安装)

1. 在文件夹 `%APPDATA%\kingsoft\wps\jsaddons` 下创建 `publish.xml`
2. 写入以下内容

```xml
<jsplugins>
    <jspluginonline name="论文" url="https://cubxx.github.io/wpsAcademic/论文/" type="wps"/>
</jsplugins>
```

## 2.3. 启用加载项（可选）
打开 WPS 安装目录，进入 wps.exe 程序同级目录

在`cfgs`文件夹下新建或打开`oem.ini`文件

在`oem.ini`文件最后添加如下配置

```ini
[Support]
##启用加载项
JsApiPlugin = true
##启用网页调试
JsApiShowWebDebugger = false
```
