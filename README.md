- [介绍](#介绍)
- [安装](#安装)
  - [本地安装](#本地安装)
  - [在线安装](#在线安装)
  - [启用加载项（可选）](#启用加载项可选)


# 介绍
这是一些使用JavaScript开发的WPS加载项，主要用于**搓论文**

每个文件夹表示一个加载项

详细内容见各文件夹中的`README.md`

# 安装
**视频教程** `https://www.bilibili.com/list/525570753?sid=3253331`

## 本地安装
1. 从右边的Releases界面下载安装包
2. 解压到一个你喜欢的位置就行
3. 首次安装需要根据安装包内的**安装教程**进行操作

## 在线安装
> 因为国内经常无法连接github，所以建议[**本地安装**](#本地安装)

1. 在文件夹`%APPDATA%\kingsoft\wps\jsaddons`下创建`publish.xml`
2. 写入以下内容

```xml
<jsplugins>
    <jspluginonline name="论文" url="https://cubxx.github.io/wpsAcademic/论文/" type="wps"/>
</jsplugins>
```

## 启用加载项（可选）
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
