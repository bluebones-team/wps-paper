# 介绍
这是一些使用JavaScript开发的WPS加载项，主要用于**搓论文**  
这些加载项需要在使用时**保持联网**  
每个文件夹表示一个加载项  
详细内容见文件夹中的`README.md`

# 安装
## 填写链接
进入 `WPS安装路径\11.1.0.版本号\office6\cfgs`  
找到 `oem.ini` 文件  
用记事本打开该文件，添加以下文本  
```
[Support]
JsApiPlugin=true
JsApiShowWebDebugger=false

[Server]
JSPluginsServer=https://cubxx.github.io/wps-addon/jsplugins.xml
```
> 文件夹或文件，没找到就自己新建

## 启动WPS  
在WPS中选择 `选项 > 自定义功能区` 勾选相应加载项
> 如果找不到相应加载项，就等几分钟，程序在加载

## 检查网络
可以通过点击[**这里**](https://cubxx.github.io/wps-addon/)判断能否连接至Github服务器  
不行的话就**科学上网**
