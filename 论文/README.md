# 介绍
这是一个使用JavaScript开发的WPS加载项，其主要功能是协助论文写作  
该加载项部署在服务器上，使用时需要安装相应的xml文件，见 **使用教程**  
由于加载项文件在服务端，用户无需进行手动更新，但需要在使用时保持联网  
    
参考指南 [WPS开发者文档](https://qn.cache.wpscdn.cn/encs/doc/office_v19/webhelpframe.htm)
# 使用教程
## 进入[部署网页](http://47.113.221.157:81/wps-addon/publish.html)
状态正常的话，安装加载项  
安装成功后会在本地生成xml文件，WPS通过该文件连接服务器，实现加载项功能
>安装时保持WPS关闭状态
## 检验是否安装成功
`Win`+`E`打开文件资源管理器  
在地址栏输入`%appdata%/kingsoft/wps/jsaddons`  
检查该目录下是否存在 `publish.xml`
## 打开WPS，载入xml
随便打开一个文档，`开发工具 > 加载项 > 添加`  
转到上述 `publish.xml` 文件路径，并选择该文件  
主选项卡面板将出现 `论文` 加载项
# 问题解决
## 部署网页中没有安装选项
按 `F12`，控制台 Console 报错如下:
> Access to XMLHttpRequest at `http://127.0.0.1:58890/version` from origin `http://47.113.221.157:81` has been blocked by CORS policy: The request client is not a secure context and the resource is in more-private address space `local`. 

**[WPS官方解决文档](https://www.kdocs.cn/l/cv7pyp6sqOFC)**  
进入 <chrome://flags/#block-insecure-private-network-requests>，将 `block-insecure-private-network-requests` 设置为 `Disabled`
>**注意**:  
此操作将关闭浏览器限制非本地服务器访问[本地服务器](http://localhost)的功能  
如果你开启了IIS功能，在访问需要请求本地服务器的网页中，可能存在一定风险  
可以通过对 `控制面板 > 程序 > 启用或关闭Windows功能 > Internet Information Services` 勾选或取消勾选，决定是否开启IIS功能
## 载入xml后没有出现加载项
**没有勾选该功能区**  
`文件 > 选项 > 自定义功能区` 勾选 `论文`  
**网络问题**  
我也不会
# 功能原理
## 参考文献格式化

