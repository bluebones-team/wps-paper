# 介绍
这是一些使用JavaScript开发的WPS加载项，主要用于**搓论文**  
每个文件夹表示一个加载项  
详细内容见各文件夹中的`README.md`

# 本地安装
1. 按住Ctrl点击[**下载链接**](https://cubxx.github.io/wps-addon/index.html)
2. 下载后会得到一个资源包
3. 解压到一个你喜欢的位置就行
4. 如果是更新，则需要覆盖原先的文件
5. 首次安装需要根据[LookMe.md](./LookMe.md)进行操作
6. .md文件可以用记事本打开

# 在线安装
> 因为国内经常无法连接github，所以建议[**本地安装**](#本地安装)
1. 在文件夹`%APPDATA%\kingsoft\wps\jsaddons`下创建`publish.txt`
2. 写入以下内容
```xml
    <jsplugins>
        <jspluginonline name="汇报" url="https://cubxx.github.io/wps-addon/汇报/wps-addon-build/" type="wpp"/>
        <jspluginonline name="论文" url="https://cubxx.github.io/wps-addon/论文/wps-addon-build/" type="wps"/>
    </jsplugins>
```
3. `publish.txt`文件名改成`publish.xml`