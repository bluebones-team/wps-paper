# 一般步骤
1. 保证加载项文件夹、*jsplugins.xml*、*deploying.bat*、*addonServer.exe*在同一文件夹下
2. 双击启动*addonServer.exe*，当前文件夹会出现*record.txt*文件
3. 打开*record.txt*，查看是否有报错，并依照提示操作，然后关闭*record.txt*，**再次启动***addonServer.exe*，直到成功部署
4. 双击运行*deploying.bat*，可能会有弹窗，按F
5. 打开WPS，随便开个文档，查看自定义功能区内是否有相应加载项
6. 一次部署成功，以后就不需要再启动*addonServer.exe*了，因为*addonServer.exe*已经加入了[开机自启动](#deployingbat)

# 遇到问题
1. 进入*record.txt*文件中的网址，如果有数据出现则部署成功，打不开则部署失败，此时建议**再次启动***addonServer.exe*
2. 检查目录`%APPDATA%\kingsoft\wps\jsaddons`下是否只有publish.xml文件，多余文件请删除
3. 保持WPS打开，加载项可能需要10+秒加载

# 文件说明
## *addonServer.exe*
1. 打开任务管理器，按名称排序，可以在后台进程中找到*addonServer.exe*。运行1次*addonServer.exe*将创建2个进程，需要**再次启动**时可以关闭它们
2. 占用内存很低，适合加入开机自启动，省得每次都需要启动*addonServer.exe*
## *jsplugins.xml*
1. 加载项链接文件，部署失败时可能需要手动修改**端口号**
2. 例如：`http://127.0.0.1:6304` 中，`6304` 是**端口号**
## *deploying.bat*
1. 将*addonServer.exe*快捷方式添加至`shell:startup`文件夹中，实现开机自启
2. 将*jsplugins.xml*文件复制到`%APPDATA%\kingsoft\wps\jsaddons`，并重命名为publish.xml

# 实现原理
1. 打开WPS程序中的文档时，WPS会查看`%APPDATA%\kingsoft\wps\jsaddons`文件夹下是否有`jsplugins.xml`或`publish.xml`
2. 如果有，WPS将根据该文件内的url（网址）获取加载项资源，加载速度取决于WPS访问文件内url的速度；如果没有，WPS将不会载入加载项
3. 所以*addonServer.exe*根据*jsplugins.xml*中的url和name属性构建本地文件服务器，将相应加载项文件夹部署在本地端口上，实现通过网址访问本地加载项文件
4. 网址 - 文件夹 的对应关系写在*record.txt*中，WPS将通过这些网址访问对应本地文件夹（加载项文件夹）
5. 因为是访问本地文件，所以速度非常快，也不占用流量，就是每次更新都需要手动进行