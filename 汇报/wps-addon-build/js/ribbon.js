function OnAddinLoad(ribbonUI) {
    if (typeof (wps.ribbonUI) != "object") {
        wps.ribbonUI = ribbonUI
    }
    if (typeof (wps.Enum) != "object") { // 如果没有内置枚举值
        wps.Enum = WPS_Enum
    }
    wps.PluginStorage.setItem("EnableFlag", false) //往PluginStorage中设置一个标记，用于控制两个按钮的置灰
    wps.PluginStorage.setItem("ApiEventFlag", false) //往PluginStorage中设置一个标记，用于控制ApiEvent的按钮label
    return true
}
var pre = Application.ActivePresentation,
    sel = Application.Windows.Item(1).Selection,
    _debug = { isDefalutPath: true };

//导入框架
function importFrame() {
    // script
    let text = _debug.isDefalutPath ?
        getFile(`${Application.Env.GetTempPath()}/wps/frameFile.json`) :
        dialog();
    if (text) {
        Array.prototype.unique = function () { //对象数组去重
            const map = new Map();
            return this.filter(e => !map.has(e.title) && map.set(e.title, null));
        }
        var config = JSON.parse(text),
            format = {}; //缩进级别:[段落号]
        delAllSildes();
        //标题
        addSlide(config.head, wps.Enum.ppLayoutTitle);
        //目录
        addSlide({
            title: '目录',
            body: config.body.map(e => { return { title: e.title, body: [] } }).unique()
        });
        //body
        for (let obj of config.body) {
            format = {};
            addSlide(obj);
        }
        //结束
        addSlide(config.end, wps.Enum.ppLayoutTitle);
        //选中第一个幻灯片
        pre.Slides.Item(1).Select();
    }

    // function group
    function addSlide(obj, PpSlideLayout = wps.Enum.ppLayoutText, index = pre.Slides.Count + 1) {
        //添加幻灯片
        let slide = pre.Slides.Add(index, PpSlideLayout);
        setSlide(slide, obj);
        return slide
    }
    function setSlide(slide, { title, body }) {
        //设置幻灯片属性
        slide.Shapes.Title.TextFrame.TextRange.Text = title;
        let bodyRange = slide.Shapes.Item(2).TextFrame.TextRange;
        bodyRange.Text = '';
        pWrite(bodyRange, body, 1);
        pFormat(bodyRange, format)
    }
    function pWrite(Range, body, n) {
        //迭代写入文本
        if (!Object.keys(format).includes(n + '')) format[n] = [];
        for (let obj of body) {
            Range.Text += obj.title + '\r';
            format[n].push(Range.Paragraphs().Count);
            pWrite(Range, obj.body, n + 1);
        }
    }
    function pFormat(Range, format) {
        //根据format确定缩进级别
        for (let level in format) {
            if (level == '1') continue;
            for (let i of format[level]) {
                Range.Paragraphs(i).IndentLevel = parseInt(level);
            }
        }
    }
    function delAllSildes() {
        console.log('清空幻灯片');
        while (pre.Slides.Count) { pre.Slides.Item(1).Delete() }
    }
    function getFile(path) {
        if (path === null) return false;
        else if (!path.includes('frameFile.json')) return dialog('路径错误');
        else {
            try { return wps.FileSystem.ReadFile(path) } //返回文本字符
            catch (err) { return dialog(err.message) }
        }
    }
    function dialog(message) {
        //打开文件窗口，直到路径正确
        message && alert(`ERROR: ${message}\n请选择 frameFile.json 文件`);
        let fd = wps.FileDialog(wps.Enum.msoFileDialogOpen);
        fd.AllowMultiSelect = false;
        fd.InitialFileName = Application.Env.GetTempPath() + '/';
        fd.Show();
        return getFile(fd.SelectedItems.Item(1));
    }
    return !0
}

// 显示帮助文档
function ShowHelp(control) {
    let helpID = wps.PluginStorage.getItem("help_id")
    if (!helpID) {
        let helpPanel = wps.CreateTaskPane(GetUrlPath() + "/ui/help.html");
        wps.PluginStorage.setItem("help_id", helpPanel.ID),
            helpPanel.Visible = true
    } else {
        let helpPanel = wps.GetTaskPane(helpID)
        helpPanel.Visible = !helpPanel.Visible
    }
    return true
}