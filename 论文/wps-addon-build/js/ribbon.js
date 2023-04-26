function Paper_onload(ribbonUI) {
    if (typeof (wps.ribbonUI) != "object") {
        wps.ribbonUI = ribbonUI
    }
    if (typeof (wps.Enum) != "object") { // 如果没有内置枚举值
        wps.Enum = WPS_Enum
    }
    wps.PluginStorage.setItem("EnableFlag", !1) //往PluginStorage中设置一个标记，用于控制两个按钮的置灰
    wps.PluginStorage.setItem("ApiEventFlag", !1) //往PluginStorage中设置一个标记，用于控制ApiEvent的按钮label
    return true
}

// var dependent_variables = [], independent_variables = [];
function RangStr(range) {
    if (Array.isArray(range)) { let str = ''; return (range.forEach(e => { str += e.Text }), str) }
    else return range.Text;
}
function fontInit(Font, config = {}) {
    Object.assign(Font, {
        Size: 9,
        Bold: !1,
        Italic: !1,
        Color: 0,
        Name: '',
        NameAscii: "Times New Roman",
        NameFarEast: "宋体",
    }, config);
}
function BookmarkToField(source, refer, name) {
    //添加 书签-域 链接
    source.Bookmarks.Add(name);
    refer.Fields.Add(refer).Code.Text = ` REF ${name} `;
}
function loop(Items, func) {
    // Items.Count 可能会动态变化
    for (let i = 1; i <= Items.Count; i++) func(Items.Item(i), i, Items);
}
Function.prototype.step = function (note) {
    // 创建撤销组
    wps.UndoRecord.StartCustomRecord(note);
    this();
    wps.UndoRecord.EndCustomRecord();
}
var referIndex = { values: [], ranges: [], init: function () { this.values = [], this.ranges = [] } }, //参考文献
    sel = wps.Selection,
    doc = wps.ActiveDocument,
    _debug = {
        log: '', segments: {}, reds: { now: [], before: [], index: 0 }, num: 0, dbg: void 0,
        init: function () { this.log = '', this.segments = {}, this.reds.now = [], this.num = 0 },
        rcd_segs: function (segs) {
            this.segments = {}
            segs.forEach((e, i) => { this.segments[i + 'N_' + (e.length || 1)] = RangStr(e); })
            referIndex.values.push(Object.values(this.segments));
            referIndex.ranges.push(segs);
        },
        time: function (func, c) {
            let a = new Date().getTime(); func();
            let d = new Date().getTime() - a;
            c && console.log('\n-----\n运行时间：' + d);
            return d
        }
    }
function delComment(range) {
    range.Comments.Count && range.Comments.Item(1).Delete();
}
function addComment(range, content) {
    delComment(range);
    range.Comments.Add(range, content);
    range.Comments.Item(1).Author = '*Output';
}

// 索引格式检查
function CheckIndexFormat(control) {
    {//功能组
        var RegStr = function (...regs) {
            let str = '';
            regs.forEach(e => {
                if (typeof e == 'string') str += e;
                else str += e.toString().slice(1, -1);
            });
            return str
        }
        var RegExpTest = function (range, reg, ctrl = true) {
            if (range === null) return 1;
            let str = '', islist = Array.isArray(range);
            if (islist) {
                if (reg.test((range.forEach(e => { str += e.Text }), str))) { ctrl && range.forEach(e => { e.Font.Color = 0 }); return true }
                else { ctrl && (range.forEach(e => { e.Font.Color = 255 }), range[0] && _debug.reds.now.push(range[0])); return false }
            } else {
                if (reg.test(range.Text)) { ctrl && (range.Font.Color = 0); return true }
                else { ctrl && (range.Font.Color = 255, _debug.reds.now.push(range)); return false }
            }
        }
        var isNetname = function (str) { let i = 0; net_head.forEach(e => { i += str.includes(e) }); return i }
        var formated = function () {
            _debug.init();
            _debug.num = sel.Words.Count; //初始词数
            sel.Text = sel.Text.replace(/\r\r/g, '\r'); //删除过多换行
            //纠正符号错误
            for (let c = 0; c < chr_swap[0].length; c++) { sel.Text = sel.Text.replace(RegExp(chr_swap[0][c], 'g'), chr_swap[1][c]) }
            //删除突出强调
            sel.Range.HighlightColorIndex = 0;
            //整体字体格式
            fontInit(sel.Font);
            //遍历每段
            for (let n = 1; n <= pars.Count; n++) {
                let p = pars.Item(n);
                if (p.Range.Text.length == 1) { p_space.push(p); continue } //跳过空行
                function p_format(p) {
                    //非法符号替换，规定符号后加空格、纠正符号组合
                    function add_spac(Cells) { //查找句子 > 查找单词
                        let str = '';
                        for (let s = 1; s <= Cells.Count; s++) {
                            let stc_t = Cells.Item(s).Text;
                            if (isNetname(stc_t)) {
                                if (Cells.Item(s).Words.Count == 1) stop_add = true;
                                else stc_t = add_spac(Cells.Item(s).Words);
                            }
                            else if (!stop_add) for (let c of chr_spac) {
                                stc_t = stc_t.replace(new RegExp('[ ]*\\' + c + '[ ]*', 'g'), c == '&' ? ' & ' : c + ' ')
                            }
                            str += stc_t;
                        }
                        return str;
                    }
                    let stop_add = false;
                    p.Range.Text = add_spac(p.Range.Sentences)
                        .replace(/\. ,/g, '.,')
                        .replace(/\. \. \./g, '...') //省略号
                        .replace(/,\.\.\./g, ', ...') //逗号+省略号
                        .replace(/ {2,}/g, ' ') //删除过多空格
                        .replace(/\? \. /g, '? ')
                        .replace(/\r{2,}/g, '\r'); //删除空行
                    //分段
                    function p_slice(p) {
                        function r_merge(rs, a, b) { let arr = []; for (let i = a; i <= b; i++) { arr.push(rs.Item(i)) } return arr }
                        let authors, date, title, source; //各成分
                        let stcs = p.Range.Sentences, stc_n = stcs.Count;
                        if (stc_n < 4) return false;
                        //最后一句首词为 net_head ? link,so : so
                        [1, 3].forEach(() => {
                            let last = (n) => { return stcs.Item(n).Words.Item(1).Text }
                            if (isNetname(last(stc_n)) || last(stc_n)[0] == ']' ||
                                isNetname(last(stc_n - 1)) || last(stc_n - 1)[0] == ']') stc_n--;
                        });
                        source = stcs.Item(stc_n);
                        //匹配日期 ? date : 标红加粗
                        for (let i = 1; i < stc_n; i++) {
                            let s = stcs.Item(i);
                            if (RegExpTest(s, da, false)) {
                                if (1 > i - 1 || i + 1 > stc_n - 1) console.log('date位置错误 ' + i + ' ' + stc_n);
                                date = s;
                                authors = r_merge(stcs, 1, i - 1);
                                title = r_merge(stcs, i + 1, stc_n - 1);
                                break;
                            }
                        }
                        if (date === undefined)
                            return false;
                        else
                            return [authors, date, title, source]; //可能含有 []
                    }
                    let p_segs = p_slice(p), isTh = false;
                    p_segs && _debug.rcd_segs(p_segs);
                    //检测类型：期刊文献，学位论文
                    if (!p_segs) {
                        p.Range.Font.Bold = !0;
                        p.Range.Font.Color = 255;
                        _debug.reds.now.push(p.Range);
                    } else if (/((硕|博)士学位论文|Unpublished (master\'s thesis|doctorial dissertation))/.test(p.Range.Text)) {
                        RegExpTest(p_segs[0], au), RegExpTest(p_segs[1], da), RegExpTest(p_segs[2], th_ti), RegExpTest(p_segs[3], th_so);
                        _debug.log += 'thesis\n' + [au, da, th_ti, th_so].join('\n');
                        isTh = true;
                    } else {
                        RegExpTest(p_segs[0], au), RegExpTest(p_segs[1], da), RegExpTest(p_segs[2], ge_ti), RegExpTest(p_segs[3], ge_so);
                        _debug.log += 'journal\n' + [au, da, ge_ti, ge_so].join('\n');
                    }
                    //斜体
                    let trial_range = isTh ? p_segs[2].slice(-1)[0] : p_segs[3],
                        wrds = trial_range && trial_range.Words,
                        has_trial = '';
                    if (wrds) for (let w = 1; w <= wrds.Count; w++) {
                        has_trial += wrds.Item(w).Text;
                        if (trial_range.Text.includes('(') && /\($/.test(has_trial.slice(-1))) break;
                        else if (/ ?\d+, ?/.test(has_trial.slice(-3))) break;
                        wrds.Item(w).Italic = !0;
                    }
                    //段落格式
                    p.Space15(), //*倍行距
                        p.SpaceAfter = 0, //段后间距
                        p.SpaceBefore = 0,
                        p.CharacterUnitLeftIndent = 0, //左缩进量
                        p.CharacterUnitRightIndent = 0,
                        p.CharacterUnitFirstLineIndent = -2; //悬挂缩进2字符
                }
                try { p_format(p) } catch (e) { p.Range.HighlightColorIndex = 7; alert('CheckIndexFormat p_format error'); console.log(p.Range.Text + '\n', e) }
            }
            return !0;
        }
    }
    {//变量
        var cn_au = /(^\[?[\u4e00-\u9fa5]{2,4}(, [\u4e00-\u9fa5]{2,4}){0,5}(, (\.\.\. )?[\u4e00-\u9fa5]{2,4})?\. $)/,
            en_au = /(^[a-zA-Z\-\']+,( [A-Z]\.){1,3}(, [a-zA-Z\-\']+,( [A-Z]\.){1,3}){0,5}(, (\&|\.\.\.) [a-zA-Z\-\']+,( [A-Z]\.){1,3})? $)/,
            au = new RegExp('^(' + RegStr(cn_au) + '|' + RegStr(en_au) + ')'),
            da = /(\(\d{4}[a-z]?\)\. )/, //(1234b). /
            ti = /([^]+[\.\?]|[^]+\? [^]+\.) /,
            ge_so = /([a-zA-Z\u4e00-\u9fa5,\-\&\:\(\)\' ]{2,}, \d+(\(\d+(\-\d+)?\))?, (Article )?[a-zA-Z]?\d+([\-\+][a-zA-Z]?\d+){0,2}\. \r*)/,
            th_so = /(^[a-zA-Z\u4e00-\u9fa5\' ]{4,}(, [a-zA-Z\u4e00-\u9fa5\' ]{2,}){0,3}\. \r*$)/,
            so = new RegExp('^(' + RegStr(ge_so) + '|' + RegStr(th_so) + ')\]?\\r*'),
            link = /(^(https?|doi)\:\/\/[a-zA-Z\#\?\.\/]*\.\s*)?$/,
            reg = new RegExp(RegStr(au, da, ti, so)),
            net_head = ['doi', 'http'];
        var pars = sel.Paragraphs,
            p_space = [],
            chr_swap = ["（）【】，−－–：’？！ ．", "()[],---:'?! ."], chr_spac = ',.&:?…!',
            ge_ti = /[^\(\)]*. /,
            th_ti = /[^]*[a-zA-Z\u4e00-\u9fa5] \(((硕|博)士学位论文|Unpublished (master\'s thesis|doctorial dissertation))\). /;
    }
    switch (control.Id) {
        case 'format':
            if (pars.Count != 1) { //不止选择一段
                console.log('\n======cif======');
                referIndex.init();
                console.log('每秒词数：' + 1000 / (_debug.time(() => { formated() }, !1) / _debug.num));
                _debug.log = _debug.reds.now.length + '个疑似错误^_^';
                alert('检查成功，' + _debug.reds.now.length + '个疑似错误');
                _debug.reds.now[0] && _debug.reds.now[0].Select(); //选中第一个red
                _debug.reds.before = _debug.reds.now; //更新reds库
            } else formated();
            break;
        case 'up_err':
            if (_debug.reds.index != undefined) {
                0 < _debug.reds.index && _debug.reds.index--;
                _debug.reds.before[0] && _debug.reds.before[_debug.reds.index].Select();
            } break;
        case 'down_err':
            if (_debug.reds.index != undefined) {
                _debug.reds.before.length - 1 > _debug.reds.index && _debug.reds.index++;
                _debug.reds.before[0] && _debug.reds.before[_debug.reds.index].Select();
            } break;
        default: break;
    }
    return !0;
}

// 上下文核对
function CheckContext() {
    //字符串正则不匹配，返回空数组
    const _match = String.prototype.match;
    String.prototype.match = function (reg) {
        let res = _match.call(this, reg);
        return res == null ? [] : res;
    }
    //判断字符串是否包含数组中某个字符
    String.prototype.includeSome = function (list) {
        let res = false
        list.forEach(e => this.includes(e) && (res = true));
        return res
    }
    //删除数组重复元素
    Array.prototype.delRepeatElements = function () {
        let _this = [];
        this.forEach((e, i) => { _this[i] = e + '' }); //元素转化为字符串
        _this = [...new Set(_this)[Symbol.iterator]()]; //删除元素
        _this.forEach((e, i) => { _this[i] = e.split(',') }); //元素转化为数组
        return _this;
    }
    var sel = wps.Selection, txt = sel.Text;
    //判断是否有错误字符
    chr_swap = ['（(', '）)'];
    if (sel.Text.includeSome(chr_swap.map(e => { return e[0] }))) {
        //纠正错误字符
        for (let chr of chr_swap) {
            [sel.Find.Text, sel.Find.Replacement.Text] = chr;
            while (sel.Find.Execute()) { sel.Text = chr[1] } //查找替换
        }
        alert('错误字符纠正完毕\n请重新进行核对');
    } else {
        //主程序
        console.log('\n======ctx======');
        try {
            //提取正文索引
            var mt = /[^a-zA-Z\u4e00-\u9fa5\(\)]+[a-zA-Z\u4e00-\u9fa5]+(等人)?\(\d{4}\)/g, //。abc等人(2333)
                gt = /\([^\(\)]+\d{4}\)/g; //(abc et al., 2333)
            var textIndex = [], arr = [];
            for (let i of txt.match(gt)) {
                arr.push(...i.split('; '))
            }
            textIndex = [...txt.match(mt), ...arr]; //正文索引  使用正则从正文提取字符串速度快，所以不采用逐步分析range的方式
            console.log('正文索引:', textIndex)
            console.log('参考文献:', referIndex.values);
            //比较正文索引和参考文献
            var match_ok = [], match_lack = [];
            referIndex.values.forEach((ref, i) => { //遍历参考文献，事先标蓝
                // referIndex.ranges[i][1].Font.ColorIndex = 2;
                addComment(referIndex.ranges[i][1], '多余索引');
            });
            textIndex.forEach(cite => { //遍历正文索引
                let find_times = 0;
                referIndex.values.forEach((ref, i) => { //遍历参考文献
                    let name = ref[0].split(', ')[0].replace(/\. /g, ''), //Dickman
                        date = ref[1].split('. ')[0].slice(1, -1); //2020
                    if (cite.includes(name) && cite.includes(date)) {
                        find_times++;
                        match_ok.push([name, date]);
                        referIndex.ranges[i][1].Font.ColorIndex = 2; //1黑
                        delComment(referIndex.ranges[i][1]);
                    }
                });
                find_times == 0 && match_lack.push(cite);
            });
            console.log('匹配成功:', match_ok.delRepeatElements());
            console.log('缺少参考文献:', match_lack);
            //写入批注
            addComment(sel.Range, match_lack[0]
                ? '缺少参考文献的正文索引\r' + match_lack.join('\r')
                : '核对成功');
            alert('请在下次核对前\n先删除所有批注');
        } catch (err) { alert('CheckContext error'); console.log('\n', err) }
    }
    return !0;
}

// 挑选统计方法
function pickMethod(c) {
    return !0
}

// 撰写结果
function writeResult() {
    shellExecuteByOAAssist(`${GetUrlPath()}/ui/writeResult.html`);
    return !0;
}

// 作图
function addFigure() {
    let ps = sel.Paragraphs;
    //图表
    let shape = doc.InlineShapes.AddChart(wps.Enum.xlColumnClustered);
    shape.Chart.HasTitle = false;
    shape.Chart.SetElement(wps.Enum.msoElementErrorBarStandardError); // 添加标准误
    shape.Range.Paragraphs.Alignment = wps.Enum.wdAlignParagraphCenter;
    shape.Select();
    // shape.ConvertToShape(); // 转化为Shape对象
    ps.Add(); ps.Add(); ps.Add(); ps.Add();
    //图注
    fontInit(sel.Font);
    ps.Item(2).Range.Text = '注：±1个标准误\r';
    ps.Item(2).Range.Paragraphs.Alignment = wps.Enum.wdAlignParagraphLeft;
    //图题
    ps.Item(3).Range.Text = `图${doc.Tables.Count + 1}\t${'标题'}`;
    ps.Item(3).Range.Fields.Add(ps.Item(3).Range.Words.Item(2)).Code.Text = ' SEQ 图 \\* ARABIC '; // 添加域
    ps.Item(3).Alignment = wps.Enum.wdAlignParagraphCenter;
    doc.Fields.Update(); // 更新文档所有域
    return !0;
}

// 作表
function addTable() {
    sel.Text = '';
    let ps = sel.Paragraphs, table;
    ps.Add(); ps.Add(); ps.Add(); ps.Add();
    fontInit(sel.Font);
    //表题
    ps.Item(1).Range.Text = `表${doc.Tables.Count + 1}\t标题\r`;
    ps.Item(1).Range.Font.Name = '黑体';
    ps.Item(1).Range.Fields.Add(ps.Item(1).Range.Words.Item(2)).Code.Text = ' SEQ 表 \\* ARABIC '; // 添加域
    ps.Item(1).Alignment = wps.Enum.wdAlignParagraphCenter;
    doc.Fields.Update(); // 更新文档所有域
    //表注
    ps.Item(3).Range.Text = '注：*表示p < 0.05';
    ps.Item(3).Range.Words.Item(5).Font.Italic = true; // p斜体
    ps.Item(3).Range.Paragraphs.Alignment = wps.Enum.wdAlignParagraphLeft;
    //创建表格
    let tableCount = doc.Tables.Count;
    ps.Item(2).Range.Select();
    sel.PasteAndFormat(wps.Enum.wdChart); // 粘贴表格
    if (doc.Tables.Count !== tableCount + 1) {
        doc.Undo(); // 撤销粘贴
        table = doc.Tables.Add(sel.Range, 3, 3);
        sel.Text = '作表失败 请先复制Excel表格';
    } else table = sel.Tables.Item(1);
    //设置表格格式
    loop(table.Rows, e => e.Height = 13.8); // 行高 13.8
    table.AutoFitBehavior(2); // 根据活动窗口的宽度自动调整表格大小
    table.Borders.InsideLineStyle = 0;
    table.Borders.OutsideLineStyle = 0;
    Object.assign(table.Rows.Item(1).Borders.Item(wps.Enum.wdBorderTop), { LineStyle: 1, LineWidth: wps.Enum.wdLineWidth150pt });
    Object.assign(table.Rows.Item(2).Borders.Item(wps.Enum.wdBorderTop), { LineStyle: 1, LineWidth: wps.Enum.wdLineWidth075pt });
    Object.assign(table.Rows.Item(table.Rows.Count).Borders.Item(wps.Enum.wdBorderBottom), { LineStyle: 1, LineWidth: wps.Enum.wdLineWidth150pt });
    //设置内容格式
    table.Range.Select();
    sel.ClearFormatting();
    fontInit(sel.Font);
    return !0;
}

// 流程图
function addFlowChart() {
    let s = wps.ActiveDocument.Shapes,
        config = { wh: [60, 40], position: { abs: [0, 0], rel: [0, 0] } };
    s.AddTextbox(1, 0, 0, ...config.wh).TextFrame.TextRange.Text = '程序1';
    s.AddTextbox(1, 0, 0, ...config.wh).TextFrame.TextRange.Text = '程序2';
    s.AddTextbox(1, 0, 0, ...config.wh).TextFrame.TextRange.Text = '程序3';
    wps.ShowDialog('cubxx.github.io/Game-Lib/', '', 1280, 720, true);
    return !0;
}

// 创建框架，用于生成ppt
function createFrame() {
    let config = {
        "head": { "title": "标题", "body": [] },
        "body": [/* {
            "title": "一级标题",
            "body": [{
                "title": "二级标题",
                "body": []
            }]
        } */],
        "end": { "title": "感谢观看", "body": [] }
    }
    //确定标题/副标题
    for (var pnHead = 1; pnHead <= sel.Paragraphs.Count; pnHead++) {
        let p = sel.Paragraphs.Item(pnHead),
            text = p.Range.Text.replace(/\r/g, '');
        if (p.Alignment == wps.Enum.wdAlignParagraphCenter) { //居中
            switch (pnHead) {
                case 1: config.head.title = text; break;
                case 2: config.head.body.push({ "title": text, "body": [] }); break;
                default: break;
            }
        } else break;
    }
    //遍历body段落，写入config.body
    let objs = [];
    for (let pn = pnHead; pn <= sel.Paragraphs.Count; pn++) {
        let p = sel.Paragraphs.Item(pn),
            text = p.Range.Text.replace(/\r/g, '');
        //段落数据对象
        let obj = new (class {
            constructor(pn, level, title) {
                this.super = config;
                this.pn = pn;
                this.OutlineLevel = level;
                this.data = { title: title, body: [] };
            }
            addSub(obj) { this.data.body.push(obj.data); obj.super = this.data }
            addBro(obj) { this.super.body.push(obj.data); obj.super = this.super }
        })(pn, p.OutlineLevel, text)
        objs[pn] = obj;
        //通过迭代，构造树形结构
        (function structure(p, n) {
            let pp = objs[p.pn - n]; //前一段
            if (!pp) p.addBro(p);  //第一段
            else {
                if (pp.OutlineLevel < p.OutlineLevel) { pp.addSub(p); }
                else if (pp.OutlineLevel == p.OutlineLevel) pp.addBro(p);
                else structure(p, n + 1);
            }
        })(obj, 1);
    }
    //一张幻灯片内容太多，则分成多份
    Array.prototype.sum = function () { let s = 0; this.forEach(e => s += e); return s } //数组求和
    for (let o of config.body) {
        let oooNums = [];
        for (let oo of o.body) oooNums.push(oo.body.length);
        const maxNum = 4; //一张幻灯片上的三级标题最多数量
        //三级对象总数大于最大值，则拆分
        if (oooNums.sum() > maxNum) {
            //判断拆分位置
            let start = 0; //开始识别位置
            for (let len = 1; len <= oooNums.length; len++) { //识别长度
                if (start + 1 < oooNums.length && //不识别最后一个元素
                    oooNums.slice(start, start + len).sum() > maxNum) { //识别元素之和大于最大值
                    let cutLen = len == 1 ? 1 : len - 1; //裁剪长度
                    config.body.splice(config.body.indexOf(o), 0, { title: o.title, body: o.body.splice(0, cutLen) }); //拆分
                    start += cutLen;
                    len = 0;
                }
            }
        }
    }
    //写入本地
    let path = `${Application.Env.GetTempPath()}/wps`;
    wps.FileSystem.Mkdir(path); //创建wps文件夹
    path += '/frameFile.json';
    if (wps.FileSystem.WriteFile(path, JSON.stringify(config, null, 4))) alert(`创建成功`);
    else alert(`创建失败`);
    // wps.ShowDialog('file:///' + path, path); //查看文件内容
    return !0
}

// 识别大纲
function identifyOutline() {
    for (let pn = 1; pn <= sel.Paragraphs.Count; pn++) {
        let p = sel.Paragraphs.Item(pn);
        //识别标题
        if (p.Alignment == wps.Enum.wdAlignParagraphCenter) {
            switch (pn) {
                case 1: p.Style = '中文题目'; break;
                case 2: p.Style = '作者姓名'; break;
                default: addComment(p.Range, '为什么这个段落居中了'); break;
            }
        } else {
            //识别正文级别
            if (check(/^\d\.? ([\s\S]*)/, p)) p.Style = '标题 1';
            else if (check(/^\d\.\d\.? ([\s\S]*)/, p)) p.Style = '标题 2';
            else if (check(/^\d\.\d\.\d\.? ([\s\S]*)/, p)) p.Style = '标题 3';
            else p.Style = '正文';
        }
    }
    //正则提取
    function check(reg, p) {
        let result = reg.exec(p.Range.Text);
        if (result) {
            p.Range.Text = result[1];
            return true;
        }
    }
    return !0
}

// 测试功能
function test() {
    let sel = wps.Selection,
        inp = (e) => { sel.TypeText(e) },
        font = sel.Font,
        align = (e) => { sel.ParagraphFormat.Alignment = e };
    let all_idioms = [];
    try {
        for (let n = 1; n <= 14; n++) {
            //title
            font.Bold = true; font.Size = 14; font.NameAscii = "Times New Roman"; font.NameFarEast = "Times New Roman";
            align(1); //center
            inp('Week ' + n + '\n');
            //idiom 3
            font.Bold = false;
            align(0); //left
            let idioms = all_idioms.slice(3 * n - 3, 3 * n);
            idioms.forEach((e, i) => {
                inp((i + 1) + ')\t“' + e[0] + '”\n');
                inp('\tIt means ' + e[1] + '\n');
                inp('\tE.G. ' + e[2].replace(/\;/g, ',') + '\n\n');
            });
            //reflection 50 words
            align(1);
            inp('Weekly Reflection\n');
            align(0);
            inp(all_idioms[n - 1][3] + '\n');
        }
    } catch (err) { console.log(err) }
    return !0;
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

function Edit(control) {
    debugger;
    console.log(control);
    return !0
}