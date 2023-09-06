function onload(ribbonUI) {
    wps.ribbonUI ??= ribbonUI;
    wps.Enum ??= {
        msoCTPDockPositionLeft: 0,
        msoCTPDockPositionRight: 2
    };
    return !0;
}
function checkCiteFormat(control) {
    //功能组
    function Reg2Str(...regs) {
        return regs.map(e =>
            typeof e == 'string' ? e : e.toString().slice(1, -1)
        ).join('');
    }
    function RegExpTest(range, reg, isMark = true) {
        if (range === null) return 1;
        const ranges = [range].flat();
        const res = reg.test(ranges.map(e => e.Text).join(''));
        if (!isMark) return res;
        if (res) {
            ranges.forEach(e => { e.Font.Color = 0 });
        } else {
            ranges.forEach(e => { e.Font.Color = 255 });
            ranges[0] && error_ranges.push(ranges[0]);
        }
        return res;
    }
    function setParagraph(p) {
        //非法符号替换，规定符号后加空格、纠正符号组合
        function add_space(Words) { //查找句子 > 查找单词
            return collection_operator(Words).map(e => {
                let text = e.Text;
                if (hasLink(text)) {
                    if (e.Words.Count == 1) isStop = true;
                    else return add_space(e.Words);
                }
                isStop || [...chr_spac].forEach(c =>
                    text = text.replace(new RegExp(` *\\${c} *`, 'g'), c == '&' ? ' & ' : c + ' ')
                );
                return text;
            }, 1).join('');
        }
        let isStop = false;
        p.Range.Text = add_space(p.Range.Sentences)
            .replace(/\. ,/g, '.,')
            .replace(/\. \. \./g, '...') //省略号
            .replace(/,\.\.\./g, ', ...') //逗号+省略号
            .replace(/ {2,}/g, ' ') //删除过多空格
            .replace(/\? \. /g, '? ')
            .replace(/\r{2,}/g, '\r'); //删除空行
        //分段
        const p_segments = cutParagraph(p);
        //检测类型：期刊文献，学位论文
        if (!p_segments) {
            p.Range.Font.Bold = !0;
            p.Range.Font.Color = 255;
            error_ranges.push(p.Range);
            return;
        }
        $.refer_cites.push({
            range: p.Range,
            segments: p_segments.map(e => [e].flat().map(e => e.Text).join('')),
        });
        const [authors, date, title, source] = p_segments;
        const isTh = /((硕|博)士学位论文|Unpublished (master\'s thesis|doctorial dissertation))/.test(p.Range.Text);
        RegExpTest(authors, reg_author);
        RegExpTest(date, reg_date);
        if (isTh) {
            RegExpTest(title, reg_th_title);
            RegExpTest(source, reg_th_source);
        } else {
            RegExpTest(title, reg_jo_title);
            RegExpTest(source, reg_jo_source);
        }
        //斜体
        const target = isTh ? title.at(-1) : source,
            words = target?.Words;
        let pass_string = '';
        if (!words) {
            return;
        }
        for (let w = 1; w <= words.Count; w++) {
            const word = words.Item(w);
            pass_string += word.Text;
            if (target.Text.includes('(') && /\($/.test(pass_string.slice(-1))) break;
            if (/ ?\d+, ?/.test(pass_string.slice(-3))) break;
            word.Italic = !0;
        }
    }
    function cutParagraph(p) {
        const sentences_operator = collection_operator(p.Range.Sentences);
        let authors, date, title, source; //各成分
        let last_marker = sentences_operator.obj.Count;
        if (last_marker < 4) {
            add_comment(p.Range, '缺少成分');
            return;
        }
        //最后一句首词为 net_head ? link,so : so
        if (hasLink(sentences_operator.at(-1).Words.Item(1).Text)) {
            last_marker--;
        }
        source = sentences_operator.at(last_marker - 1);
        //匹配日期 ? date : 标红加粗
        for (let i = 1; i < last_marker; i++) {
            const date_range = sentences_operator.at(i - 1);
            if (!RegExpTest(date_range, reg_date, false)) {
                continue;
            }
            date = date_range;
            authors = sentences_operator.slice(0, i - 1);
            title = sentences_operator.slice(i, last_marker - 1);
            break;
        }
        if (date) {
            return [authors, date, title, source]; //可能含有 []
        } else {
            add_comment(p.Range, '无法获取日期');
        }
    }
    const hasLink = str => ['doi', 'http'].some(e => str.includes(e));
    function main() {
        error_ranges = [];
        $.refer_cites = [];
        //删除过多换行
        sel().Text = sel().Text.replace(/\r{2,}/g, '\r');
        //纠正符号
        replaceAll(...chr_swap);
        //删除高亮
        sel().Range.HighlightColorIndex = 0;
        //格式
        set_font_format(sel().Font, { Size: 9 });
        set_paragraph_format(paragraphs_operator.obj, { CharacterUnitFirstLineIndent: -2 });
        //遍历每段
        paragraphs_operator.map(p => {
            if (p.Range.Text.length == 1) { //跳过空行
                return;
            }
            try {
                setParagraph(p);
            } catch (e) {
                p.Range.HighlightColorIndex = Enum.wdYellow;
                alert('CheckIndexFormat p_format error');
                console.error(p.Range.Text + '\n', e);
            }
        });
        return !0;
    }

    //常量
    const reg_cn_author = /(^\[?[\u4e00-\u9fa5]{2,4}(, [\u4e00-\u9fa5]{2,4}){0,5}(, (\.\.\. )?[\u4e00-\u9fa5]{2,4})?\. $)/,
        reg_en_author = /(^[a-zA-Z\-\']+,( [A-Z]\.){1,3}(, [a-zA-Z\-\']+,( [A-Z]\.){1,3}){0,5}(, (\&|\.\.\.) [a-zA-Z\-\']+,( [A-Z]\.){1,3})? $)/,
        reg_author = new RegExp('^(' + Reg2Str(reg_cn_author) + '|' + Reg2Str(reg_en_author) + ')'),
        reg_date = /(\(\d{4}[a-z]?\)\. )/, //(1234b). /
        reg_jo_title = /[^\(\)]*. /,
        reg_th_title = /[^]*[a-zA-Z\u4e00-\u9fa5] \(((硕|博)士学位论文|Unpublished (master\'s thesis|doctorial dissertation))\). /,
        reg_title = /([^]+[\.\?]|[^]+\? [^]+\.) /,
        reg_jo_source = /([a-zA-Z\u4e00-\u9fa5,\-\&\:\(\)\' ]{2,}, \d+(\(\d+(\-\d+)?\))?, (Article )?[a-zA-Z]?\d+([\-\+][a-zA-Z]?\d+){0,2}\. \r*)/,
        reg_th_source = /(^[a-zA-Z\u4e00-\u9fa5\' ]{4,}(, [a-zA-Z\u4e00-\u9fa5\' ]{2,}){0,3}\. \r*$)/,
        reg_source = new RegExp('^(' + Reg2Str(reg_jo_source) + '|' + Reg2Str(reg_th_source) + ')\]?\\r*'),
        link = /(^(https?|doi)\:\/\/[a-zA-Z\#\?\.\/]*\.\s*)?$/,
        reg = new RegExp(Reg2Str(reg_author, reg_date, reg_title, reg_source));
    const paragraphs_operator = collection_operator(sel().Paragraphs),
        chr_swap = ["（）【】，−－–：’？！ ．", "()[],---:'?! ."],
        chr_spac = ',.&:?…!';
    let error_ranges = [],
        timerRun = fn => {
            const st = +new Date();
            fn();
            return +new Date() - st;
        };

    //主逻辑
    del_comment(sel().Range);
    if (paragraphs_operator.obj.Count != 1) { //不止选择一段
        console.log('每秒词数：' + sel().Words.Count * 1e3 / timerRun(main));
        alert('检查成功，' + error_ranges.length + '个疑似错误');
    } else {
        main();
    }
    return !0;
}
function matchCites() {
    del_comment(sel().Range);
    replaceAll('（）', '()');
    const text = sel().Text;
    const reg_mt = /[^a-zA-Z\u4e00-\u9fa5\(\)]+[a-zA-Z\u4e00-\u9fa5]+(等人)?\(\d{4}\)/g, //。abc等人(2333)
        reg_gt = /\([^\(\)]+\d{4}\)/g; //(abc et al., 2333)
    //提取正文索引
    const main_cites = [ //使用正则从正文提取字符串速度快，所以不采用逐步分析range的方式
        ...text.matchAll(reg_mt),
        ...text.matchAll(reg_gt),
    ].flatMap(e =>
        e[0].split(';')
    );
    console.log('正文索引:', main_cites)
    console.log('参考文献:', $.refer_cites.map(({ segments }) => segments));
    //比较正文索引和参考文献
    const ok_cites = new Set(),
        ok_refer_cites = new Set(),
        fail_main_cites = [];
    main_cites.forEach(cite => { //遍历正文索引
        let find_times = 0;
        $.refer_cites.forEach(ref => { //遍历参考文献
            const name = ref.segments[0].match(/^([^,\.]+)/)[1], //Dickman
                date = ref.segments[1].match(/^\((\d{4}[a-z]?)\)/)[1]; //2020
            if (cite.includes(name) && cite.includes(date)) {
                ok_cites.add([name, date].join(', '));
                ok_refer_cites.add(ref);
                // ref.range.Font.ColorIndex = Enum.wdBlue;
                find_times++;
            }
        });
        find_times || fail_main_cites.push(cite);
    });
    $.refer_cites.forEach(ref => {
        ok_refer_cites.has(ref) || add_comment(ref.range, '多余参考文献');
    });
    console.log('匹配成功:', ok_cites);
    console.log('匹配成功:', ok_refer_cites);
    console.log('缺少参考文献:', fail_main_cites);
    //写入批注
    add_comment(sel().Range, fail_main_cites.length
        ? '缺少参考文献的正文索引\r' + fail_main_cites.join('\r')
        : '核对成功');
    return !0;
}
function writeResult() {
    confirm('是否打开最新版？')
        ? open_url_in_local(`https://cubxx.github.io/wpsAcademic/论文/ui/writeResult.html`)
        : open_url_in_local(location.origin + '/ui/writeResult.html');
    return !0;
}
const { createLink, insertLink } = function () {
    let uncited_bookmark_name = '';
    function get_last_bookmark() {
        return collection_operator(doc().Bookmarks).at(-1);
    }
    return {
        createLink() {
            // 删除之前没被引用的书签
            const last_bookmark = get_last_bookmark();
            if (uncited_bookmark_name === last_bookmark?.Name) {
                last_bookmark.Delete();
            }
            if (sel().Type == Enum.wdSelectionNormal) {
                uncited_bookmark_name = add_bookmark(sel().Range).Name;
                alert('创建成功');
            } else {
                alert('请先选中文本');
            }
            return !0;
        },
        insertLink() {
            const bookmark = get_last_bookmark();
            if (bookmark) {
                cite_bookmark(bookmark);
                uncited_bookmark_name = '';
            } else {
                alert('请先创建文本链接');
            }
            return !0;
        }
    }
}();

function addFigure() {
    sel().Text = '';
    //图表
    // const shape = doc().InlineShapes.AddChart(Enum.xlColumnClustered);
    // shape.Chart.HasTitle = false;
    // shape.Chart.SetElement(Enum.msoElementErrorBarStandardError); // 添加标准误
    // shape.Range.Paragraphs.Alignment = Enum.wdAlignParagraphCenter;
    // shape.Select();
    // shape.ConvertToShape(); // 转化为Shape对象
    //
    const ps = sel().Paragraphs;
    ps.Add();
    ps.Add();
    ps.Add();
    ps.Add();
    set_font_format(sel().Font);
    set_paragraph_format(sel().Paragraphs);
    //图注
    ps.Item(2).Range.Text = '注：±1个标准误\r';
    ps.Item(2).Range.Paragraphs.Alignment = Enum.wdAlignParagraphLeft;
    //图题
    ps.Item(3).Range.Text = `图${doc().Tables.Count + 1}\t${'标题'}`;
    ps.Item(3).Range.Fields.Add(ps.Item(3).Range.Words.Item(2)).Code.Text = ' SEQ 图 \\* ARABIC '; // 添加域
    ps.Item(3).Alignment = Enum.wdAlignParagraphCenter;
    sel().ParagraphFormat.CharacterUnitFirstLineIndent = 0;
    doc().Fields.Update(); // 更新文档所有域
    return !0;
}
function addTable() {
    sel().Text = '';
    const ps = sel().Paragraphs;
    ps.Add();
    ps.Add();
    ps.Add();
    ps.Add();
    set_font_format(sel().Font);
    set_paragraph_format(sel().Paragraphs);
    //表题
    ps.Item(1).Range.Text = `表${doc().Tables.Count + 1}\t标题\r`;
    ps.Item(1).Range.Font.Name = '黑体';
    ps.Item(1).Range.Fields.Add(ps.Item(1).Range.Words.Item(2)).Code.Text = ' SEQ 表 \\* ARABIC '; // 添加域
    ps.Item(1).Alignment = Enum.wdAlignParagraphCenter;
    //表注
    ps.Item(3).Range.Text = '注：N = 20, *p < 0.05, **p < 0.01, ***p < 0.001。';
    collection_operator(ps.Item(3).Range.Words).map(e => {
        if (/^ ?[Np] ?$/.test(e.Text)) {
            e.Font.Italic = true;
        }
    });
    ps.Item(3).Range.Paragraphs.Alignment = Enum.wdAlignParagraphLeft;
    sel().ParagraphFormat.CharacterUnitFirstLineIndent = 0;
    doc().Fields.Update(); // 更新文档所有域
    //创建表格
    const tableCount = doc().Tables.Count;
    ps.Item(2).Range.Select();
    sel().PasteAndFormat(Enum.wdChart); // 粘贴表格
    if (doc().Tables.Count !== tableCount + 1) {
        sel().Text = '作表失败 请先复制Excel表格\r';
        return !0;
    }
    const table = sel().Tables.Item(1);
    //设置表格格式
    const rows_operator = collection_operator(table.Rows);
    rows_operator.map(row => {
        row.SetHeight(13.8, Enum.wdRowHeightAuto);
    });
    table.AutoFitBehavior(2); // 根据活动窗口的宽度自动调整表格大小
    table.Borders.InsideLineStyle = 0;
    table.Borders.OutsideLineStyle = 0;
    function set_border_line(row, orientation, LineWidth) {
        row.Borders.Item(orientation).set({
            LineStyle: Enum.wdLineStyleSingle,
            LineWidth,
        });
    };
    set_border_line(rows_operator.at(0), Enum.wdBorderTop, Enum.wdLineWidth150pt);
    set_border_line(rows_operator.at(0), Enum.wdBorderBottom, Enum.wdLineWidth075pt);
    set_border_line(rows_operator.at(-1), Enum.wdBorderBottom, Enum.wdLineWidth150pt);
    //设置内容格式
    table.Range.Select();
    sel().ClearFormatting();
    set_font_format(sel().Font);
    return !0;
}
function addFlowChart() {
    const shp = doc().Shapes,
        config = { wh: [60, 40], position: { abs: [0, 0], rel: [0, 0] } };
    shp.AddTextbox(1, 0, 0, ...config.wh).TextFrame.TextRange.Text = '程序1';
    shp.AddTextbox(1, 0, 0, ...config.wh).TextFrame.TextRange.Text = '程序2';
    shp.AddTextbox(1, 0, 0, ...config.wh).TextFrame.TextRange.Text = '程序3';
    return !0;
}
// function createFrame() {
//     let config = {
//         "head": { "title": "标题", "body": [] },
//         "body": [/* {
//             "title": "一级标题",
//             "body": [{
//                 "title": "二级标题",
//                 "body": []
//             }]
//         } */],
//         "end": { "title": "感谢观看", "body": [] }
//     }
//     //确定标题/副标题
//     for (var pnHead = 1; pnHead <= sel().Paragraphs.Count; pnHead++) {
//         let p = sel().Paragraphs.Item(pnHead),
//             text = p.Range.Text.replace(/\r/g, '');
//         if (p.Alignment == Enum.wdAlignParagraphCenter) { //居中
//             switch (pnHead) {
//                 case 1: config.head.title = text; break;
//                 case 2: config.head.body.push({ "title": text, "body": [] }); break;
//                 default: break;
//             }
//         } else break;
//     }
//     //遍历body段落，写入config.body
//     let objs = [];
//     for (let pn = pnHead; pn <= sel().Paragraphs.Count; pn++) {
//         let p = sel().Paragraphs.Item(pn),
//             text = p.Range.Text.replace(/\r/g, '');
//         //段落数据对象
//         let obj = new (class {
//             constructor(pn, level, title) {
//                 this.super = config;
//                 this.pn = pn;
//                 this.OutlineLevel = level;
//                 this.data = { title: title, body: [] };
//             }
//             addSub(obj) { this.data.body.push(obj.data); obj.super = this.data }
//             addBro(obj) { this.super.body.push(obj.data); obj.super = this.super }
//         })(pn, p.OutlineLevel, text)
//         objs[pn] = obj;
//         //通过迭代，构造树形结构
//         (function structure(p, n) {
//             let pp = objs[p.pn - n]; //前一段
//             if (!pp) p.addBro(p);  //第一段
//             else {
//                 if (pp.OutlineLevel < p.OutlineLevel) { pp.addSub(p); }
//                 else if (pp.OutlineLevel == p.OutlineLevel) pp.addBro(p);
//                 else structure(p, n + 1);
//             }
//         })(obj, 1);
//     }
//     //一张幻灯片内容太多，则分成多份
//     Array.prototype.sum = function () { let s = 0; this.forEach(e => s += e); return s } //数组求和
//     for (let o of config.body) {
//         let oooNums = [];
//         for (let oo of o.body) oooNums.push(oo.body.length);
//         const maxNum = 4; //一张幻灯片上的三级标题最多数量
//         //三级对象总数大于最大值，则拆分
//         if (oooNums.sum() > maxNum) {
//             //判断拆分位置
//             let start = 0; //开始识别位置
//             for (let len = 1; len <= oooNums.length; len++) { //识别长度
//                 if (start + 1 < oooNums.length && //不识别最后一个元素
//                     oooNums.slice(start, start + len).sum() > maxNum) { //识别元素之和大于最大值
//                     let cutLen = len == 1 ? 1 : len - 1; //裁剪长度
//                     config.body.splice(config.body.indexOf(o), 0, { title: o.title, body: o.body.splice(0, cutLen) }); //拆分
//                     start += cutLen;
//                     len = 0;
//                 }
//             }
//         }
//     }
//     //写入本地
//     let path = `${Application.Env.GetTempPath()}/wps`;
//     wps.FileSystem.Mkdir(path); //创建wps文件夹
//     path += '/frameFile.json';
//     if (wps.FileSystem.WriteFile(path, JSON.stringify(config, null, 4))) alert(`创建成功`);
//     else alert(`创建失败`);
//     // wps.ShowDialog('file:///' + path, path); //查看文件内容
//     return !0
// }
function syntaxParser() {
    const paragraphs_operator = collection_operator(sel().Paragraphs).set({
        parse() {
            this.paragraph_info_objs = this.map((p, index) => {
                if (p.Range.Text.length === 1) {
                    return {};
                }
                const matchRes = p.Range.Text.match(/^\\([a-z]+) /);
                if (matchRes) {
                    p.Range.Text = matchRes.input.replace(matchRes[0], '');
                }
                return {
                    index,
                    identifier: matchRes?.[1] ?? 'p',
                };
            });
        },
        apply(fns) {
            this.paragraph_info_objs.map(({ index, identifier }) => {
                fns[identifier]?.(this.at(index - 1));
            });
        },
    });
    const identifier_fns = {
        title(p) { p.Alignment = Enum.wdAlignParagraphCenter; p.Range.Font.set({ Size: 22, NameFarEast: '黑体', }); },
        author(p) { p.Alignment = Enum.wdAlignParagraphCenter; p.Range.Font.set({ Size: 14, NameFarEast: '仿宋', }); },
        address(p) { },
        abstract(p) { p.Range.InsertBefore('摘要  '); p.Range.Font.set({ Size: 10.5, }); p.Range.Words.Item(1).Font.set({ NameFarEast: '黑体', }); },
        keywords(p) { p.Range.InsertBefore('关键词  '); p.Range.Font.set({ Size: 10.5, }); p.Range.Words.Item(1).Font.set({ NameFarEast: '黑体', }); },
        h(p) { p.Range.Font.set({ Size: 14, }); },
        hh(p) { p.Range.Font.set({ Size: 10.5, NameFarEast: '黑体', }); },
        hhh(p) { p.Range.Font.set({ Size: 10.5, NameFarEast: '黑体', }); },
        hhhh(p) { },
        p(p) { p.Range.Font.set({ Size: 10.5, }); p.set({ CharacterUnitFirstLineIndent: 2, }); operator_parser(p.Range); },
        link(p) { },
        table(p) { p.Range.Font.set({ Size: 9, }); },
        figure(p) { p.Range.Font.set({ Size: 7.5, }); },
        cite(p) { p.Range.Font.set({ Size: 9, }); },
    };
    const operator_fns = {
        '\\'(range) { range.Font.Italic = !0; },
        '_'(range) { range.Font.Subscript = !0; },
        '^'(range) { range.Font.Superscript = !0; },
    };
    const special_symbols = [
        ['alpha', 'α'],
        ['beta', 'β'],
        ['chi', 'χ'],
        ['eta', 'η'],
        ['times', '×'],
    ];
    function operator_parser(range) {
        const operators = Object.keys(operator_fns);
        function is_useless(e) {
            return e.Text.length === 1 && '\r\v '.includes(e.Text);
        }
        progressive_search(range, 'Sentences', (e, i, arr) => {
            if (is_useless(e)) {
                return;
            }
            if (operators.some(operator => e.Text === operator)) {
                // e: 操作符
                const next_e = arr.Item(i + 1);
                if (is_useless(next_e)) {
                    return;
                }
                const previous_e = arr.Item(i - 1);
                if (previous_e?.Text === '\\') {
                    return;
                }
                if (e.Text === '\\' && operators.some(operator => next_e.Text[0] === operator)) {
                    // e: 转义符
                } else {
                    operator_fns[e.Text](next_e);
                }
                e.Text = '';
                return;
            }
            if (operators.some(operator => e.Text.includes(operator))) {
                // e: 含有操作符的句子/单词
                return true;
            }
        });
    }
    function main() {
        if (paragraphs_operator.at(-1).Range.Text.length !== 1) {
            alert('选中文本的最后一行应为空行');
            return !0;
        }
        // 替换特殊符号
        replaceAll(...Object.keys(special_symbols[0]).map(col =>
            special_symbols.map(row =>
                (!!+col ? '' : '\\') + row[col]
            )
        ));
        // 构建list
        sel().ClearFormatting();
        const list_num = doc().Lists.Count;
        paragraphs_operator.parse();
        paragraphs_operator.apply({
            h(p) { p.Style = '标题 1'; },
            hh(p) { p.Style = '标题 2'; },
            hhh(p) { p.Style = '标题 3'; },
            hhhh(p) { p.Style = '标题 4'; },
            p(p) { p.Style = '正文'; },
        });
        if (doc().Lists.Count === list_num + 1) { //构建成功
            const list = doc().Lists.Item(list_num + 1);
            list?.ApplyListTemplate(wps.ListGalleries.Item(3).ListTemplates.Item(7)); //多级编号
        }
        // 识别标识符
        set_font_format(sel().Font);
        set_paragraph_format(sel().Paragraphs);
        paragraphs_operator.apply(identifier_fns);
    }
    main();
    return !0;
}
function showHelp(control) {
    const helpID = wps.PluginStorage.getItem("helpPanel_id");
    if (!helpID) {
        const helpPanel = wps.CreateTaskPane(location.origin + "/ui/help.html");
        wps.PluginStorage.setItem("helpPanel_id", helpPanel.ID);
        helpPanel.Visible = true
    } else {
        const helpPanel = wps.GetTaskPane(helpID)
        helpPanel.Visible = !helpPanel.Visible
    }
    return true
}
function update() {
    const get_latest_api = 'https://api.github.com/repos/Cubxx/wpsAcademic/releases/latest';
    const controller = new AbortController();
    setTimeout(() => controller.abort(), 3e3);
    fetch(get_latest_api, {
        signal: controller.signal,
    }).then(async res => {
        const data = await res.json();
        if (res.ok) {
            const { tag_name, body, zipball_url } = data;
            if ($.version == tag_name) {
                alert('已是最新版');
            } else if (confirm(`是否下载最新版?\n版本: ${tag_name}\n描述: ${body}`)) {
                open_url_in_local(zipball_url);
            }
        } else {
            alert('返回无效响应\n' + data.message);
        }
    }).catch(err => {
        const info = err.name == 'AbortError'
            ? '连接GitHub超时'
            : '无法连接GitHub';
        alert(info + '\n请尝试科学上网');
    });
    return !0;
}
function test() {
    return !0;
}

const UI = function (ui) {
    [
        'Enabled',
        'Label',
        'Screentip',
        'Supertip',
    ].forEach(prop => {
        ui['get' + prop] = function ({ Id }) {
            const value = ui.controls[Id]?.[prop] ?? '';
            return typeof value == 'function' ? value() : value;
        };
    });
    return ui;
}({
    controls: {
        r1: {
            Enabled: true,
            Label: '检查索引',
            Screentip: '检查参考文献是否按APA格式进行索引',
            Supertip: '请先选中参考文献',
        },
        r2: {
            Enabled: true,
            Label: '核对索引',
            Screentip: '核对正文索引和参考文献列表是否一一对应',
            Supertip: '请先选中正文',
        },
        w1: {
            Enabled: true,
            Label: '撰写结果',
            Screentip: '通过输入统计数据自动生成结果部分',
            Supertip: '单击打开网页程序',
        },
        w2: {
            Enabled: true,
            Label: '创建文链',
            Screentip: '创建文本链接',
            Supertip: '请先选中文本',
        },
        w3: {
            Enabled: true,
            Label: '插入文链',
            Screentip: '插入文本链接',
            Supertip: '将之前创建的文链插入光标处',
        },
        c1: {
            Enabled: true,
            Label: '作图',
            Screentip: '',
            Supertip: '',
        },
        c2: {
            Enabled: true,
            Label: '作表',
            Screentip: '以三线表格式粘贴表格',
            Supertip: '请先复制表格',
        },
        c3: {
            Enabled: true,
            Label: '流程图',
            Screentip: '',
            Supertip: '',
        },
        o1: {
            Enabled: true,
            Label: '格式解析',
            Screentip: '对格式文本进行解析',
            Supertip: '请先选中需要解析的文本',
        },
        o2: {
            Enabled: true,
            Label: '帮助',
            Screentip: '',
            Supertip: '',
        },
        o3: {
            Enabled: true,
            Label: '更新',
            Screentip: '当前版本',
            Supertip: $.version,
        },
    },
});