import { sel, doc, Enum, $ } from "./util.js";
import { Collection_decorator, Range_decorator } from './decorator.js'
const actions = {
    onload(ribbonUI) {
        wps.ribbonUI ??= ribbonUI;
        wps.Enum ??= {
            msoCTPDockPositionLeft: 0,
            msoCTPDockPositionRight: 2
        };
        return !0;
    },
    writeResult() {
        confirm('是否打开稳定版？')
            ? $.open_url_in_local(location.origin + '/ui/writeResult.html')
            : $.open_url_in_local(`https://cubxx.github.io/wpsAcademic/论文/ui/writeResult.html`)
        return !0;
    },
    addFigure() {
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
        new Range_decorator(sel().Range).set_default_text_format();
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
    },
    addTable() {
        sel().Text = '';
        const ps = sel().Paragraphs;
        ps.Add();
        ps.Add();
        ps.Add();
        ps.Add();
        new Range_decorator(sel().Range).set_default_text_format();
        //表题
        ps.Item(1).Range.Text = `表${doc().Tables.Count + 1}\t标题\r`;
        ps.Item(1).Range.Font.Name = '黑体';
        ps.Item(1).Range.Fields.Add(ps.Item(1).Range.Words.Item(2)).Code.Text = ' SEQ 表 \\* ARABIC '; // 添加域
        ps.Item(1).Alignment = Enum.wdAlignParagraphCenter;
        //表注
        ps.Item(3).Range.Text = '注：N = 20, *p < 0.05, **p < 0.01, ***p < 0.001。';
        new Collection_decorator(ps.Item(3).Range.Words).map(e => {
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
        const rows_operator = new Collection_decorator(table.Rows);
        rows_operator.map(row => {
            row.SetHeight(13.8, Enum.wdRowHeightAuto);
        });
        table.AutoFitBehavior(2); // 根据活动窗口的宽度自动调整表格大小
        table.Borders.InsideLineStyle = 0;
        table.Borders.OutsideLineStyle = 0;
        set_border_line(rows_operator.at(0), Enum.wdBorderTop, Enum.wdLineWidth150pt);
        set_border_line(rows_operator.at(0), Enum.wdBorderBottom, Enum.wdLineWidth075pt);
        set_border_line(rows_operator.at(-1), Enum.wdBorderBottom, Enum.wdLineWidth150pt);
        function set_border_line(row, orientation, LineWidth) {
            row.Borders.Item(orientation).set({
                LineStyle: Enum.wdLineStyleSingle,
                LineWidth,
            });
        };
        //设置内容格式
        table.Range.Select();
        sel().ClearFormatting();
        new Range_decorator(sel().Range).set_default_text_format();
        return !0;
    },
    addFlowChart() {
        const shp = doc().Shapes,
            config = { wh: [60, 40], position: { abs: [0, 0], rel: [0, 0] } };
        shp.AddTextbox(1, 0, 0, ...config.wh).TextFrame.TextRange.Text = '程序1';
        shp.AddTextbox(1, 0, 0, ...config.wh).TextFrame.TextRange.Text = '程序2';
        shp.AddTextbox(1, 0, 0, ...config.wh).TextFrame.TextRange.Text = '程序3';
        return !0;
    },
    syntaxParser() {
        const paragraphs = new Collection_decorator(sel().Paragraphs);
        if (!$.is_null_range(paragraphs.at(-1).Range)) {
            alert('选中文本的最后一行应为空行');
            return !0;
        }
        const paragraphs_operator = {
            infos: paragraphs.map((p, index) => {
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
            }),
            apply(fns) {
                this.infos.map(({ index, identifier }) => {
                    fns[identifier]?.(paragraphs.at(index - 1));
                });
            },
        };
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
            new Range_decorator(range).progressive_search('Sentences', (e, i, arr) => {
                if ($.is_null_range(e)) {
                    return;
                }
                if (operators.some(operator => e.Text === operator)) {
                    // e: 操作符
                    const next_e = arr.Item(i + 1);
                    if ($.is_null_range(next_e)) {
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
            // 替换特殊符号
            $.replaceAll(...Object.keys(special_symbols[0]).map(col =>
                special_symbols.map(row =>
                    (!!+col ? '' : '\\') + row[col]
                )
            ));
            // 构建list
            sel().ClearFormatting();
            const list_num = doc().Lists.Count;
            paragraphs_operator.apply({
                h(p) { p.Style = '标题 1'; },
                hh(p) { p.Style = '标题 2'; },
                hhh(p) { p.Style = '标题 3'; },
                hhhh(p) { p.Style = '标题 4'; },
                p(p) { p.Style = '正文'; },
            });
            if (doc().Lists.Count === list_num + 1) { //构建成功
                const list = doc().Lists.Item(list_num + 1);
                list.ApplyListTemplate(wps.ListGalleries.Item(3).ListTemplates.Item(7)); //多级编号
            }
            // 识别标识符
            new Range_decorator(sel().Range).set_default_text_format();
            paragraphs_operator.apply(identifier_fns);
        }
        main();
        return !0;
    },
    showHelp() {
        const helpID = wps.PluginStorage.getItem("helpPanel_id");
        if (!helpID) {
            const helpPanel = wps.CreateTaskPane(location.origin + "/ui/help.html");
            wps.PluginStorage.setItem("helpPanel_id", helpPanel.ID);
            helpPanel.Visible = true
        } else {
            const helpPanel = wps.GetTaskPane(helpID)
            helpPanel.Visible = !helpPanel.Visible
        }
        return !0
    },
    update() {
        const get_latest_api = 'https://api.github.com/repos/Cubxx/wpsAcademic/releases/latest';
        const controller = new AbortController();
        setTimeout(() => controller.abort(), 3e3);
        fetch(get_latest_api, {
            signal: controller.signal,
        }).then(async res => {
            const data = await res.json();
            if (res.ok) {
                const { tag_name, body, zipball_url } = data;
                if ($.version === tag_name) {
                    alert('已是最新版');
                } else if (confirm(`最新版: ${tag_name}\n${body}\n是否下载?`)) {
                    $.open_url_in_local(zipball_url);
                }
            } else {
                alert('返回无效响应\n' + data.message);
            }
        }).catch(err => {
            const info = err.name === 'AbortError'
                ? '连接GitHub超时'
                : '无法连接GitHub';
            alert(info + '\n请尝试科学上网');
        });
        return !0;
    },
    test() {
        return !0;
    },
};
const cite_actions = function () {
    /** @type{Map<{range: Range, segments: string[]}, string[]>} */
    const rbCite_pairs = new Map();
    return {
        checkCites() {
            if ($.is_null_range(sel().Range)) {
                alert('请先选中参考文献');
                return;
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
                reg_cite = new RegExp(Reg2Str(reg_author, reg_date, reg_title, reg_source));
            var p_range = new Range_decorator();
            //功能组
            const hasLink = str => ['doi', 'http'].some(e => str.includes(e));
            function Reg2Str(...regs) {
                return regs.map(e =>
                    typeof e == 'string' ? e : e.toString().slice(1, -1)
                ).join('');
            }
            function RegExpTest(range, reg, shouldMark = true) {
                if (range === null) return 1;
                const ranges = [range].flat();
                const res = reg.test(ranges.map(e => e.Text).join(''));
                if (!shouldMark) return res;
                if (res) {
                    ranges.forEach(e => { e.Font.Color = 0 });
                } else {
                    ranges.forEach(e => { e.Font.Color = 255 });
                }
                return res;
            }
            function setParagraph(p) {
                //非法符号替换，规定符号后加空格、纠正符号组合
                function add_space(Words) { //查找句子 > 查找单词
                    return new Collection_decorator(Words).map(e => {
                        let text = e.Text;
                        if (hasLink(text)) {
                            if (e.Words.Count == 1) isStop = true;
                            else return add_space(e.Words);
                        }
                        if (!isStop) {
                            [...',.&:?…!'].forEach(c =>
                                text = text.replace(new RegExp(` *\\${c} *`, 'g'), c == '&' ? ' & ' : c + ' ')
                            );
                        }
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
                    return;
                }
                rbCite_pairs.set({
                    range: p_segments[1],
                    segments: p_segments.map(e => [e].flat().map(e => e.Text).join('')),
                }, []);
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
                const sentences = new Collection_decorator(p.Range.Sentences);
                let last_marker = sentences.obj.Count;
                if (last_marker < 4) {
                    p_range.add_comment('这可能是不完整的索引');
                    return;
                }
                if (hasLink(sentences.at(-1).Words.Item(1).Text)) { //最后一句是否为网址
                    last_marker--;
                }
                let authors, date, title, source; //各成分
                source = sentences.at(last_marker - 1);
                for (let i = 1; i < last_marker; i++) {
                    const date_range = sentences.at(i - 1);
                    if (!RegExpTest(date_range, reg_date, false)) { //匹配日期
                        continue;
                    }
                    date = date_range;
                    authors = sentences.slice(0, i - 1);
                    title = sentences.slice(i, last_marker - 1);
                    break;
                }
                if (date) {
                    return [authors, date, title, source]; //可能含有 []
                } else {
                    p_range.add_comment('程序无法获取索引日期');
                    return;
                }
            }
            function main() {
                const range = new Range_decorator(sel().Range);
                const paragraphs = new Collection_decorator(sel().Paragraphs);
                rbCite_pairs.clear();
                sel().Text = sel().Text.replace(/\r{2,}/g, '\r'); //删除过多换行
                $.replaceAll("（）【】，−－–：’？！ ．", "()[],---:'?! ."); //纠正符号
                range.del_comment();
                range.set_font_format({ Size: 9 });
                range.set_paragraph_format({ CharacterUnitFirstLineIndent: -2 });
                function loop() {
                    paragraphs.map(p => {
                        if ($.is_null_range(p.Range)) { //跳过空行
                            return;
                        }
                        p_range = new Range_decorator(p.Range);
                        try {
                            setParagraph(p);
                        } catch (e) {
                            p.Range.HighlightColorIndex = Enum.wdYellow;
                            p_range.add_comment('意料外的错误\v' + e, 'err');
                            console.error(e);
                        }
                    });
                }
                console.log('每秒词数：' + sel().Words.Count * 1e3 / loop.timed());
                alert('检查成功\n记录了' + rbCite_pairs.size + '条参考文献');
            }
            main();
            return !0;
        },
        matchCites() {
            if ($.is_null_range(sel().Range)) {
                alert('请先选中正文');
                return;
            }
            if (rbCite_pairs.size === 0) {
                alert('还没有记录参考文献\n请先检查索引');
                return;
            }
            const sel_range = new Range_decorator(sel().Range);
            sel_range.del_comment();
            $.replaceAll('（）', '()');
            //提取正文索引
            const text = sel().Text;
            const reg_mt = /[^a-zA-Z\u4e00-\u9fa5\(\)]{1}[a-zA-Z\u4e00-\u9fa5]+\(\d{4}[a-z]?\)/g, //。abc(2333a)
                reg_gt = /\([^\(\)]+\d{4}[a-z]?\)/g; //(abc et al., 2333a)
            const body_cites = [
                ...text.matchAll(reg_mt),
                ...text.matchAll(reg_gt),
            ].flatMap(e =>
                e[0].split(';')
            );
            const refer_cites = [...rbCite_pairs.keys()];
            console.log('正文索引', body_cites)
            console.log('参考文献', refer_cites.map(({ segments }) => segments));
            //匹配
            const fail_body_cites = [];
            loop_body: for (let cite of body_cites) {
                loop_refer: for (let ref of refer_cites) {
                    const name = ref.segments[0].match(/^([^,\.]+)/)[1], //abc
                        date = ref.segments[1].match(/^\((\d{4}[a-z]?)\)/)[1]; //2333a
                    const isOK = [name, date].every(e => cite.includes(e));
                    if (isOK) {
                        rbCite_pairs.get(ref).push(cite);
                        continue loop_body;
                    }
                }
                fail_body_cites.push(cite);
            }
            refer_cites.forEach(ref => {
                rbCite_pairs.get(ref).length || new Range_decorator(ref.range).add_comment('这可能是未被引用的文献\v需要人工核对');
            });
            console.log('匹配结果', [...rbCite_pairs.entries()]);
            //写入批注
            fail_body_cites.length
                ? sel_range.add_comment('这些正文索引缺少参考文献\v' + fail_body_cites.join('\v'))
                : sel_range.add_comment('正文核对成功', 'ok');
            sel_range.add_comment('参考文献核对情况见参考文献列表\v没有批注则核对成功', 'info');
            return !0;
        },
    }
}();
const link_actions = function () {
    let hasCited = true;
    const get_last_bookmark = () => new Collection_decorator(doc().Bookmarks).at(-1);
    const get_sel_range = () => new Range_decorator(sel().Range);
    return {
        createLink() {
            if (!hasCited) { //删除之前没被引用的书签
                get_last_bookmark().Delete();
            }
            if (sel().Type === Enum.wdSelectionNormal) {
                get_sel_range().add_bookmark();
                hasCited = false;
                alert('创建成功');
            } else {
                alert('请先选中文本');
            }
        },
        insertLink() {
            const bookmark = get_last_bookmark();
            if (bookmark) {
                get_sel_range().cite_bookmark(bookmark);
                hasCited = true;
            } else {
                alert('请先创建文本链接');
            }
            return !0;
        }
    }
}();
actions.set(cite_actions, link_actions);
// 添加撤销组
function step(...args) {
    wps.UndoRecord.StartCustomRecord('论文_' + this.name);
    try {
        this.apply(actions, args);
    } catch (e) {
        new Range_decorator(sel().Range).add_comment('意料外的错误\v请撤销这个操作\v' + e, 'err');
        console.error(e);
    }
    wps.UndoRecord.EndCustomRecord();
    return !0
}
[
    'checkCites', 'matchCites',
    'createLink', 'insertLink',
    'addFigure', 'addTable', 'addFlowChart',
    'syntaxParser',
].forEach(key => {
    actions[key] = step.bind(actions[key]);
});
export default actions;