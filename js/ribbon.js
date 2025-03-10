import { config } from '../config.js';
import { open_url_in_wps, root_path } from '../lib/web.js';
import { Collection, Range } from './decorator.js';
import { $, Enum, doc, sel } from './util.js';
const app_base = {
    onload(ribbonUI) {
        wps.ribbonUI ??= ribbonUI;
        wps.Enum ??= {
            msoCTPDockPositionLeft: 0,
            msoCTPDockPositionRight: 2,
        };
        return !0;
    },
    update() {
        // open_url_in_wps(config.ui.update.html, '更新', 550, 150);
        wps.OAAssist.ShellExecute('powershell', `irm ${root_path}/setup.ps1 | iex`);
        return !0;
    },
    test() {
        return !0;
    },
};
const app_ui = {
    write_result() {
        open_url_in_wps(config.ui.write_result.html, '撰写结果工具', 900, 550);
        // open_url_in_local(config.ui.write_result.html);
        return !0;
    },
    show_help() {
        const helpID = wps.PluginStorage.getItem('helpPanel_id');
        if (!helpID) {
            const helpPanel = wps.CreateTaskPane(root_path + config.ui.help.html);
            wps.PluginStorage.setItem('helpPanel_id', helpPanel.ID);
            helpPanel.Visible = true;
        } else {
            const helpPanel = wps.GetTaskPane(+helpID);
            helpPanel.Visible = !helpPanel.Visible;
        }
        return !0;
    },
};
const other = {
    add_figure() {
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
        new Range(sel().Range).set_default_text_format();
        //图注
        ps.Item(2).Range.Text = '注：±1个标准误\r';
        ps.Item(2).Range.Paragraphs.Alignment = Enum.wdAlignParagraphLeft;
        //图题
        ps.Item(3).Range.Text = `图${doc().Tables.Count + 1}\t${'标题'}`;
        ps.Item(3).Range.Fields.Add(ps.Item(3).Range.Words.Item(2)).Code.Text =
            ' SEQ 图 \\* ARABIC '; // 添加域
        ps.Item(3).Alignment = Enum.wdAlignParagraphCenter;
        sel().ParagraphFormat.CharacterUnitFirstLineIndent = 0;
        doc().Fields.Update(); // 更新文档所有域
    },
    add_table() {
        sel().Text = '';
        const ps = sel().Paragraphs;
        ps.Add();
        ps.Add();
        ps.Add();
        ps.Add();
        new Range(sel().Range).set_default_text_format();
        //表题
        ps.Item(1).Range.Text = `表${doc().Tables.Count + 1}\t标题\r`;
        ps.Item(1).Range.Font.Name = '黑体';
        ps.Item(1).Range.Fields.Add(ps.Item(1).Range.Words.Item(2)).Code.Text =
            ' SEQ 表 \\* ARABIC '; // 添加域
        ps.Item(1).Alignment = Enum.wdAlignParagraphCenter;
        //表注
        ps.Item(3).Range.Text = '注：N = 20, *p < 0.05, **p < 0.01, ***p < 0.001。';
        new Collection(ps.Item(3).Range.Words).map((e) => {
            if (/^ ?[Np] ?$/.test(e.Text)) {
                e.Font.Italic = 1;
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
            return;
        }
        const table = sel().Tables.Item(1);
        //设置表格格式
        const rows_operator = new Collection(table.Rows);
        rows_operator.map((row) => {
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
        }
        //设置内容格式
        table.Range.Select();
        sel().ClearFormatting();
        new Range(sel().Range).set_default_text_format();
    },
    add_flowChart() {
        doc().InlineShapes.AddWebShapeEx('/ui/mermaid.html');
    },
    syntax_parser() {
        const paragraphs = new Collection(sel().Paragraphs);
        if (!$.is_null_range(paragraphs.at(-1).Range)) {
            paragraphs.at(-1).Range.Text += '\r';
        }
        const paragraphs_operator = {
            infos: paragraphs.map((p, index) => {
                if ($.is_null_range(p.Range)) {
                    return {};
                }
                const matchRes = p.Range.Text.match(/^\\([a-z]+) /);
                if (matchRes) {
                    p.Range.Text = matchRes.input?.replace(matchRes[0], '') ?? '';
                }
                return {
                    index,
                    identifier: matchRes?.[1] ?? 'p',
                };
            }),
            /**
             * @param {Record<string,(p:Wps.WpsParagraph)=>void>} fns
             */
            apply(fns) {
                this.infos.map(({ index, identifier }) => {
                    if (!index) return;
                    fns[identifier]?.(paragraphs.at(index - 1));
                });
            },
        };
        const identifier_fns = {
            title(p) {
                p.Alignment = Enum.wdAlignParagraphCenter;
                p.Range.Font.set({ Size: 22, NameFarEast: '黑体' });
            },
            author(p) {
                p.Alignment = Enum.wdAlignParagraphCenter;
                p.Range.Font.set({ Size: 14, NameFarEast: '仿宋' });
            },
            address(p) {},
            abstract(p) {
                p.Range.InsertBefore('摘要  ');
                p.Range.Font.set({ Size: 10.5 });
                p.Range.Words.Item(1).Font.set({ NameFarEast: '黑体' });
            },
            keywords(p) {
                p.Range.InsertBefore('关键词  ');
                p.Range.Font.set({ Size: 10.5 });
                p.Range.Words.Item(1).Font.set({ NameFarEast: '黑体' });
            },
            h(p) {
                p.Range.Font.set({ Size: 14 });
            },
            hh(p) {
                p.Range.Font.set({ Size: 10.5, NameFarEast: '黑体' });
            },
            hhh(p) {
                p.Range.Font.set({ Size: 10.5, NameFarEast: '黑体' });
            },
            hhhh(p) {},
            p(p) {
                p.Range.Font.set({ Size: 10.5 });
                p.set({ CharacterUnitFirstLineIndent: 2 });
                operator_parser(p.Range);
            },
            link(p) {},
            table(p) {
                p.Range.Font.set({ Size: 9 });
            },
            figure(p) {
                p.Range.Font.set({ Size: 7.5 });
            },
            cite(p) {
                p.Range.Font.set({ Size: 9 });
            },
        };
        const operator_fns = {
            '\\'(range) {
                range.Font.Italic = !0;
            },
            _(range) {
                range.Font.Subscript = !0;
            },
            '^'(range) {
                range.Font.Superscript = !0;
            },
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
            new Range(range).progressive_search('Sentences', (e, i, arr) => {
                if ($.is_null_range(e)) {
                    return;
                }
                if (operators.some((operator) => e.Text === operator)) {
                    // e: 操作符
                    const next_e = arr.Item(i + 1);
                    if ($.is_null_range(next_e)) {
                        return;
                    }
                    const previous_e = arr.Item(i - 1);
                    if (previous_e?.Text === '\\') {
                        return;
                    }
                    if (
                        e.Text === '\\' &&
                        operators.some((operator) => next_e.Text[0] === operator)
                    ) {
                        // e: 转义符
                    } else {
                        operator_fns[e.Text](next_e);
                    }
                    e.Text = '';
                    return;
                }
                if (operators.some((operator) => e.Text.includes(operator))) {
                    // e: 含有操作符的句子/单词
                    return true;
                }
            });
        }
        function main() {
            // 替换特殊符号
            $.replace_all(
                special_symbols.map((e) => e[0]),
                special_symbols.map((e) => e[1]),
            );
            // 构建list
            sel().ClearFormatting();
            const list_num = doc().Lists.Count;
            paragraphs_operator.apply({
                h(p) {
                    p.Style = '标题 1';
                },
                hh(p) {
                    p.Style = '标题 2';
                },
                hhh(p) {
                    p.Style = '标题 3';
                },
                hhhh(p) {
                    p.Style = '标题 4';
                },
                p(p) {
                    p.Style = '正文';
                },
            });
            if (doc().Lists.Count === list_num + 1) {
                //构建成功
                const list = doc().Lists.Item(list_num + 1);
                list.ApplyListTemplate(wps.ListGalleries.Item(3).ListTemplates.Item(7)); //多级编号
            }
            // 识别标识符
            new Range(sel().Range).set_default_text_format();
            paragraphs_operator.apply(identifier_fns);
        }
        main();
    },
};
const cite = (function () {
    /** @typedef {{range: Wps.WpsRange, segments: string[]}} reference*/
    /** @type {Map<reference, string[]>} */
    const rbCite_map = new Map();
    return {
        check_cites() {
            if ($.is_null_range(sel().Range)) {
                alert('请先选中参考文献');
                return;
            }
            //常量
            const reg_cn_author =
                    /(^\[?[\u4e00-\u9fa5]{2,4}(, [\u4e00-\u9fa5]{2,4}){0,5}(, (\.\.\. )?[\u4e00-\u9fa5]{2,4})?\. $)/,
                reg_en_author =
                    /(^[a-zA-Z\-\']+,( [A-Z]\.){1,3}(, [a-zA-Z\-\']+,( [A-Z]\.){1,3}){0,5}(, (\&|\.\.\.) [a-zA-Z\-\']+,( [A-Z]\.){1,3})? $)/,
                reg_author = new RegExp(
                    '^(' + Reg2Str(reg_cn_author) + '|' + Reg2Str(reg_en_author) + ')',
                ),
                reg_date = /(\(\d{4}[a-z]?\)\. )/, //(1234b). /
                reg_jo_title = /[^\(\)]*. /,
                reg_th_title =
                    /[^]*[a-zA-Z\u4e00-\u9fa5] \(((硕|博)士学位论文|Unpublished (master\'s thesis|doctorial dissertation))\). /,
                reg_title = /([^]+[\.\?]|[^]+\? [^]+\.) /,
                reg_jo_source =
                    /([a-zA-Z\u4e00-\u9fa5,\-\&\:\(\)\' ]{2,}, \d+(\(\d+(\-\d+)?\))?, (Article )?[a-zA-Z]?\d+([\-\+][a-zA-Z]?\d+){0,2}\. \r*)/,
                reg_th_source =
                    /(^[a-zA-Z\u4e00-\u9fa5\' ]{4,}(, [a-zA-Z\u4e00-\u9fa5\' ]{2,}){0,3}\. \r*$)/,
                reg_source = new RegExp(
                    '^(' + Reg2Str(reg_jo_source) + '|' + Reg2Str(reg_th_source) + ')]?\\r*',
                ),
                link = /(^(https?|doi)\:\/\/[a-zA-Z\#\?\.\/]*\.\s*)?$/,
                reg_cite = new RegExp(Reg2Str(reg_author, reg_date, reg_title, reg_source));
            //功能组
            const hasLink = (str) => ['doi', 'http'].some((e) => str.includes(e));
            function Reg2Str(...regs) {
                return regs
                    .map((e) => (typeof e == 'string' ? e : e.toString().slice(1, -1)))
                    .join('');
            }
            /**
             * @param {Wps.WpsRange|Wps.WpsRange[]} range
             * @param {RegExp} reg
             */
            function RegExpTest(range, reg, shouldMark = true) {
                const ranges = [range].flat();
                const isOK = reg.test(ranges.map((e) => e.Text).join(''));
                if (!shouldMark) {
                    return isOK;
                }
                if (isOK) {
                    ranges.forEach((e) => {
                        e.Font.Color = Enum.wdColorBlack;
                    });
                } else {
                    ranges.forEach((e) => {
                        e.Font.Color = Enum.wdColorRed;
                    });
                }
                return isOK;
            }
            function modify_full_text() {
                //纠正符号
                $.replace_all('（）【】，−－–：’？！ ．', "()[],---:'?! .");
                //删除多余空格/换行
                sel().Text = sel()
                    .Text.replace(/ {2,}/g, ' ')
                    .replace(/\r{2,}/g, '\r');
            }
            function modify_paragraph_text(p_range) {
                function add_space(collection) {
                    let isExecute = true;
                    return new Collection(collection)
                        .map((e) => {
                            if (hasLink(e.Text)) {
                                if (e.Words.Count === 1) {
                                    isExecute = false;
                                } else {
                                    return add_space(e.Words);
                                }
                            }
                            return isExecute ? e.Text.replaceAll(/[',\.&:?…!']/g, '$& ') : e.Text;
                        }, true)
                        .join('');
                }
                p_range.obj.Text = add_space(p_range.obj.Sentences) //重新赋值，需要重设段落格式
                    .replaceAll(/ {2,}/g, ' ')
                    .replaceAll('. ,', '.,')
                    .replaceAll('. . .', '...')
                    .replaceAll(/[\r\n\f\v]{2,}/g, '\r');
            }
            /**
             * @param {Range} p_range
             */
            function cut_paragraph(p_range) {
                let authors, date, title, source; //各成分
                const sentences = new Collection(p_range.obj.Sentences);
                let last_marker = sentences.obj.Count;
                if (last_marker < 4) {
                    return '这可能是不完整的索引';
                }
                if (hasLink(sentences.at(-1).Words.Item(1).Text)) {
                    //最后一句是否为网址
                    last_marker--;
                }
                source = sentences.at(last_marker - 1);
                for (let i = 1; i < last_marker; i++) {
                    const date_range = sentences.at(i - 1);
                    if (!RegExpTest(date_range, reg_date, false)) {
                        //匹配日期
                        continue;
                    }
                    date = date_range;
                    authors = sentences.slice(0, i - 1);
                    title = sentences.slice(i, last_marker - 1);
                    switch (0) {
                        case authors.length: {
                            return '程序无法获取索引作者';
                        }
                        case title.length: {
                            return '程序无法获取索引标题';
                        }
                    }
                    break;
                }
                if (date && authors && title) {
                    /**@type {[Wps.WpsRange[],Wps.WpsRange,Wps.WpsRange[],Wps.WpsRange]} */
                    const data = [authors, date, title, source];
                    return data;
                } else {
                    return '程序无法获取索引日期';
                }
            }
            /**
             * @param {Range} p_range
             */
            function set_paragraph(p_range) {
                modify_paragraph_text(p_range);
                //分段
                const p_segments = cut_paragraph(p_range);
                if (!(p_segments instanceof Array)) {
                    p_range.obj.Font.Color = Enum.wdColorBlue;
                    p_range.add_comment(p_segments + '\v需要人工检查');
                    return;
                }
                rbCite_map.set(
                    {
                        range: p_segments[1],
                        segments: p_segments.map((e) =>
                            [e]
                                .flat()
                                .map((e) => e.Text)
                                .join(''),
                        ),
                    },
                    [],
                );
                //检测类型：期刊文献，学位论文
                const [authors, date, title, source] = p_segments;
                const isTh =
                    /((硕|博)士学位论文|Unpublished (master\'s thesis|doctorial dissertation))/.test(
                        p_range.obj.Text,
                    );
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
                //段落格式
                p_range.set_paragraph_format({ CharacterUnitFirstLineIndent: -2 });
            }
            function set_paragraphs() {
                const paragraphs = new Collection(sel().Paragraphs);
                paragraphs.map((p) => {
                    if ($.is_null_range(p.Range)) {
                        //跳过空行
                        return;
                    }
                    if ($.has_special_char(p.Range.Text)) {
                        //跳过含有特殊字符的段落
                        return;
                    }
                    const p_range = new Range(p.Range);
                    try {
                        set_paragraph(p_range);
                    } catch (e) {
                        p_range.add_comment('意料外的错误\v' + e, 'err');
                        console.error(e);
                    }
                });
                alert('检查成功\n记录了' + rbCite_map.size + '条参考文献');
            }
            function timer(fn) {
                const st = +new Date();
                fn();
                return +new Date() - st;
            }
            function main() {
                rbCite_map.clear();
                modify_full_text();
                const range = new Range(sel().Range);
                range.del_comment();
                range.set_font_format({ Size: 9 });
                console.log('每秒词数：' + (sel().Words.Count * 1e3) / timer(set_paragraphs));
            }
            main();
        },
        match_cites() {
            if ($.is_null_range(sel().Range)) {
                alert('请先选中正文');
                return;
            }
            if (rbCite_map.size === 0) {
                alert('还没有记录参考文献\n请先检查索引');
                return;
            }
            $.replace_all('（）', '()');
            //提取索引
            /**@type {string}*/
            const text = sel().Text;
            const refer_cites = [...rbCite_map.keys()];
            const body_cites = [
                ...text.matchAll(
                    /[^a-zA-Z\u4e00-\u9fa5\(\)]{1}[a-zA-Z\u4e00-\u9fa5]+\(\d{4}[a-z]?\)/g,
                ), //。abc(2333a)
                ...text.matchAll(/\([^\(\)]+\d{4}[a-z]?\)/g), //(abc et al., 2333a)
            ].flatMap((e) => e[0].split(';'));
            //功能组
            function main() {
                const success_body_cites = filter_by_array(
                    body_cites,
                    get_refer_infos(refer_cites),
                    (cite, [ref, info]) => {
                        const isOK = info.every((e) => cite.includes(e));
                        if (isOK) {
                            rbCite_map.get(ref).push(cite);
                        }
                        return isOK;
                    },
                );
                const fail_body_cites = diff(body_cites, success_body_cites);
                addComments(body_cites, fail_body_cites);
                console.log('正文索引', body_cites);
                console.log('匹配结果', [...rbCite_map.entries()]);
            }
            /**
             * @param {string[]} body_cites
             * @param {string[]} fail_body_cites
             */
            function addComments(body_cites, fail_body_cites) {
                //正文批注
                const sel_range = new Range(sel().Range);
                sel_range.del_comment();
                if (body_cites.length === 0) {
                    sel_range.add_comment('这段正文中好像没有索引');
                    return;
                }
                sel_range.add_comment('参考文献核对情况见参考文献列表\v没有批注则核对成功', 'info');
                fail_body_cites.length
                    ? sel_range.add_comment(
                          '这些正文索引缺少参考文献\v' + fail_body_cites.join('\v'),
                      )
                    : sel_range.add_comment('正文核对成功', 'ok');
                //参考文献批注
                refer_cites.forEach((ref) => {
                    if (rbCite_map.get(ref).length !== 0) {
                        return;
                    }
                    const ref_range = new Range(ref.range);
                    ref_range.add_comment('这可能是未被引用的文献\v需要人工核对');
                    //再次查找符合该ref的cites
                    const maybe_body_cites = filter_by_array(
                        body_cites,
                        get_refer_infos([ref]),
                        (cite, [ref, info]) => info.every((e) => cite.includes(e)),
                    );
                    if (maybe_body_cites.length !== 0) {
                        ref_range.add_comment(
                            '找到一些可能对应的正文索引\v' + maybe_body_cites.join('\v'),
                            'info',
                        );
                    }
                });
            }
            /**
             * 基于 condition_arr 中的元素筛选 target_arr 中的元素
             * @template T,R
             * @param {T[]} target_arr
             * @param {R[]} condition_arr
             * @param {(target:T,condition:R)=>boolean} fn
             */
            function filter_by_array(target_arr, condition_arr, fn) {
                const filtered_arr = [];
                loop_target: for (let target of target_arr) {
                    loop_condition: for (let condition of condition_arr) {
                        if (fn(target, condition)) {
                            filtered_arr.push(target);
                            continue loop_target;
                        }
                    }
                }
                return filtered_arr;
            }
            /**
             * 求数组差集
             * @template T
             * @param {T[]} target
             * @param {T[]} arr
             */
            function diff(target, arr) {
                const set = new Set(arr);
                return [...new Set(target)].filter((x) => !set.has(x));
            }
            /**
             * 获取参考文献关键信息
             * @param {reference[]} refers
             * @returns {[reference, [string,string]][]}
             */
            function get_refer_infos(refers) {
                return refers.map((ref) => {
                    const author = ref.segments[0].match(/^([^,\.]+)/)?.[1], //abc
                        date = ref.segments[1].match(/^\((\d{4}[a-z]?)\)/)?.[1]; //2333a
                    if (!author) throw '无法获取参考文献的作者';
                    if (!date) throw '无法获取参考文献的日期';
                    return [ref, [author, date]];
                });
            }
            main();
        },
    };
})();
const link = (function () {
    let cur_bookmark;
    const get_sel_range = () => new Range(sel().Range);
    return {
        create_link() {
            cur_bookmark?.Delete(); //删除之前没被引用的书签
            if (sel().Type === Enum.wdSelectionNormal) {
                cur_bookmark = get_sel_range().add_bookmark();
                alert('创建成功');
            } else {
                alert('请先选中文本');
            }
        },
        insert_link() {
            if (cur_bookmark) {
                get_sel_range().cite_bookmark(cur_bookmark);
                cur_bookmark = null;
            } else {
                alert('请先创建文本链接');
            }
        },
    };
})();
const app_step = other.set(cite, link);
/**
 * 添加撤销组
 * @this {()=>void}
 * @param  {...any} args
 * @returns
 */
function step(...args) {
    wps.WpsApplication().UndoRecord.StartCustomRecord('论文_' + this.name);
    try {
        this.apply(null, args);
    } catch (e) {
        new Range(sel().Range).add_comment('意料外的错误\v请撤销这个操作\v' + e, 'err');
        console.error(e);
    }
    wps.WpsApplication().UndoRecord.EndCustomRecord();
    return !0;
}
app_step.each((value, key) => {
    app_step[key] = step.bind(app_step[key]);
});
export default app_base.set(app_ui, app_step);
