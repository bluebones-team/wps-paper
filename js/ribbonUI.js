import { config } from '../config.js'
import app from './ribbon.js'
export default (function (UI) {
    [
        'Label',
        'Screentip',
        'Supertip',
        'fn',
    ].forEach(prop => {
        UI[prop] = function ({ Id }) {
            const value = UI.controls[Id]?.[prop] ?? '';
            return typeof value === 'function' ? value() : value;
        };
    });
    return UI;
})({
    onload: app.onload,
    controls: {
        r1: {
            Label: '检查索引',
            Screentip: '检查参考文献是否按APA格式进行索引',
            Supertip: '请先选中参考文献',
            fn: app.check_cites,
        },
        r2: {
            Label: '核对索引',
            Screentip: '核对正文索引和参考文献列表是否一一对应',
            Supertip: '请先选中正文',
            fn: app.match_cites,
        },
        w1: {
            Label: '撰写结果',
            Screentip: '通过输入统计数据自动生成结果部分',
            Supertip: '单击打开网页程序',
            fn: app.write_result,
        },
        w2: {
            Label: '创建文链',
            Screentip: '创建文本链接',
            Supertip: '请先选中文本',
            fn: app.create_link,
        },
        w3: {
            Label: '插入文链',
            Screentip: '插入文本链接',
            Supertip: '将之前创建的文链插入光标处',
            fn: app.insert_link,
        },
        c1: {
            Label: '作图',
            Screentip: '',
            Supertip: '',
            fn: app.add_figure,
        },
        c2: {
            Label: '作表',
            Screentip: '以三线表格式粘贴表格',
            Supertip: '请先复制表格',
            fn: app.add_table,
        },
        c3: {
            Label: '流程图',
            Screentip: '',
            Supertip: '',
            fn: app.add_flowChart,
        },
        o1: {
            Label: '格式解析',
            Screentip: '对格式文本进行解析',
            Supertip: '请先选中需要解析的文本',
            fn: app.syntax_parser,
        },
        o2: {
            Label: '帮助',
            Screentip: '',
            Supertip: '',
            fn: app.show_help,
        },
        o3: {
            Label: '更新',
            Screentip: '当前版本',
            Supertip: config.version,
            fn: app.update,
        },
    },
});