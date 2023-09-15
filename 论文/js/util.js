const sel = () => wps.Selection;
const doc = () => wps.ActiveDocument;
const Enum = wps.Enum;
const $ = {
    version: 'v23.9-alpha.4',
    rbCite_pairs: new Map(),
    comment_Author: '＞︿＜',
};
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
/**
 * 同时修改多个属性
 */
Object.prototype.set = function (...configs) {
    return Object.assign(this, ...configs);
};
Function.prototype.set({
    /**
     * 添加撤销组
     */
    step(...args) {
        wps.UndoRecord.StartCustomRecord('JSA-' + this.name);
        try {
            this(...args);
        } catch (e) {
            new Range_decorator(sel().Range).add_comment('意料外的错误\v请撤销刚刚的操作\v' + e);
            console.error(e);
        }
        wps.UndoRecord.EndCustomRecord();
        return !0
    },
    /**
     * 执行计时
     */
    timed(...args) {
        const st = +new Date();
        this(...args);
        return +new Date() - st;
    },
});
function open_url_in_local(url) {
    wps.OAAssist.ShellExecute(url);
}
function open_url_in_wps(url, caption, width, height) {
    wps.ShowDialog(url, caption, width, height, true);
}
function is_null_range(range) {
    return range.Text.length === 1 && '\r\v '.includes(range.Text);
}
/**
 * 批量替换文本
 * @param {Iterable<string>} oldTexts 
 * @param {Iterable<string>} newTexts 
 */
function replaceAll(oldTexts, newTexts, range = sel().Range) {
    if (oldTexts.length != newTexts.length) {
        throw '查找数组和替换数组的长度不匹配';
    }
    for (let i = 0; i < oldTexts.length; i++) {
        range.Find.Execute(oldTexts[i], true, true, false, false, false, true, Enum.wdFindContinue, false, newTexts[i], Enum.wdReplaceAll);
    }
}
