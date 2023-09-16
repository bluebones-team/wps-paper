export const sel = () => wps.Selection;
export const doc = () => wps.ActiveDocument;
export const Enum = wps.Enum;
export const $ = {
    version: 'v23.9-alpha.5',
    comment_Authors: {
        err: 'Σ(ﾟдﾟ;)',
        warn: '＞︿＜',
        ok: 'ヾ(≧▽≦*)o',
        info: '_(:з」∠)_',
    },
    open_url_in_local(url) {
        wps.OAAssist.ShellExecute(url);
    },
    open_url_in_wps(url, caption, width, height) {
        wps.ShowDialog(url, caption, width, height, true);
    },
    is_null_range(range) {
        const char_num = range.Text.length;
        switch (char_num) {
            case 0: return true;
            case 1: return '\r\v '.includes(range.Text);
            default: return false;
        }
    },
    /**
     * 批量替换文本
     * @param {Iterable<string>} oldTexts 
     * @param {Iterable<string>} newTexts 
     */
    replaceAll(oldTexts, newTexts, range = sel().Range) {
        if (oldTexts.length != newTexts.length) {
            throw '查找数组和替换数组的长度不匹配';
        }
        for (let i = 0; i < oldTexts.length; i++) {
            range.Find.Execute(oldTexts[i], true, true, false, false, false, true, Enum.wdFindContinue, false, newTexts[i], Enum.wdReplaceAll);
        }
    },
};
/**
 * 同时修改多个属性
 */
Object.prototype.set = function (...configs) {
    return Object.assign(this, ...configs);
};
/**
 * 执行计时
 */
Function.prototype.timed = function (...args) {
    const st = +new Date();
    this(...args);
    return +new Date() - st;
}
