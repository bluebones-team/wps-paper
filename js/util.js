export const sel = () => wps.Selection;
export const doc = () => wps.ActiveDocument;
export const Enum = window.wps?.Enum;
export const $ = {
    open_url_in_local(url) {
        if (url[0] === '/') {
            url = location.origin + url;
        }
        wps.OAAssist.ShellExecute(url);
    },
    open_url_in_wps(url, caption, width, height) {
        wps.ShowDialog(url, caption, width, height, true);
    },
    is_null_range(range) {
        const char_num = range.Text.length;
        switch (char_num) {
            case 0: return true;
            case 1: return '\r\v\f '.includes(range.Text);
            default: return false;
        }
    },
    has_special_char(text) {
        return wps.CleanString(text) !== text
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
