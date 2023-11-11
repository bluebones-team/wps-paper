export const sel = () => wps.WpsApplication().Selection;
export const doc = () => wps.WpsApplication().ActiveDocument;
export const Enum = wps.Enum;
export const $ = {
    /**
     * @param {Wps.WpsRange} range 
     */
    is_null_range(range) {
        const char_num = range.Text.length;
        switch (char_num) {
            case 0: return true;
            case 1: return '\r\v\f '.includes(range.Text);
            default: return false;
        }
    },
    /**
     * @param {string} text 
     */
    has_special_char(text) {
        return wps.WpsApplication().CleanString(text) !== text
    },
    /**
     * 批量替换文本
     * @param {string|string[]} oldTexts 
     * @param {string|string[]} newTexts 
     */
    replace_all(oldTexts, newTexts, range = sel().Range) {
        if (oldTexts.length != newTexts.length) {
            throw '查找数组和替换数组的长度不匹配';
        }
        for (let i = 0; i < oldTexts.length; i++) {
            range.Find.Execute(oldTexts[i], true, true, false, false, false, true, Enum.wdFindContinue, false, newTexts[i], Enum.wdReplaceAll);
        }
    },
};
Object.prototype.set = function (...configs) {
    return Object.assign(this, ...configs);
};
Object.prototype.each = function (fn) {
    for (const key in this) {
        if (Object.hasOwnProperty.call(this, key)) {
            fn(this[key], key, this);
        }
    }
};