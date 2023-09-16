import { sel, doc, Enum, $ } from "./util.js";
class Decorator {
    constructor(obj) {
        this.obj = obj;
    }
}
class Collection_decorator extends Decorator {
    /**
    * @param {boolean} order true：正序
    */
    map(fn, order) {
        const len = this.obj.Count;
        const arr = [];
        for (let i = len; i > 0; i--) {
            const index = order ? len - i + 1 : i;
            arr[index - 1] = fn(this.obj.Item(index), index, this.obj);
        }
        return arr;
    }
    /**
     * @param {number} index 0：第一个item
     */
    at(index) {
        const len = this.obj.Count;
        if (index >= 0 && index < len)
            return this.obj.Item(index + 1);
        if (index < 0 && index >= -len)
            return this.obj.Item(len + index + 1);
    }
    /**
     * 数组切片
     */
    slice(start = 0, end = this.obj.Count) {
        const arr = [];
        for (let i = start; i < end; i++) {
            arr.push(this.obj.Item(i + 1));
        }
        return arr;
    }
}
class Range_decorator extends Decorator {
    add_comment(content, type = 'warn') {
        return this.obj.Comments.Add(this.obj, content).set({ Author: $.comment_Authors[type] });
    }
    del_comment() {
        new Collection_decorator(this.obj.Comments).map(e => {
            if ($.comment_Authors.includes(e.Author)) {
                e.Delete();
            }
        });
    }
    add_bookmark(name = `_link_${+new Date()}`) {
        return doc().Bookmarks.Add(name, this.obj);
    }
    cite_bookmark(bookmark) {
        const field = doc().Fields.Add(this.obj);
        field.Code.Text = ` REF ${bookmark.Name} \\h `;
        field.Update();
    }
    set_font_format(config = {}) {
        this.obj.HighlightColorIndex = Enum.wdNoHighlight;
        this.obj.Font.Reset();
        this.obj.Font.set({
            Bold: false,
            Italic: false,
            Color: Enum.wdColorAutomatic, //自动配色
            Name: '',
            NameAscii: "Times New Roman",
            NameFarEast: "宋体",
            NameOther: "Times New Roman",
        }, config);
    }
    set_paragraph_format(config = {}) {
        this.obj.Paragraphs.set({
            CharacterUnitLeftIndent: 0, //左缩进量
            CharacterUnitRightIndent: 0,
            CharacterUnitFirstLineIndent: 0, //首行缩进
            SpaceAfter: 0, //段后间距
            SpaceBefore: 0,
            LineUnitAfter: 0, //段后间距（网格线）
            LineUnitBefore: 0,
            LineSpacingRule: Enum.wdLineSpace1pt5, //1.5倍行距
            AutoAdjustRightIndent: false, //不自动调整右缩进
            DisableLineHeightGrid: true, //不与网格线对齐
        }, config);
    }
    set_default_text_format() {
        this.set_font_format();
        this.set_paragraph_format();
    }
    /**
     * 递进查找
     * @param {'Paragraphs'|'Sentences'|'Words'|'Characters'} name 起始集合
     * @param {(e,i,arr) => boolean} fn
     */
    progressive_search(name, fn) {
        const names = ['Paragraphs', 'Sentences', 'Words', 'Characters'];
        !function recur(collection, recur_num) {
            new Collection_decorator(collection).map((e, i, arr) => {
                if (fn(e, i, arr)) {
                    recur(e[names[recur_num + 1]], recur_num + 1);
                }
            });
        }(this.obj[name], names.indexOf(name));
    }
}
export { Collection_decorator, Range_decorator }