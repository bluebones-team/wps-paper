import { doc, Enum } from './util.js';
import { config } from '../config.js';
/**
 * 自定义集合
 * @template {{Count:number,Item(index:number):ReturnType<T['Item']>}} T
 */
class Collection {
    /**
     * @param {T} obj
     */
    constructor(obj) {
        this.obj = obj;
    }
    /**
     * @template R
     * @param {(item:ReturnType<T['Item']>,index:number,obj:T)=>R} fn
     * @param {boolean} order true：正序
     */
    map(fn, order = false) {
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
        if (index >= 0 && index < len) return this.obj.Item(index + 1);
        if (index < 0 && index >= -len) return this.obj.Item(len + index + 1);
        throw '索引超过集合长度';
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
    /**
     * 迭代器
     */
    *[Symbol.iterator]() {
        const len = this.obj.Count;
        for (let i = 1; i <= len; i++) {
            // 每次获取一次
            yield this.obj.Item(i);
        }
    }
}
const comment_values = new Set(Object.values(config.comment));
/**自定义Range */
class Range {
    /**
     * @param {Wps.WpsRange} obj
     */
    constructor(obj) {
        this.obj = obj;
    }
    /**
     * @param {string} content
     */
    add_comment(content, type = 'warn') {
        return this.obj.Comments.Add(this.obj, content).set({
            Author: config.comment[type],
        });
    }
    del_comment() {
        new Collection(this.obj.Comments).map((e) => {
            if (comment_values.has(e.Author)) {
                e.Delete();
            }
        });
    }
    add_bookmark(name = `_wps_paper_link_${+new Date()}`) {
        return doc().Bookmarks.Add(name, this.obj);
    }
    /**
     * @param {Wps.WpsBookmark} bookmark
     */
    cite_bookmark(bookmark) {
        const field = doc().Fields.Add(this.obj);
        field.Code.Text = ` REF ${bookmark.Name} \\h `;
        field.Update();
    }
    /**
     * @param {Partial<Wps.WpsFont>} config
     */
    set_font_format(config = {}) {
        this.obj.HighlightColorIndex = Enum.wdNoHighlight;
        this.obj.Font.Reset();
        this.obj.Font.set({
            Bold: 0,
            Italic: 0,
            Color: Enum.wdColorAutomatic, //自动配色
            Name: '',
            NameAscii: 'Times New Roman',
            NameFarEast: '宋体',
            NameOther: 'Times New Roman',
        }).set(config);
    }
    set_paragraph_format(config = {}) {
        this.obj.Paragraphs.set({
            Alignment: Enum.wdAlignParagraphJustify,
            CharacterUnitLeftIndent: 0, //左缩进量
            CharacterUnitRightIndent: 0,
            CharacterUnitFirstLineIndent: 0, //首行缩进
            SpaceAfter: 0, //段后间距
            SpaceBefore: 0,
            LineUnitAfter: 0, //段后间距（网格线）
            LineUnitBefore: 0,
            LineSpacingRule: Enum.wdLineSpace1pt5, //1.5倍行距
            AutoAdjustRightIndent: 0, //不自动调整右缩进
            DisableLineHeightGrid: 1, //不与网格线对齐
        }).set(config);
    }
    set_default_text_format() {
        this.set_font_format();
        this.set_paragraph_format();
    }
    /**
     * 递进查找
     * @param {'Sentences'|'Words'|'Characters'} name 起始集合
     * @param {(e:Wps.WpsRange,i:number,arr:Wps.WpsSentences | Wps.WpsWords | Wps.WpsCharacters) => boolean|void} fn
     */
    progressive_search(name, fn) {
        const names = ['Sentences', 'Words', 'Characters'];
        !(function recur(collection, recur_num) {
            return new Collection(collection).map((e, i, arr) => {
                if (fn(e, i, arr)) {
                    recur(e[names[recur_num + 1]], recur_num + 1);
                }
            });
        })(this.obj[name], names.indexOf(name));
    }
}
export { Collection, Range };
