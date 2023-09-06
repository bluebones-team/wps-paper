const sel = () => wps.Selection;
const doc = () => wps.ActiveDocument;
const Enum = wps.Enum;
const $ = {
    version: 'v23.9-alpha.3',
    refer_cites: [],
    comments: [],
};
/**
 * 同时修改多个属性
 */
Object.prototype.set = function (...configs) {
    return Object.assign(this, ...configs);
};
/**
 * 数组去重，包括对象数组
 */
Array.prototype.deduplication = function () {
    let arr = this.map(e => JSON.stringify(e));
    arr = [...new Set(arr)]; //去重
    return arr.map(e => JSON.parse(e));
}
/**
 * 添加撤销组
 */
Function.prototype.step = function (note) {
    note ??= this.name;
    return (...args) => {
        wps.UndoRecord.StartCustomRecord(`JSA-${note}`);
        try {
            this(...args);
        } catch (err) {
            alert(err + '');
            console.error(err);
        }
        wps.UndoRecord.EndCustomRecord();
        return !0
    }
};
function open_url_in_local(url) {
    wps.OAAssist.ShellExecute(url);
}
function open_url_in_wps(url, caption, width, height) {
    wps.ShowDialog(url, caption, width, height, true);
}
function del_comment(range) {
    collection_operator(range.Comments).map(e => $.comments.includes(e) && e.Delete());
}
function add_comment(range, content) {
    del_comment(range);
    const comment = range.Comments.Add(range, content);
    comment.Author = '*Output';
    $.comments.push(comment);
}
function set_font_format(Font, config = {}) {
    Font.Reset();
    Font.set({
        Bold: !1,
        Italic: !1,
        Color: 0,
        Name: '',
        NameAscii: "Times New Roman",
        NameFarEast: "宋体",
        NameOther: "Times New Roman",
    }, config);
}
function set_paragraph_format(Paragraphs, config = {}) {
    Paragraphs.set({
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
function add_bookmark(range, name = `_link_${+new Date()}`) {
    return doc().Bookmarks.Add(name, range);
}
function cite_bookmark(bookmark, range = sel().Range) {
    const field = doc().Fields.Add(range);
    field.Code.Text = ` REF ${bookmark.Name} \\h `;
    field.Update();
}
/**
 * 返回集合对象操作器
 */
function collection_operator(collection) {
    return {
        obj: collection,
        /**
         * @param {boolean} order true：正序
         */
        map(fn, order) {
            const len = collection.Count;
            const arr = [];
            for (let i = len; i > 0; i--) {
                const index = order ? len - i + 1 : i;
                arr[index - 1] = fn(collection.Item(index), index, collection);
            }
            return arr;
        },
        /**
         * @param {number} index 0：第一个item
         */
        at(index) {
            const len = collection.Count;
            if (index >= 0 && index < len)
                return collection.Item(index + 1);
            if (index < 0 && index >= -len)
                return collection.Item(len + index + 1);
        },
        /**
         * 数组切片
         */
        slice(start = 0, end = collection.Count) {
            const arr = [];
            for (let i = start; i < end; i++) {
                arr.push(collection.Item(i + 1));
            }
            return arr;
        },
    };
}
/**
 * 批量替换文本
 * @param {Iterable<string>} oldTexts 
 * @param {Iterable<string>} newTexts 
 */
function replaceAll(Range, oldTexts, newTexts) {
    if (oldTexts.length != newTexts.length) {
        throw '查找数组和替换数组的长度不匹配';
    }
    for (let i = 0; i < oldTexts.length; i++) {
        Range.Find.Execute(oldTexts[i], true, true, false, false, false, true, Enum.wdFindContinue, false, newTexts[i], Enum.wdReplaceAll);
    }
}
/**
 * 递进查找
 * @param {'Paragraphs'|'Sentences'|'Words'|'Characters'} name 集合对象名
 * @param {(e,i,arr) => boolean} fn
 */
function progressive_search(range, name, fn) {
    const names = ['Paragraphs', 'Sentences', 'Words', 'Characters'];
    !function recur(collection, recur_num) {
        collection_operator(collection).map((e, i, arr) => {
            if (fn(e, i, arr)) {
                recur(e[names[recur_num + 1]], recur_num + 1);
            }
        });
    }(range[name], names.indexOf(name));
}