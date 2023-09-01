const Enum = wps.Enum;
const $ = {
    refer_cites: [],
    comments: [],
};
/**
 * 同时修改多个属性
 */
Object.prototype.set = function (config = {}) {
    Object.assign(this, config);
    return this;
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
            alert(err.message);
            console.error(err);
        }
        wps.UndoRecord.EndCustomRecord();
        return !0
    }
};
const sel = () => wps.Selection;
const doc = () => wps.ActiveDocument;
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
    Object.assign(Font, {
        Bold: !1,
        Italic: !1,
        Color: 0,
        Name: '',
        NameAscii: "Times New Roman",
        NameFarEast: "宋体",
    }, config);
}
function set_paragraph_format(Paragraph, config = {}) {
    Paragraph.Reset();
    Paragraph.Space15(); //1.5倍行距
    Object.assign(Paragraph, {
        SpaceAfter: 0, //段后间距
        SpaceBefore: 0,
        CharacterUnitLeftIndent: 0, //左缩进量
        CharacterUnitRightIndent: 0,
        CharacterUnitFirstLineIndent: 0, //首行缩进
    }, config);
}
/**
 * 添加 书签-域 链接
 */
function add_BookmarkField_link(source, refer, name) {
    doc().Bookmarks.Add(name, source);
    doc().Fields.Add(refer).Code.Text = ` REF ${name} \\h `;
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
            throw '索引超出长度'
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