function Expand(locked, p, o) {
    if (locked) {
        // 添加不可删除、不可修改、不可枚举的属性
        Object.keys(o).forEach(k => o[k] = { value: o[k] });
        Object.defineProperties(p, o);
    } else {
        Object.assign(p, o);
    }
};
const { toFixed } = Number.prototype;
const { insertBefore, appendChild } = Node.prototype;
Expand(0, Number.prototype, {
    toFixed(n) {
        return toFixed.call(+`${Math.round(`${this}e${n}`)}e-${n}`, n)
    },
});
Expand(1, Object.prototype, {
    or(k, a, b) {
        const L = this[k] === a;
        this[k] = L ? b : a;
        return L
    },
    each(fn) {
        for (let key in this) {
            if (Object.hasOwnProperty.call(this, key)) {
                fn(this[key], key, this);
            }
        }
    },
});
Expand(1, Array.prototype, {
    mean() {
        const obj = {};
        this.forEach(e => {
            e.each((value, key) => {
                obj[key] = (obj[key] ?? 0) + value;
            });
        });
        obj.each((v, k) => { obj[k] = v / this.length });
        return obj;
    },
    choose(size) {
        var allResult = [];
        (function (arr, size, result) {
            var arrLen = arr.length;
            if (size > arrLen) return;
            if (size == arrLen) {
                allResult.push([].concat(result, arr));
            } else {
                for (var i = 0; i < arrLen; i++) {
                    var newResult = [].concat(result);
                    newResult.push(arr[i]);
                    if (size == 1) {
                        allResult.push(newResult);
                    } else {
                        var newArr = [].concat(arr);
                        newArr.splice(0, i + 1);
                        arguments.callee(newArr, size - 1, newResult);
                    }
                }
            }
        })(this, size, []);
        return allResult;
    }
});
Expand(1, HTMLInputElement.prototype, {
    add(max = 4) {
        const tree = this.parentElement.parentElement;
        const source = tree.lastElementChild;
        if (tree.childElementCount < max) {
            function selfAdd(attribute) {
                if (attribute) {
                    const index = parseInt(attribute.slice(-1));
                    if (!isNaN(index)) return attribute.slice(0, -1) + (index + 1);
                }
                return ''
            }
            tree.appendChild(Object.assign(source.cloneNode(true), {
                id: selfAdd(source.id),
                placeholder: selfAdd(source.placeholder),
                title: selfAdd(source.title),
                value: selfAdd(source.value),
            }));
        }
    },
    dec(min = 2) {
        const tree = this.parentElement.parentElement;
        if (tree.childElementCount == 3 && /^indep\d$/.test(tree.id)) tree.appendChild = e => e;
        if (tree.childElementCount > min) tree.lastElementChild.remove();
    },
    del() {
        this.parentElement.remove();
    },
    info() { },
});
Expand(0, HTMLDivElement.prototype, {
    init() {
        this.Mode = Mode[db.mode];
        this.editor = this.addEditor(this.onCheck, this.onEdit);
        // 更新元素
        const newHTML = this.update?.(); if (newHTML) this.$('.data').innerHTML = newHTML;
        // label 对齐
        const list = this.$('label', 1),
            max = Math.max(...list.map(e => parseFloat(getComputedStyle(e).width)));
        list.forEach(e => e.style.width = max + 'px');
        // 最后执行
        this.initFunc?.();
        return this;
    },
    datas() { return [...this.$(`.data`).children] },
    dataElm(id, exp = 'name|placeholder|type|other,', config = {}) {
        const propNames = exp.split(',').map(e => e.split('|'));
        const inputHTMLs = propNames.map(e =>
            `<input id="${id}_${e[0]}" type="${e[2] || 'number'}" placeholder="${e[1] || e[0]}" ${e[3] || ''}>`
        ).join('');
        const dataElm = addElm(config.tag || 'div', {
            id: `${id}_data`,
            innerHTML: inputHTMLs,
        }, config);
        // 插入元素 -> 父容器
        config.elms?.forEach(e => {
            typeof e.index == 'number' ?
                dataElm.children[e.index].outerHTML += e.elm :
                dataElm.innerHTML = e.elm + dataElm.innerHTML;
        });
        // 添加附件
        if (config.label) dataElm.innerHTML = `<label>${config.label}</label>` + dataElm.innerHTML;
        if (config.hasDel) dataElm.innerHTML += '<input class="del tool" type="button" value="×" onclick="this.del()">';
        return dataElm.outerHTML;
    },
    readTextareaData() {
        const textarea = this.editor.$('textarea');
        const text = textarea.value.replace(/ +/g, '').replace(/\n$/, ''),
            arr = text.split('\n').map(e => e.split('\t'));
        textarea.value = '';
        if (arr[1]) {
            const list = this.datas().map(e => e.$('input[placeholder]', 1));
            arr.forEach((e, r) => {
                e.forEach((ee, c) => { if (list[r]?.[c]) list[r][c].value = ee; });
            });
        }
    },
    writeData(dataElm, prop) {
        // 获取id
        let id = dataElm.id.split('_')[0];
        if (id === '' && dataElm.className == 'data') {
            id = dataElm.parentElement.id.split('_')[0];
        }
        if (id === '') {
            throw `提取id为空字符 ${dataElm.outerHTML.slice(0, 50)}`;
        }
        // 获取obj
        let obj = window[id] ||= { name: id };
        if (typeof prop == 'string') {
            // 多个元素 写入数组
            !function set(propNames) {
                const name = propNames.shift();
                const hasElm = !!propNames.length;
                obj = obj[name] = hasElm ? {} : [];
                hasElm && set(propNames);
            }(prop.split('.'));
            Object.assign(obj, dataElm.$('div[id]', 1).map(elm2obj));
        } else {
            // 单个元素 写入对象
            Object.assign(obj, elm2obj(dataElm));
        }
        function elm2obj(elm) {
            // 根据单个元素 返回数据对象
            const obj = {};
            elm.$('input[id]', 1).forEach(e => {
                obj[e.id.match(/_([^_]+)$/)[1]] = e.type == 'number' ? +e.value : e.value;
            });
            return obj;
        }
        return window[id];
    },
    onCheck() {
        this.readTextareaData();
        const colors = isOK => ['#f005', '#0f05'][+isOK];
        const errInput = this.$('input[placeholder]', 1).filter(e => {
            e.style.backgroundColor = colors(e.pass = true);
            return e.value === '';
        });
        errInput[0]
            ? errInput.forEach(e => e.style.backgroundColor = colors(e.pass = false))
            : (this.confirm(), console.log(this.id + '\n', this.output()));
    },
    onEdit() {
        this.$('input[placeholder]', 1).forEach(e => {
            e.style.backgroundColor = e.pass = '';
        });
    },
    output() { return '' },
});
Expand(1, Node.prototype, {
    $: _$,
    addTools(max = 9, min = 1) {
        return this.insertBefore(addElm('span', {
            id: `${this.id}_tools`,
            className: 'tools',
            innerHTML: `<input class="add tool" type="button" value="+" onclick="this.add(${max + 1})">
                        <input class="dec tool" type="button" value="-" onclick="this.dec(${min + 1})">`,
        }), this.children[0]);
    },
    addEditor(func1, func2) {
        const _this = this;
        const elm = this.appendChild(addElm('span', {
            id: `${this.id}_editor`,
            className: 'editor',
            innerHTML: `<input type="button" value="确认">
                       <!-- <input type="button" value="i" class="info tool" onclick="this.info()"> -->
                        <textarea cols="2" rows="1"></textarea>`,
        }));
        elm.$('input[value=确认]').onclick = function () {
            _this.$('input', 1).filter(e => e !== this).forEach(e => e.disabled = !e.disabled);
            elm.$('textarea').or('disabled', true, false);
            elm.$('textarea').style.or('resize', '', 'none');
            [func1, func2][+!this.or('value', '确认', '编辑')]?.call(_this);
            //自动输出
            if (db.autoOutput && $('script ~.box:not(#output)', 1).every(e =>
                e.$('input[placeholder]', 1).every(i => i.pass))) Mode.output();
        };
        return elm;
    },
    insertBefore(node, pos) {
        const o = node.id && $(`#${node.id}`);
        return o
            ? (console.log(`已存在`, o), o)
            : insertBefore.call(this, node, pos);
    },
    appendChild(node) {
        const o = node.id && $(`#${node.id}`);
        return o
            ? (console.log(`已存在`, o), o)
            : appendChild.call(this, node);
    },
});