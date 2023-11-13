const op = {
    MSD(...level) {
        const { MSDdata } = $('#MSDtable');
        if (!MSDdata) throw '请先填写MSDtable';
        let w = MSDdata.rows.indexOf(level[0]),
            h = MSDdata.columns.indexOf(level[0]);
        if ((w != -1) == (h != -1)) {
            throw (w != -1) ? '水平重复' : `水平缺失-${level[0]}`;
        } else if (w != -1) {
            h = MSDdata.columns.indexOf(level[1] || 'mean');
            if (h == -1) throw `水平缺失-${level[1]}`;
        } else {
            w = MSDdata.rows.indexOf(level[1] || 'mean');
            if (w == -1) throw `水平缺失-${level[1]}`;
        }
        const { M, SD } = MSDdata.data[w][h];
        return {
            M, SD,
            text: `(M = ${M.toFixed(2)}, SD = ${SD.toFixed(2)})`
        };
    },
    ifRes(type, arg1, arg2) {
        switch (type) {
            case 'p***': return arg1 < 0.001 ? 'p < 0.001' : 'p = ' + arg1.toFixed(3);
            case 'p*': return arg1 < 0.05 ? '显著' : '不显著';
            case 'M': return arg1.M > arg2.M ? '大于' : '小于';
            case 'ia': return arg1 + (arg1.includes('×') ? '交互作用' : '主效应');
            case 'orient': return arg1 > 0 ? '正' : '反';
            case '?': return arg2[+arg1];
        }
    },
    sign(p, l1, l2, msd1, msd2) {
        return p < 0.05 ?
            `${l1}的${d1.name}${msd1.text}显著${op.ifRes('M', msd1, msd2)}${l2}${msd2.text}` :
            `${l1}的${d1.name}${msd1.text}与${l2}${msd2.text}无显著差异`;
    }
};
class Mode {
    constructor(name) {
        Mode[Mode.length++] = this;
        $('#mode').innerHTML += `<option>${name}</option>`;
    }
    /**@param {Record<string,{}>} config */
    set config(config) {
        Object.keys(config).forEach(k => this[k] = config[k]);
    }
    init() {
        $('script ~.box', 1).forEach(e => e.remove());
        this.addBox('base');
        Mode.update();
    }
    addBox(boxID) {
        return Mode.addBox(this.boxes.find(e => e.id === boxID));
    }
    static length = 0;
    static output() {
        return $('#output')?.update()
            ?? Mode.addBox({
                id: 'output',
                confirm() {
                    this.$('.data').contentEditable = false;
                    const { innerText, innerHTML } = this.$('.data');
                    navigator.clipboard.write([
                        new ClipboardItem({
                            'text/plain': new Blob([innerText], { type: 'text/plain' }),
                            'text/html': new Blob([innerHTML], { type: 'text/html' }),
                        })
                    ]).then(e => alert('复制成功\n可以带格式粘贴'));
                },
                onEdit() {
                    this.$('.data').contentEditable = true;
                    getSelection()?.selectAllChildren(this.$('.data'));
                },
                update() {
                    function txt(text, tag, style = "") {
                        if (tag == 'p') {
                            style += `font-family: 'Times New Roman','宋体';
                                font-size: 10.5pt;font-weight: normal;color: black;
                                text-indent: 2em;margin: 0;line-height: 1.5;`;
                            text = format(text);
                        }
                        return `<${tag} style="${style}">${text}</${tag}>`
                    }
                    function format(text) {
                        return text.replace(/[χη](2)/g, `η${txt('$1', 'sup')}`)
                            .replace(/ (F|p|t|d|M|SD|r|N|n|B)/g, ` ${txt('$1', 'i')}`)
                            .replace(/(CI)\[/g, `${txt('$1', 'i')}[`)
                            .replace(/\((M) /g, `(${txt('$1', 'i')} `)
                            .replace(/(t|F)检验/g, `${txt('$1', 'i')}检验`)
                            .replace(/(t|Z)分数/g, `${txt('$1', 'i')}分数`)
                            .replace(/ -(\d)/g, ` —$1`);
                    }
                    const arr = $('script ~.box', 1).map(e => e.output()).filter(e => e);
                    arr.unshift(arr.splice(0, 2).join('')); //前两段合并
                    this.$('.data').innerHTML = arr.map(e => txt(e, 'p')).join('');
                },
                initFunc() {
                    this.onEdit();
                    this.editor.$('textarea').style.display = 'none';
                },
            });
    }
    static addBox(config) {
        $('#output')?.remove();
        const box = $('body').appendChild(addElm('div', {
            className: 'box',
            innerHTML: `<div class="data"></div>`,
            onkeypress(e) {
                switch (e.code) {
                    case 'Enter': this.$('.editor input[type=button]').click(); break;
                    default: break;
                }
            },
        }, config)).init();
        // 自动填充
        db.autoFill
            && box.$('input', 1).forEach(e => {
                e.value ||= function () {
                    switch (e.type) {
                        case 'number': switch (e.placeholder) {
                            case 'p': return (Math.random() * 0.1).toFixed(3);
                            case 'df': return (Math.random() * 100).toFixed(0);
                            case 'df1': return (Math.random() * 100).toFixed(0);
                            case 'df2': return (Math.random() * 100).toFixed(0);
                            default: return (Math.random() * 5).toFixed(3);
                        }
                        case 'text': switch (e.placeholder) {
                            case '因变量': return 'Y';
                            case '自变量': return 'X';
                            case '中介变量': return 'M1';
                            case '调节变量': return 'M1';
                            case '水平1': return 'Lv1';
                            case '水平2': return 'Lv2';
                            default: return `${String.fromCharCode(...[0, 0].map(e => parseInt(Math.random() * (90 - 65 + 1) + 65, 10)))}`;
                        };
                        default: return e.type;
                    }
                }();
            });
        return box;
    }
    static update() {
        // 更新变量
        window['vs']?.forEach(e => delete window[e.id]);
        window['vs'] = [];
        const vArr = [{
            type: 'd', elm: $('#dep'), get(e) { return { name: e.value } },
        }, {
            type: 'i', elm: $('#indep'), get(e) { return { name: e.$('input.name').value, ls: e.$('input.level', 1).map(e => e.value), } },
        }, {
            type: 'm', elm: $('#med') || $('#mod'), get(e) { return { name: e.value, } }
        }];
        vArr.filter(v => v.elm ?? console.warn('找不到元素', v))
            .forEach(v => {
                window[`${v.type}s`] = [];
                for (let i = 1; i <= v.elm.childElementCount - 1; i++) {
                    const id = v.type + i,
                        elm = window[id] = $(`#${id}`);
                    Object.assign(elm, v.get(elm));
                    window[`${v.type}s`].push(window[id]);
                    window['vs'].push(window[id]);
                }
            });
        // 清空下级box
        $('#base ~ .box', 1).forEach(e => e.remove());
    }
}
new Mode('差异检验').config = {
    boxes: [
        {
            id: 'base',
            innerHTML: `<div id="dep"><input id="d1" type="text" placeholder="因变量"></div>
                    <div id="indep"><div id="i1">
                        <input class="name" type="text" placeholder="自变量">
                        <input class="level" type="text" placeholder="水平1">
                        <input class="level" type="text" placeholder="水平2">
                    </div></div>`,
            datas() { return $('#indep').$('div', 1); },
            confirm() {
                Mode.update();
                this.Mode.addBox('MSDtable');
                (is.length + i1.ls.length == 3)
                    ? this.Mode.addBox('Ttest')
                    : this.Mode.addBox('Ftest');
            },
            output() {
                let method = '';
                if (is.length == 1 && i1.ls.length == 2 && ds.length == 1) method = `t检验`;
                else {
                    method = `${'0单双'[is.length] || '多'}因素`;
                    if (ds.length > 1) method += '多元';
                    method += '方差分析';
                }
                const indepInfo = is.map(({ name, ls }) => `${ls.length}(${name}: ${ls.join(' / ')})`);
                return `对${ds.map(e => e.name).join('、')}采用${is.length > 1 ? indepInfo.join('×') : ''}${method}。结果发现，`

            },
            initFunc() {
                $('#dep').addTools(1);
                $('#indep').addTools(2); // 当前MSD表格只能存储2个自变量的数据
                $('#i1').addTools(5, 3);
            }
        },
        {
            id: 'MSDtable',
            innerHTML: `<table class="data"></table>`,
            datas() { return [...this.$('.data').$('tbody').children] },
            confirm() {
                const dataArr = this.datas().map(e => {
                    const row = e.$('.MSD', 1).map(ee => ({
                        M: parseFloat(ee.children[0].value),
                        SD: parseFloat(ee.children[1].value),
                    }));
                    return [...row, row.mean()]
                });
                const i2 = window['i2'] || { ls: [] };
                this.MSDdata = {
                    rows: [...i2.ls, 'mean'],
                    columns: [...i1.ls, 'mean'],
                    data: [...dataArr, dataArr[0].map((e, i) => dataArr.map(ee => ee[i]).mean())],
                };
            },
            update() {
                const i2 = window['i2'] || { name: '', ls: [''] };
                const head =
                    `<tr><th>${i2.name}\\${i1.name}</th>`
                    + i1.ls.map((e, i) =>
                        `<th id="i1l${i + 1}">${e}</th>`
                    ).join('') + '</tr>';
                const body = i2.ls.map((e, i) =>
                    `<tr><td id="i2l${i + 1}">${e}</td>`
                    + i1.ls.map((ee, j) =>
                        this.dataElm(`i2l${i + 1}&i1l${j + 1}_MSD`, 'M,SD', {
                            tag: 'td',
                            className: 'MSD',
                        })).join('') + '</tr>'
                ).join('');
                return `<thead>${head}</thead><tbody>${body}</tbody>`;
            },
        },
        {
            id: 'Ttest',
            confirm() { this.datas().map(this.writeData) },
            update() {
                return this.dataElm('i1_Ttest', 't,df,p,CohenD|Cohen d,ciL|CI下限,ciU|CI上限', {
                    elms: [{
                        elm: `<label>${i1.name}</label>`,
                    }]
                });
            },
            output() {
                return (function ({ ls, t, df, p, CohenD, ciL, ciU }) {
                    const [l1, l2] = ls;
                    const msd1 = op.MSD(l1),
                        msd2 = op.MSD(l2);
                    return `${op.sign(p, l1, l2, msd1, msd2)}, t(${df}) = ${t.toFixed(2)}, ${op.ifRes('p***', p)}, Cohen d = ${CohenD.toFixed(2)}, 90%CI[${`${ciL.toFixed(2)}, ${ciU.toFixed(2)}`}]。`
                })(i1);
            }
        },
        {
            id: 'Ftest',
            confirm() {
                $('.post,.simple', 1).forEach(e => e.remove());
                this.is = this.datas().map(e => {
                    const i = this.writeData(e);
                    const { p, id, ls } = i;
                    if (p < 0.05) {
                        if (id.includes('×')) Mode.addBox(this.Mode['simple'](id));
                        else if (ls.length > 2) Mode.addBox(this.Mode['post'](id));
                    }
                    return i;
                });
            },
            update() {
                return is.flatMap((e, i) =>
                    is.choose(i + 1).map(es => {
                        const id = es.map(e => e.id).join('×');
                        if (i > 0) window[id] = {
                            id,
                            name: es.map(e => e.name).join('×'),
                            is: es,
                        };
                        return this.dataElm(`${id}_Ftest`, 'F,df1,df2,p,eta2|η2,ciL|CI下限,ciU|CI上限', {
                            elms: [{
                                elm: `<label>${window[id].name}</label>`,
                            }]
                        });
                    })
                ).join('');
            },
            output() {
                return this.is.map(({ name, F, df1, df2, p, eta2, ciL, ciU }) =>
                    `${op.ifRes('ia', name)}${op.ifRes('p*', p)}, F(${df1}, ${df2}) = ${F.toFixed(2)}, ${op.ifRes('p***', p)}, η2 = ${eta2.toFixed(2)}, 90%CI[${`${ciL.toFixed(2)}, ${ciU.toFixed(2)}`}]`
                ).join('；') + '。';
            },
        },
    ],
    post(id) {
        return {
            i: window[id],
            id: `${id}_post`,
            className: 'post box',
            title: `${id} 事后检验`,
            innerHTML: `<label>${window[id].name}</label>
                    <div class="data"></div>`,
            confirm() { this.writeData(this.$(`.data`), 'post'); },
            update() {
                return this.i.ls.choose(2).map((e, i) =>
                    this.dataElm(`${id}_post${i}`, `l1|水平1|text|value="${e[0]}",l2|水平2|text|value="${e[1]}",p`, { hasDel: 1 })
                ).join('');
            },
            output() {
                return `对${this.i.name}进行事后检验发现，${this.i.post.map(({ l1, l2, p }) =>
                    `${op.sign(p, l1, l2, op.MSD(l1), op.MSD(l2))}, ${op.ifRes('p***', p)}`
                ).join('；')}。`;
            },
            initFunc() { this.$('textarea').style.display = 'none'; },
        }
    },
    simple(id) {
        return {
            i: window[id],
            id: `${id}_simple`,
            className: 'simple box',
            title: `${id} 简单效应分析`,
            innerHTML: `<label>${window[id].name}</label>
                    <div class="data"></div>`,
            confirm() { this.writeData(this.$(`.data`), 'simple'); },
            update() {
                const simples = window[id].is.flatMap((e, i, arr) =>
                    arr[i].ls.flatMap(l =>
                        arr[+!i].ls.choose(2).map(ll => ({
                            fi: e.name,
                            fil: l,
                            l1: ll[0],
                            l2: ll[1]
                        }))
                    )
                );
                return simples.map((e, i) =>
                    this.dataElm(`${id}_simple${i + 1}`,
                        `fi|固定自变量|text|value="${e.fi}",fil|固定水平|text|value="${e.fil}",l1|水平1|text|value="${e.l1}",l2|水平2|text|value="${e.l2}",p`, {
                        elms: [{
                            elm: ' :',
                            index: 1
                        }],
                        hasDel: 1
                    })).join('');
            },
            output() {
                return `${this.i.name}的简单效应分析发现，${this.i.simple.map(({ l1, l2, p, fi, fil }) => {
                    const msd1 = op.MSD(fil, l1),
                        msd2 = op.MSD(fil, l2);
                    return `${fi}为${fil}时，${op.sign(p, l1, l2, msd1, msd2)}, ${op.ifRes('p***', p)}`
                }).join('；')}。`
            },
            initFunc() { this.$('textarea').style.display = 'none'; },
        }
    },
};
new Mode('线性回归').boxes = [
    {
        id: 'base',
        innerHTML: `<div id="dep"><input id="d1" type="text" placeholder="因变量"></div>
                    <div id="indep"><div id="i1">
                        <input class="name" type="text" placeholder="自变量">
                        <input class="level" type="text" placeholder="水平1">
                        <input class="level" type="text" placeholder="水平2">
                    </div></div>`,
        datas() { return $('#indep').$('div', 1); },
        confirm() {
            Mode.update();
            this.Mode.addBox('LinearModel');
            this.Mode.addBox('Linear');
        },
        output() { return `以${ds[0].name}为因变量，${is.map(e => e.name).join('、')}为自变量，进行标准${is.length > 1 ? '多元' : ''}线性回归，各变量的相关矩阵见表1。`; },
        initFunc() {
            $('#dep').addTools(1);
            $('#indep').addTools(9);
            $('#i1').addTools(3);
        }
    },
    {
        id: 'LinearModel',
        confirm() { this.writeData(this.$('#model_data')); },
        update() {
            const html = this.dataElm('model', `R2|调整后R方,F,df1,df2,p`, {
                elms: [{
                    elm: `<label>回归模型</label>`
                }]
            });
            return html.replace(/df2\"/, `$& onchange="$('#Linear .data div[id] input:nth-child(4)',1).forEach(e=>e.value=this.value)"`);
        },
        output() {
            const { R2, F, df1, df2, p } = window['model'];
            return `回归分析结果显示，整体模型调整后的解释方差为${(R2 * 1e2).toFixed(1)}%, F(${df1}, ${df2}) = ${F.toFixed(2)}, ${op.ifRes('p***', p)}。`;
        },
        initFunc() { this.$('textarea').style.display = 'none'; },
    },
    {
        id: 'Linear',
        confirm() { this.is = this.datas().map(this.writeData); },
        update() {
            return is.map(({ id }) => this.dataElm(`${id}_Linear`, 'beta|标准β,t,df,p', {
                elms: [{
                    elm: `<label>${window[id].name}</label>`
                }],
            })).join('');
        },
        output() {
            return '当其他自变量不变时，' + this.is.map(({ name, beta, t, df, p, ls }) =>
                `${name}的${op.ifRes('orient', beta)}向预测作用${op.ifRes('p*', p)}, β = ${beta.toFixed(2)}, t(${df}) = ${t}, ${op.ifRes('p***', p)}`
            ).join('；') + '。';
        },
    }
];
new Mode('中介模型').boxes = [
    {
        id: 'base',
        innerHTML: `<div id="dep"><input id="d1" type="text" placeholder="因变量"></div>
                    <div id="indep"><div id="i1">
                        <input class="name" type="text" placeholder="自变量">
                    </div></div>
                    <div id="med"><input id="m1" class="name" type="text" placeholder="中介变量"></div>`,
        datas() { return $('#indep').$('div', 1); },
        confirm() {
            Mode.update();
            this.Mode.addBox('Mediation');
            this.Mode.addBox('Indirect');
        },
        output() { return `通过回归分析考察在${i1.name}对${d1.name}的影响中${ms.map(e => e.name).join('、')}的中介作用，结果如图1所示。`; },
        initFunc() {
            $('#dep').addTools(1);
            $('#indep').addTools(1);
            $('#med').addTools(3);
        }
    },
    {
        id: 'Mediation',
        datas() { return this.$('.data>div[id]', 1) },
        confirm() { this.datas().forEach(e => this.writeData(e, 'args')); },
        update() {
            this.models = [];
            [i1, ...ms, d1].choose(2).forEach(e => {
                const model = this.models.find(t => t.y == e[1]);
                model?.x.push(e[0]) ?? this.models.push({ y: e[1], x: [e[0]] });
            });
            return this.models.map(({ y, x }) =>
                `<div id="${y.id}_model">${x.map((e, i) =>
                    this.dataElm(e.id, `name|预测变量|text|value="${e.name}",beta|标准β,p`, {
                        hasDel: 1,
                        label: i ? ' ' : y.name,
                    })
                ).join('')}</div>`
            ).join('<hr>');
        },
        output() {
            return '结果表明，' + this.models.map(({ y }) =>
                `以${y.name}为因变量，` + y.args.map(({ name, beta, p }) =>
                    `${name}对${y.name}的${op.ifRes('orient', beta)}向预测作用${op.ifRes('p*', p)}，β = ${beta.toFixed(2)}, ${op.ifRes('p***', p)}`
                ).join('，')
            ).join('；') + '。';
        },
    },
    {
        id: 'Indirect',
        confirm() { this.paths = this.datas().map(this.writeData); },
        update() {
            return ms.map((e, i, arr) =>
                arr.choose(i + 1).map(es => {
                    const path = [i1.name, ...es.map(e => e.name), d1.name].join('→');
                    return this.dataElm(`${path}`, 'effect|间接效应,rate|占比,ciL|CI下限,ciU|CI上限', {
                        hasDel: 1,
                        label: path
                    })
                }).join('')
            ).join('');
        },
        output() {
            return this.paths.map(({ name, effect, rate, ciL, ciU }) =>
                `${name}：间接效应为${effect.toFixed(2)}，占总效应${(rate * 1e2).toFixed(2)}%，95%BootCI[${ciL.toFixed(2)}, ${ciU.toFixed(2)}]`
            ).join('；') + '。';
        },
    }
];
new Mode('调节模型').boxes = [
    {
        id: 'base',
        innerHTML: `<div id="dep"><input id="d1" type="text" placeholder="因变量"></div>
                    <div id="indep"><div id="i1">
                        <input class="name" type="text" placeholder="自变量">
                    </div></div>
                    <div id="mod"><input id="m1" class="name" type="text" placeholder="中介变量"></div>`,
        datas() { return $('#indep').$('div', 1); },
        confirm() {
            Mode.update();
            this.Mode.addBox(confirm('简单斜率分析\n选择Johnson-Neyman法?') ? 'simpleJN' : 'simple');
        },
        output() { return ``; },
        initFunc() {
            $('#dep').addTools(1);
            $('#indep').addTools(1);
            $('#mod').addTools(1);
        }
    },
    {
        id: 'Moderation',
        confirm() { $('#Moderation ~ .box', 1).forEach(e => e.remove()); },
        update() { return this.dataElm('demo', 'a,b,c,d', { label: '示例' }) },
        output() { return `` },
    },
    {
        id: 'simple',
        title: '简单斜率分析',
        confirm() { this.writeData(this.$('.data'), 'simple') },
        update() {
            return `选点法` + ['-1', '0', '+1'].map(e =>
                this.dataElm(`m1_simple1`, 'B,t,df,p', { label: e })
            ).join('').replace(/df\"/, `$& onchange="$('#m1_simple div[id] input[placeholder=df]',1).forEach(e=>e.value=this.value)"`);
        },
        output() {
            return '简单斜率分析发现，'
                + m1.simple.map(({ B, t, df, p }, i) =>
                    `当${m1.name}${'低等高'[i]}于${i == 1 ? '均值' : '1个标准差'}时，${i1.name}对${d1.name}的${op.ifRes('orient', B)}向预测作用${op.ifRes('p*', p)}，`
                    + ` B = ${B.toFixed(2)}, t(${df}) = ${t.toFixed(2)}, ${op.ifRes('p***', p)}`
                ).join('；') + '。';
        },
        initFunc() { this.$('.data').id = 'm1_simple'; },
    },
    {
        id: 'simpleJN',
        title: '简单斜率分析',
        confirm() { this.writeData(this.$('.data'), 'simple') },
        update() {
            return 'Johnson-Neyman' + this.dataElm(`m1_simple1`, 'pLine|正基准线,nLine|负基准线', {
                label: m1.name,
            });
        },
        output() {
            let { pLine, nLine } = m1.simple[0];
            pLine = pLine.toFixed(2);
            nLine = nLine.toFixed(2);
            // pLine < nLine 图像斜率为正
            const desc = `${m1.name}${pLine < nLine ? '低' : '高'}于${pLine}个标准差时，${i1.name}对${d1.name}的负向预测作用显著；`
                + `${m1.name}在${pLine}和${nLine}两个标准差之间时，${i1.name}对${d1.name}的预测作用不显著；`
                + `${m1.name}${pLine < nLine ? '高' : '低'}于${nLine}个标准差时，${i1.name}对${d1.name}的正向预测作用显著`;
            return `采用Johnson-Neyman方法考察${i1.name}对${d1.name}随${m1.name}的变化。结果如图1所示，以${m1.name}的Z分数${pLine}和${nLine}为分界。${desc}。`
        },
        initFunc() {
            this.$('.data').id = 'm1_simple';
            this.editor.$('textarea').style.display = 'none';
        },
    }
];