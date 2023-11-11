export const config = {
    name: 'wps-paper',
    version: 'v2.0',
    comment: {
        err: 'Σ(ﾟдﾟ;)',
        warn: '＞︿＜',
        ok: 'ヾ(≧▽≦*)o',
        info: '(｀・ω・´)',
    },
    update: {
        v: [
            'https://raw.gitmirror.com/Cubxx/wps-paper/main/config.js',
            'https://raw.githubusercontent.com/Cubxx/wps-paper/main/config.js',
        ],
        zip: 'https://api.github.com/repos/Cubxx/wps-paper/zipball',
    },
    ui: {
        write_result: {
            html: '/ui/write_result/index.html',
        },
        help: {
            html: '/ui/help/index.html',
            md_web: 'https://github.com/Cubxx/wps-paper/blob/main/help.md',
            video: 'https://www.bilibili.com/list/525570753?sid=3253331',
        },
    },
};