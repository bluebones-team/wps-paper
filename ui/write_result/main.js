// 加载js-cookie
import { load_jsonp } from "../../lib/web.js";
load_jsonp('../../lib/js.cookie.min.js').then(
    () => {
        const data = JSON.parse(Cookies.get('db') ?? null);
        data ? Object.assign(db, data) : console.warn('cookie已过期');
    },
    () => { console.warn('js-cookie加载失败, 但不影响使用'); }
).then(() => {
    function init() {
        $('html').style['color-scheme'] = db.night ? 'dark' : 'light';
        $('input[title=输出]').style.display = db.autoOutput ? 'none' : '';
        $('#mode').selectedIndex = db.mode;
        $('#mode').dispatchEvent(new Event('change', { bubbles: true }));
    }
    function errInfo(line, message) {
        $('#error').innerHTML += `<li>${line}&nbsp;&nbsp;&nbsp;${message}</li>`;
    };
    // 注册事件
    window.onerror = function (message, file, line) {
        errInfo(line, message);
    };
    window.onunhandledrejection = function (...es) {
        es.forEach(e => {
            const [message, line] = e.reason.stack.split('\n');
            errInfo(line.split(':').at(-2), message);
        });
    };
    window.onbeforeunload = function () {
        Cookies.set('db', JSON.stringify(db), { expires: 7 });
    };
    window.addEventListener('click', e => {
        if (e.target.tagName != 'INPUT') {
            return e.target.closest('#app-tools') && callback();
        }
        switch (e.target.title) {
            case '设置': {
                $('#app-db').style.or('display', '', 'block');
                return callback();
            }
            case '输出': {
                Mode.output();
                return callback();
            }
            case '夜间': {
                db.night = +!$('html').style.or('color-scheme', 'dark', 'light');
                return callback();
            }
        }
        function callback() {
            $('#app-db').innerHTML = Object.keys(db).map(k => `<li>${k}<input name="${k}" type="number" value="${db[k]}"></li>`).join('');
        }
    });
    window.addEventListener('change', e => {
        if (e.target.tagName == 'INPUT' && e.target.closest('#app-db')) {
            $('#app-db>*', 1).forEach(e => db[e.innerText] = +e.$('input').value || 0);
            init();
            return;
        }
        if (e.target == $('#mode')) {
            db.mode = $('#mode').selectedIndex;
            Mode[db.mode].init();
            return;
        }
    });
    // 初始化
    init();
});