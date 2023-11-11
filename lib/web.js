/**
 * 当前环境根目录
 */
export const root_path = location.href.replace('/index.html', '');
/**
 * 通过jsonp加载网址
 */
export function load_jsonp(url) {
    return new Promise((resolve, reject) => {
        document.head.appendChild(
            Object.assign(document.createElement('script'), {
                src: url,
                onload: resolve,
                onerror: reject,
            })
        );
    });
}
/**
 * 通过fetch加载网址
 * @param {(abort:AbortController.abort)=>void} abort_fn 
 * @returns {Promise<Response>}
 */
export function load_fetch(url, abort_fn) {
    let config = {};
    if (abort_fn) {
        const controller = new AbortController();
        abort_fn(() => controller.abort());
        config = {
            signal: controller.signal,
        };
    }
    return new Promise((resolve, reject) => {
        fetch(url, config).then(async res => {
            if (res.ok) {
                resolve(res);
            } else {
                reject(res);
            }
        }).catch(err => {
            reject(err);
        });
    });
}
/**
 * 依次加载网址，直到网址有效
 * @template T
 * @param {string[]} urls
 * @param {(url:string)=>Promise<T>} load 加载函数
 * @returns {Promise<T>}
 */
export function loads_orderly(urls, load) {
    return new Promise((resolve, reject) => {
        !function send(urls) {
            if (!urls.length) {
                return reject();
            }
            load(urls.shift()).then(
                resolve,
                () => send(urls)
            );
        }([...urls]);
    })
}
/**
 * 打开系统默认浏览器
 */
export function open_url_in_local(url) {
    if (url[0] === '/') {
        url = root_path + url;
    }
    wps.OAAssist.ShellExecute(url);
}
/**
 * 打开wps内置浏览器
 */
export function open_url_in_wps(url, caption, width, height) {
    if (url[0] === '/') {
        url = root_path + url;
    }
    wps.ShowDialog(url, caption, width, height, true);
}