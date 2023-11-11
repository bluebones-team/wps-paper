import { config } from '../../config.js'
import { load_jsonp, root_path, open_url_in_local } from '../../lib/web.js'
let mdFile_path = './help.md';
const { help: { video } } = config.ui;
const logger = {
    content: document.querySelector('#md'),
    fill(text) {
        this.content.innerHTML = text;
    },
    log(text, link) {
        this.content.innerHTML += '<br>' + link
            ? `<a href="${link}">${text}</a>`
            : text;
    },
    err(msg, line = '') {
        this.content.innerHTML += '<br>' + `<i style='color:red;'>${line}&emsp;${msg.replaceAll('\n', '<br>')}</i>`;
    },
};
window.addEventListener('error', ({ message, lineno }) => logger.err(message, lineno));
window.addEventListener('unhandledrejection', ({ reason }) => logger.err(reason.stack));
document.querySelector('a').onclick = () => open_url_in_local(video);
if (location.protocol !== 'file:') { //在线安装
    const res = await fetch(mdFile_path);
    const text = await res.text();
    parseMD(text);
} else { //本地安装
    mdFile_path = root_path.replace('file:///', '') + mdFile_path.slice(1);
    const text = wps.FileSystem.ReadFile(mdFile_path);
    parseMD(text);
}
function parseMD(text) {
    Promise.all([
        load_jsonp('../../lib/marked.min.js'),
    ]).then(() => {
        logger.fill(marked.parse(text));
    }, () => {
        logger.log('资源请求错误', mdFile_path);
    });
}