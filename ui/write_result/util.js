function addElm(tag, ...config) {
    return Object.assign(document.createElement(tag), ...config);
}
function _$(selectors, isAll) {
    return isAll ? [...this.querySelectorAll(selectors)] : this.querySelector(selectors);
}
const $ = _$.bind(document);
const db = {
    autoFill: 0,
    autoOutput: 1,
    mode: 0,
    night: 0,
};
