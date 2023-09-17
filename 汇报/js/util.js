const sel = () => wps.Selection;
const pre = () => wps.ActivePresentation;

Object.assign(Object.prototype, {
    loop(func) {
        const len = this.Count;
        while (len--) func(this.Item(len + 1), len + 1, this);
    },
});