import xhtml from '@hai2007/tool/xhtml';

export default function () {

    // 顶部菜单
    this._menuDom = xhtml.append(this._el, "<div></div>");

    xhtml.setStyles(this._menuDom, {
        "width": "100%",
        "height": "100px",
        "overflow": "auto"
    });



};
