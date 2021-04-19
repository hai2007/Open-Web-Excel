import xhtml from '@hai2007/tool/xhtml';

export default function () {

    // 顶部操作栏
    let topDom = xhtml.append(this._el, "<div></div>");

    xhtml.setStyles(topDom, {
        "width": "100%",
        "height": "62px",// 26 + 36
        "overflow": "auto"
    });

    // 菜单
    this._menuDom = xhtml.append(topDom, `<div>
        <span>编辑</span>
        <span>插入</span>
        <span>格式</span>
        <span>公式</span>
        <span>数据</span>
        <span>帮助</span>
    </div>`);

    xhtml.setStyles(this._menuDom, {
        "border-bottom": "1px solid #d6cccb",
        'padding': "0 20px"
    });

    let menuItems = xhtml.find(this._menuDom, () => true, 'span');
    for (let i = 0; i < menuItems.length; i++) {
        xhtml.setStyles(menuItems[i], {
            'display': "inline-block",
            'line-height': '26px',
            'padding': "0 10px",
            'font-size': '12px',
            'cursor': 'pointer',
            'color': "#555555"
        });
    }

};
