import xhtml from "@hai2007/tool/xhtml";

export default function (style) {

    // 更新顶部菜单

    let menuItems = xhtml.find(this.__menuDom, node => node.getAttribute('def-type'), 'span');
    for (let i = 0; i < menuItems.length; i++) {

        // 获取按钮类型
        let defType = menuItems[i].getAttribute('def-type');

        // 粗体
        if (defType == 'bold') {

            if (style['font-weight'] == 'bold') {
                xhtml.addClass(menuItems[i], 'active');
            } else {
                xhtml.removeClass(menuItems[i], 'active');
            }

        }

        // 粗体
        else if (defType == 'italic') {

            if (style['font-style'] == 'italic') {
                xhtml.addClass(menuItems[i], 'active');
            } else {
                xhtml.removeClass(menuItems[i], 'active');
            }

        }

        // 中划线
        else if (defType == 'underline') {

            if (style['text-decoration'] == 'underline') {
                xhtml.addClass(menuItems[i], 'active');
            } else {
                xhtml.removeClass(menuItems[i], 'active');
            }

        }

        // 下划线
        else if (defType == 'line-through') {

            if (style['text-decoration'] == 'line-through') {
                xhtml.addClass(menuItems[i], 'active');
            } else {
                xhtml.removeClass(menuItems[i], 'active');
            }

        }

        // 水平对齐方式
        else if (/^horizontal\-/.test(defType)) {

            if (defType == 'horizontal-' + style['text-align']) {
                xhtml.addClass(menuItems[i], 'active');
            } else {
                xhtml.removeClass(menuItems[i], 'active');
            }

        }

        // 垂直对齐方式
        else if (/^vertical\-/.test(defType)) {

            if (defType == 'vertical-' + style['vertical-align']) {
                xhtml.addClass(menuItems[i], 'active');
            } else {
                xhtml.removeClass(menuItems[i], 'active');
            }

        }

    }

    // 更新快速使用菜单

    let quickItems = xhtml.find(this.__menuQuickDom, node => node.getAttribute('def-type'), 'span');
    for (let i = 0; i < quickItems.length; i++) {

        // 获取按钮类型
        let defType = quickItems[i].getAttribute('def-type');

        // 文字颜色
        if (defType == 'font-color') {
            quickItems[i].getElementsByTagName('i')[0].style.backgroundColor = style.color;
        }

        // 填充色
        else if (defType == 'background-color') {
            quickItems[i].getElementsByTagName('i')[0].style.backgroundColor = style.background;
        }

        // 水平对齐方式
        else if (/^horizontal\-/.test(defType)) {

            if (defType == 'horizontal-' + style['text-align']) {
                xhtml.addClass(quickItems[i], 'active');
            } else {
                xhtml.removeClass(quickItems[i], 'active');
            }

        }

        // 垂直对齐方式
        else if (/^vertical\-/.test(defType)) {

            if (defType == 'vertical-' + style['vertical-align']) {
                xhtml.addClass(quickItems[i], 'active');
            } else {
                xhtml.removeClass(quickItems[i], 'active');
            }

        }

    }

};
