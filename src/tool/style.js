
let addUniqueNamespace = function (style) {

    let uniqueNameSpace = 'open-web-excel';

    style = style.replace(/( {0,}){/g, "{");
    style = style.replace(/( {0,}),/g, ",");

    let temp = "";
    // 分别表示：是否处于注释中、是否处于内容中、是否由于特殊情况在遇到{前完成了hash
    let isSpecial = false, isContent = false, hadComplete = false;
    for (let i = 0; i < style.length; i++) {
        if (style[i] == ':' && !isSpecial && !hadComplete && !isContent) {
            hadComplete = true;
            temp += "[" + uniqueNameSpace + "]";
        } else if (style[i] == '{' && !isSpecial) {
            isContent = true;
            if (!hadComplete) temp += "[" + uniqueNameSpace + "]";
        } else if (style[i] == '}' && !isSpecial) {
            isContent = false;
            hadComplete = false;
        } else if (style[i] == '/' && style[i + 1] == '*') {
            isSpecial = true;
        } else if (style[i] == '*' && style[i + 1] == '/') {
            isSpecial = false;
        } else if (style[i] == ',' && !isSpecial && !isContent) {
            if (!hadComplete) temp += "[" + uniqueNameSpace + "]";
            hadComplete = false;
        }

        temp += style[i];

    }

    return temp;

};

export default function () {

    if ('open-web-excel@style' in window) {
        // todo
    } else {
        window['open-web-excel@style'] = {};
    }

    let head = document.head || document.getElementsByTagName('head')[0];

    return function (keyName, styleString) {
        if (window['open-web-excel@style'][keyName]) {
            // todo
        } else {
            window['open-web-excel@style'][keyName] = true;

            // 创建style标签
            let styleElement = document.createElement('style');
            styleElement.setAttribute('type', 'text/css');

            // 写入样式内容
            // 添加统一的后缀是防止污染
            styleElement.innerHTML = addUniqueNamespace(styleString);

            // 添加到页面
            head.appendChild(styleElement);
        }
    };
};
