import { isNumber } from "@hai2007/tool/type";

export function styleToString(style) {

    let styleString = "";
    for (let key in style) {
        styleString += key + ":" + style[key] + ';';
    }

    return styleString;
};

export function newItemData() {
    return {
        value: " ", colspan: "1", rowspan: "1",
        style: {
            display: "table-cell",
            color: 'black',
            background: 'white',
            'vertical-align': 'top',
            'text-align': 'left',
            'font-weight': "normal",// bold粗体
            'font-style': 'normal',// italic斜体
            'text-decoration': 'none'// line-through中划线 underline下划线
        }
    };
};

export function formatContent(file) {

    // 如果传递了内容
    if (file && 'version' in file && file.filename == 'Open-Web-Excel') {

        // 后续如果格式进行了升级，可以格式兼容转换成最新版本
        return file.contents;
    }

    // 否则，自动初始化
    else {

        let content = [];
        for (let i = 0; i < 100; i++) {
            let rowArray = []
            for (let j = 0; j < 30; j++) {
                rowArray.push(this.$$newItemData());
            }

            content.push(rowArray);
        }
        return [{
            name: "未命名",
            content
        }];
    }

};

export function calcColName(index) {
    if (!isNumber(index)) return index;

    let codes = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z'];
    let result = "";

    while (true) {

        // 求解当前坐标
        let _index = index % 26;

        // 拼接
        result = codes[_index] + result;

        // 求解余下的数
        index = Math.floor(index / 26);

        if (index == 0) break;

        index -= 1;
    }
    return result;
};
