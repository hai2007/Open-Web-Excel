import xhtml from '@hai2007/tool/xhtml';

// 初始化结点

export function initDom() {

    this._el.innerHTML = "";
    xhtml.setStyles(this._el, {
        "background-color": "#f7f7f7"
    });

};

// 初始化视图

export function initView() {

    let tableTemplate = "";

    // 顶部的
    tableTemplate += "<tr><th style='width:50px;border: 1px solid #d6cccb;border-right:none;background-color:white;'></th>";
    for (let k = 0; k < this._contentArray[0].length; k++) {
        tableTemplate += "<th style='border: 1px solid #d6cccb;border-bottom:none;color:gray;'>" + this.$$calcColName(k) + "</th>";
    }
    tableTemplate += '</tr>';

    // 行
    for (let i = 0; i < this._contentArray.length; i++) {

        tableTemplate += "<tr><th style='width:50px;border: 1px solid #d6cccb;border-right:none;color:gray;'>" + (i + 1) + "</th>";

        //  列
        for (let j = 0; j < this._contentArray[i].length; j++) {
            if (this._contentArray[i][j] != 'null') {

                tableTemplate += '<th style="border:1px solid #000000;background-color:white;" index="' + i + '-' + j + '" row-number="' + i + '" col-number="' + j + '" colspan="' + this._contentArray[i][j].colspan + '"  rowspan="' + this._contentArray[i][j].rowspan + '">' + this._contentArray[i][j].content + '</th>';

            }
        }

        tableTemplate += "</tr>";

    }

    this._contentDom = xhtml.append(this._el, "<table>" + tableTemplate + "</table>");

    xhtml.setStyles(this._contentDom, {
        "border-collapse": "collapse",
        "width": "100%"
    });

};
