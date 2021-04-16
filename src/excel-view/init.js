import xhtml from '@hai2007/tool/xhtml';

// 初始化结点

export function initDom() {

    this._el.innerHTML = "";

};

// 初始化视图

export function initView() {

    let tableTemplate = "";

    // 行
    for (let i = 0; i < this._contentArray.length; i++) {

        tableTemplate += "<tr>";

        //  列
        for (let j = 0; j < this._contentArray[i].length; j++) {
            if (this._contentArray[i][j] != 'null') {

                tableTemplate += '<th style="border:1px solid #000000;" index="' + i + '-' + j + '" row-number="' + i + '" col-number="' + j + '" colspan="' + this._contentArray[i][j].colspan + '"  rowspan="' + this._contentArray[i][j].rowspan + '">' + this._contentArray[i][j].content + '</th>';

            }
        }

        tableTemplate += "</tr>";

    }

    this._contentDom = xhtml.append(this._el, "<table style='border-collapse: collapse;width: 100%;'>" + tableTemplate + "</table>");

};
