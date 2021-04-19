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

    this._contentDom = [];
    let tableFrame = xhtml.append(this._el, "<div></div>");

    xhtml.setStyles(tableFrame, {
        "width": "100%",
        "height": "calc(100% - 92px)",
        "overflow": "auto"
    });

    for (let index = 0; index < this._contentArray.length; index++) {

        let tableTemplate = "";

        // 顶部的
        tableTemplate += "<tr><th style='width:50px;border: 1px solid #d6cccb;border-right:none;background-color:white;'></th>";
        for (let k = 0; k < this._contentArray[index].content[0].length; k++) {
            tableTemplate += "<th style='border: 1px solid #d6cccb;border-bottom:none;color:gray;'>" + this.$$calcColName(k) + "</th>";
        }
        tableTemplate += '</tr>';

        // 行
        for (let i = 0; i < this._contentArray[index].content.length; i++) {

            tableTemplate += "<tr><th style='width:50px;border: 1px solid #d6cccb;border-right:none;color:gray;'>" + (i + 1) + "</th>";

            //  列
            for (let j = 0; j < this._contentArray[index].content[i].length; j++) {
                if (this._contentArray[index].content[i][j] != 'null') {

                    // contenteditable="true" 可编辑状态，则可点击获取焦点，同时内容也是可以编辑的
                    // tabindex="0" 点击获取焦点，内容是不可编辑的
                    tableTemplate += '<th contenteditable="true" style="vertical-align:top;min-width:50px;padding:5px;white-space: nowrap;outline:0.5px solid #555555;background-color:white;" index="' + i + '-' + j + '" row-number="' + i + '" col-number="' + j + '" colspan="' + this._contentArray[index].content[i][j].colspan + '"  rowspan="' + this._contentArray[index].content[i][j].rowspan + '">' + this._contentArray[index].content[i][j].value + '</th>';

                }
            }
            tableTemplate += '</tr>';

        }

        this._contentDom[index] = xhtml.append(tableFrame, "<table>" + tableTemplate + "</table>");

        xhtml.setStyles(this._contentDom[index], {
            "border-collapse": "collapse",
            "width": "100%",
            "display": index == 0 ? 'table' : "none"
        });

    }

    // 添加底部控制选择显示表格按钮
    let bottomBtns = xhtml.append(this._el, "<div></div>");

    xhtml.setStyles(bottomBtns, {
        "width": "100%",
        "height": "30px",
        "overflow": "auto",
        "border-top": "1px solid #d6cccb",
        'box-sizing': 'border-box'
    });

    let addBtn = xhtml.append(bottomBtns, "<span>+</span>");
    xhtml.setStyles(addBtn, {
        "line-height": "30px",
        "width": "30px",
        "text-align": "center",
        "font-size": "18px",
        'box-sizing': 'border-box',
        'vertical-align': "top",
        'display': 'inline-block',
        'cursor': 'pointer'
    });

    xhtml.bind(addBtn, 'click', () => {

    });

    this._btnDom = [];

    for (let index = 0; index < this._contentArray.length; index++) {
        let bottomBtn = xhtml.append(bottomBtns, "<span>" + this._contentArray[index].name + "</span>");
        xhtml.setStyles(bottomBtn, {
            "line-height": "30px",
            "font-size": "12px",
            'box-sizing': 'border-box',
            "padding": "0 10px",
            'vertical-align': "top",
            'display': 'inline-block',
            'cursor': 'pointer'
        });

        xhtml.bind(bottomBtn, 'click', () => {
            for (let i = 0; i < this._contentDom.length; i++) {
                if (i == index) {
                    xhtml.setStyles(this._contentDom[i], {
                        'display': 'table'
                    });
                    xhtml.setStyles(this._btnDom[i], {
                        "background-color": "white"
                    });
                } else {
                    xhtml.setStyles(this._contentDom[i], {
                        'display': 'none'
                    });
                    xhtml.setStyles(this._btnDom[i], {
                        "background-color": "transparent"
                    });
                }
            }
        });

        this._btnDom.push(bottomBtn);
    }

    // 初始化点击第一个
    this._btnDom[0].click();

};
