import xhtml from '@hai2007/tool/xhtml';

// 初始化结点

export function initDom() {

    this._el.innerHTML = "";
    xhtml.setStyles(this._el, {
        "background-color": "#f7f7f7",
        "user-select": "none"
    });

};

// 初始化视图

export function initTableView(itemTable, index, styleToString) {

    let tableTemplate = "";

    // 顶部的
    tableTemplate += "<tr><th class='top-left' open-web-excel></th>";
    for (let k = 0; k < itemTable.content[0].length; k++) {
        tableTemplate += "<th class='top-name' open-web-excel>" + this.$$calcColName(k) + "</th>";
    }
    tableTemplate += '</tr>';

    // 行
    for (let i = 0; i < itemTable.content.length; i++) {

        tableTemplate += "<tr><th class='line-num' open-web-excel>" + (i + 1) + "</th>";

        //  列
        for (let j = 0; j < itemTable.content[i].length; j++) {
            if (itemTable.content[i][j] != 'null') {

                // contenteditable="true" 可编辑状态，则可点击获取焦点，同时内容也是可以编辑的
                // tabindex="0" 点击获取焦点，内容是不可编辑的
                tableTemplate += `<th
                  row='${i + 1}'
                  col='${j + 1}'
                  contenteditable="true"
                  class="item"
                  colspan="${itemTable.content[i][j].colspan}"
                  rowspan="${itemTable.content[i][j].rowspan}"
                  style="${styleToString(itemTable.content[i][j].style)}"
                open-web-excel>${itemTable.content[i][j].value}</th>`;

            }
        }
        tableTemplate += '</tr>';

    }

    this._contentDom[index] = xhtml.append(this._tableFrame, "<table style='display:none;' class='excel-view' open-web-excel>" + tableTemplate + "</table>");

};

let bottomClick = (btnDom, contentDom, index) => {
    for (let i = 0; i < contentDom.length; i++) {
        if (i == index) {
            xhtml.setStyles(contentDom[i], {
                'display': 'table'
            });
            btnDom[i].setAttribute('active', 'yes');
        } else {
            xhtml.setStyles(contentDom[i], {
                'display': 'none'
            });
            btnDom[i].setAttribute('active', 'no');
        }
    }
};

export function initView() {

    this._contentDom = [];
    this._tableFrame = xhtml.append(this._el, "<div></div>");

    xhtml.setStyles(this._tableFrame, {
        "width": "100%",
        "height": "calc(100% - 92px)",
        "overflow": "auto"
    });

    for (let index = 0; index < this._contentArray.length; index++) {

        this.$$initTableView(this._contentArray[index], index, this.$$styleToString);

        xhtml.setStyles(this._contentDom[index], {
            "display": index == 0 ? 'table' : "none"
        });

    }

    this.$$addStyle('excel-view', `

        .excel-view{
            border-collapse: collapse;
            width: 100%;
        }

        .excel-view .top-left{
            width:50px;
            border: 1px solid #d6cccb;
            border-right:none;
            background-color:white;
        }

        .excel-view .top-name{
            border: 1px solid #d6cccb;
            border-bottom:none;
            color:gray;
        }

        .excel-view .line-num{
            width:50px;border: 1px solid #d6cccb;border-right:none;color:gray;
        }

        .excel-view .item{
            vertical-align:top;
            min-width:50px;
            padding:5px;
            white-space: nowrap;
            outline:0.5px solid #555555;
            background-color:white;
        }

    `);

    // 添加底部控制选择显示表格按钮
    let bottomBtns = xhtml.append(this._el, `<div class='bottom-btn' open-web-excel></div>`);

    let addBtn = xhtml.append(bottomBtns, "<span class='add item' open-web-excel>+</span>");

    xhtml.bind(addBtn, 'click', () => {

        // 首先，需要追加数据
        this._contentArray.push(this.$$formatContent()[0]);

        let index = this._contentArray.length - 1;

        // 然后添加table

        this.$$initTableView(this._contentArray[index], index, this.$$styleToString);

        // 添加底部按钮
        let bottomBtn = xhtml.append(bottomBtns, "<span class='name item' open-web-excel>" + this._contentArray[index].name + "</span>");
        this._btnDom.push(bottomBtn);

        xhtml.bind(bottomBtn, 'click', () => {
            bottomClick(this._btnDom, this._contentDom, index);
        });

    });

    this._btnDom = [];

    for (let index = 0; index < this._contentArray.length; index++) {
        let bottomBtn = xhtml.append(bottomBtns, "<span class='name item' open-web-excel>" + this._contentArray[index].name + "</span>");

        // 点击切换显示的视图
        xhtml.bind(bottomBtn, 'click', () => {
            bottomClick(this._btnDom, this._contentDom, index);
        });

        // 双击可以修改名字

        xhtml.bind(bottomBtn, 'dblclick', () => {
            this._btnDom[index].setAttribute('contenteditable', 'true');
        });

        xhtml.bind(bottomBtn, 'blur', () => {
            this._contentArray[index].name = bottomBtn.innerText;
        });

        // 登记起来所有的按钮
        this._btnDom.push(bottomBtn);
    }

    this.$$addStyle('bottom-btn', `

        .bottom-btn{
            width: 100%;
            height: 30px;
            overflow: auto;
            border-top: 1px solid #d6cccb;
            box-sizing: border-box;
        }

        .bottom-btn .item{
            line-height: 30px;
            box-sizing: border-box;
            vertical-align: top;
            display: inline-block;
            cursor: pointer;
        }

        .bottom-btn .add{
            width: 30px;
            text-align: center;
            font-size: 18px;
        }

        .bottom-btn .name{
            font-size: 12px;
            padding: 0 10px;
        }
        .bottom-btn .name:focus{
            outline:none;
        }

        .bottom-btn .name:hover{
            background-color:#efe9e9;
        }

        .bottom-btn .name[active='yes']{
            background-color:white;
        }

    `);

    // 初始化点击第一个
    this._btnDom[0].click();

};
