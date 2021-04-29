import xhtml from '@hai2007/tool/xhtml';
import { getTargetNode } from '../tool/polyfill';

// 初始化结点

export function initDom() {

    this.__el.innerHTML = "";
    xhtml.setStyles(this.__el, {
        "background-color": "#f7f7f7",
        "user-select": "none"
    });

};

export function itemInputHandler(event) {
    this.__contentArray[this.__tableIndex].content[+getTargetNode(event).getAttribute('row') - 1][+getTargetNode(event).getAttribute('col') - 1].value = getTargetNode(event).innerText;
};

export function itemClickHandler(event) {
    // 如果格式刷按下了
    if (this.__format == true) {

        let rowNodes = xhtml.find(this.__contentDom[this.__tableIndex], () => true, 'tr');

        let targetStyle = this.__contentArray[this.__tableIndex].content[+getTargetNode(event).getAttribute('row') - 1][+getTargetNode(event).getAttribute('col') - 1].style;

        for (let row = this.__region.info.row[0]; row <= this.__region.info.row[1]; row++) {

            let colNodes = xhtml.find(rowNodes[row], () => true, 'th');

            for (let col = this.__region.info.col[0]; col <= this.__region.info.col[1]; col++) {

                // 遍历所有的样式
                for (let key in targetStyle) {

                    // 修改界面显示
                    colNodes[col].style[key] = targetStyle[key];

                    // 修改数据
                    this.__contentArray[this.__tableIndex].content[row - 1][col - 1].style[key] = targetStyle[key];

                }

            }

        }

        // 取消标记格式刷
        this.__format = false;
        xhtml.removeClass(xhtml.find(this.__menuQuickDom, node => node.getAttribute('def-type') == 'format', 'span')[0], 'active');

    }

    this.$$moveCursorTo(getTargetNode(event), +getTargetNode(event).getAttribute('row'), +getTargetNode(event).getAttribute('col'));
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
        tableTemplate += '</tr>';

    }

    this.__contentDom[index] = xhtml.append(this.__tableFrame, "<table style='display:none;' class='excel-view' open-web-excel>" + tableTemplate + "</table>");

    // 后续动态新增的需要重新绑定

    let items = xhtml.find(this.__contentDom[index], node => xhtml.hasClass(node, 'item'), 'th');

    xhtml.bind(items, 'click', event => {

        this.$$itemClickHandler(event);

    });

    xhtml.bind(items, 'input', event => {

        this.$$itemInputHandler(event);

    });

};

let bottomClick = (target, index) => {
    for (let i = 0; i < target.__contentDom.length; i++) {
        if (i == index) {
            xhtml.setStyles(target.__contentDom[i], {
                'display': 'table'
            });
            target.__btnDom[i].setAttribute('active', 'yes');
        } else {
            xhtml.setStyles(target.__contentDom[i], {
                'display': 'none'
            });
            target.__btnDom[i].setAttribute('active', 'no');
        }
    }
    target.__tableIndex = index;

    target.$$moveCursorTo(target.__contentDom[index].getElementsByTagName('tr')[1].getElementsByTagName('th')[1], 1, 1);
};

export function initView() {

    this.__contentDom = [];
    this.__tableFrame = xhtml.append(this.__el, "<div></div>");

    xhtml.setStyles(this.__tableFrame, {
        "width": "100%",
        "height": "calc(100% - 92px)",
        "overflow": "auto"
    });

    for (let index = 0; index < this.__contentArray.length; index++) {

        this.$$initTableView(this.__contentArray[index], index, this.$$styleToString);

        xhtml.setStyles(this.__contentDom[index], {
            "display": index == 0 ? 'table' : "none"
        });

    }

    this.$$addStyle('excel-view', `

        .excel-view{
            border-collapse: collapse;
            width: 100%;
        }

        .excel-view .top-left{
            border: 1px solid #d6cccb;
            border-right:none;
            background-color:white;
        }

        .excel-view .top-name{
            border: 1px solid #d6cccb;
            border-bottom:none;
            color:gray;
            font-size:12px;
        }

        .excel-view .line-num{
            padding:0 5px;
            border: 1px solid #d6cccb;
            border-right:none;
            color:gray;
            font-size:12px;
        }

        .excel-view .item{
            min-width:50px;
            white-space: nowrap;
            border:0.5px solid rgba(85,85,85,0.5);
            outline:none;
            font-size:12px;
            padding:2px;
        }

        .excel-view .item[active='yes']{
            outline: 2px dashed red;
        }

    `);

    // 添加底部控制选择显示表格按钮
    let bottomBtns = xhtml.append(this.__el, `<div class='bottom-btn' open-web-excel></div>`);

    let addBtn = xhtml.append(bottomBtns, "<span class='add item' open-web-excel>+</span>");

    xhtml.bind(addBtn, 'click', () => {

        // 首先，需要追加数据
        this.__contentArray.push(this.$$formatContent()[0]);

        let index = this.__contentArray.length - 1;

        // 然后添加table

        this.$$initTableView(this.__contentArray[index], index, this.$$styleToString);

        // 添加底部按钮
        let bottomBtn = xhtml.append(bottomBtns, "<span class='name item' open-web-excel>" + this.__contentArray[index].name + "</span>");
        this.__btnDom.push(bottomBtn);

        xhtml.bind(bottomBtn, 'click', () => {
            bottomClick(this, index);
        });

    });

    this.__btnDom = [];

    for (let index = 0; index < this.__contentArray.length; index++) {
        let bottomBtn = xhtml.append(bottomBtns, "<span class='name item' open-web-excel>" + this.__contentArray[index].name + "</span>");

        // 点击切换显示的视图
        xhtml.bind(bottomBtn, 'click', () => {
            bottomClick(this, index);
        });

        // 双击可以修改名字

        xhtml.bind(bottomBtn, 'dblclick', () => {
            this.__btnDom[index].setAttribute('contenteditable', 'true');
        });

        xhtml.bind(bottomBtn, 'blur', () => {
            this.__contentArray[index].name = bottomBtn.innerText;
        });

        // 登记起来所有的按钮
        this.__btnDom.push(bottomBtn);
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
    this.__btnDom[0].click();

};
