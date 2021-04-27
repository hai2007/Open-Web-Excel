import xhtml from '@hai2007/tool/xhtml';

export function insertUp() {

    let rowNodes = xhtml.find(this.__contentDom[this.__tableIndex], () => true, 'tr');

    // 首先，直接在插入行前面插入一行
    let newRowNode = xhtml.before(rowNodes[this.__rowNum], '<tr><th class="line-num" open-web-excel>' + (this.__rowNum) + '</th></tr>');

    rowNodes.splice(this.__rowNum, 0, newRowNode);
    this.__contentArray[this.__tableIndex].content.splice(this.__rowNum - 1, 0, []);

    // 然后，校对数据
    for (let row = this.__rowNum + 1; row <= rowNodes.length - 1; row++) {
        let colNodes = xhtml.find(rowNodes[row], () => true, 'th');

        // 修改行数
        colNodes[0].innerText = row;

        // 依次修改记录的行数
        for (let col = 1; col < colNodes.length; col++) {
            colNodes[col].setAttribute('row', row);
        }
    }

    for (let col = 1; col <= this.__contentArray[this.__tableIndex].content[0].length; col++) {

        // 获取新的数据
        let tempNewItemData = this.$$newItemData();

        // 追加数据
        this.__contentArray[this.__tableIndex].content[this.__rowNum - 1].push(tempNewItemData);

        // 追加结点
        let newItemNode = xhtml.append(newRowNode,
            `<th row="${this.__rowNum}" col="${col}" contenteditable="true" class="item" colspan="1" rowspan="1" style="${this.$$styleToString(tempNewItemData.style)}" open-web-excel></th>`
        );

        // 绑定事件
        xhtml.bind(newItemNode, 'click', event => {
            this.$$itemClickHandler(event);
        });

    }

    // 最后标记下沉
    this.__rowNum += 1;
};

export function insertDown() {

    let rowNodes = xhtml.find(this.__contentDom[this.__tableIndex], () => true, 'tr');

    // 首先，直接在插入行前面插入一行
    let newRowNode = xhtml.after(rowNodes[this.__rowNum], '<tr><th class="line-num" open-web-excel>' + (this.__rowNum + 1) + '</th></tr>');

    rowNodes.splice(this.__rowNum + 1, 0, newRowNode);
    this.__contentArray[this.__tableIndex].content.splice(this.__rowNum, 0, []);

    // 然后，校对数据
    for (let row = this.__rowNum + 2; row <= rowNodes.length - 1; row++) {
        let colNodes = xhtml.find(rowNodes[row], () => true, 'th');

        // 修改行数
        colNodes[0].innerText = row;

        // 依次修改记录的行数
        for (let col = 1; col < colNodes.length; col++) {
            colNodes[col].setAttribute('row', row);
        }
    }

    for (let col = 1; col <= this.__contentArray[this.__tableIndex].content[0].length; col++) {

        // 获取新的数据
        let tempNewItemData = this.$$newItemData();

        // 追加数据
        this.__contentArray[this.__tableIndex].content[this.__rowNum].push(tempNewItemData);

        // 追加结点
        let newItemNode = xhtml.append(newRowNode,
            `<th row="${this.__rowNum}" col="${col}" contenteditable="true" class="item" colspan="1" rowspan="1" style="${this.$$styleToString(tempNewItemData.style)}" open-web-excel></th>`
        );

        // 绑定事件
        xhtml.bind(newItemNode, 'click', event => {
            this.$$itemClickHandler(event);
        });

    }

};

export function insertLeft() {

    let rowNodes = xhtml.find(this.__contentDom[this.__tableIndex], () => true, 'tr');

    // 先修改顶部的位置提示
    xhtml.append(rowNodes[0], "<th class='top-name' open-web-excel>" + this.$$calcColName(this.__contentArray[this.__tableIndex].content[0].length) + "</th>");

    for (let row = 1; row < rowNodes.length; row++) {
        let colNodes = xhtml.find(rowNodes[row], () => true, 'th');

        // 校对列序号
        for (let col = this.__colNum; col < colNodes.length; col++) {
            colNodes[col].setAttribute('col', col + 1);
        }

        // 获取新的数据
        let tempNewItemData = this.$$newItemData();

        // 追加数据
        this.__contentArray[this.__tableIndex].content[row - 1].splice(this.__colNum - 1, 0, tempNewItemData);

        // 追加结点
        let newItemNode = xhtml.before(colNodes[this.__colNum],
            `<th row="${row}" col="${this.__colNum}" contenteditable="true" class="item" colspan="1" rowspan="1" style="${this.$$styleToString(tempNewItemData.style)}" open-web-excel></th>`
        );

        // 绑定事件
        xhtml.bind(newItemNode, 'click', event => {
            this.$$itemClickHandler(event);
        });

    }

};

export function insertRight() {

    let rowNodes = xhtml.find(this.__contentDom[this.__tableIndex], () => true, 'tr');

    // 先修改顶部的位置提示
    xhtml.append(rowNodes[0], "<th class='top-name' open-web-excel>" + this.$$calcColName(this.__contentArray[this.__tableIndex].content[0].length) + "</th>");

    for (let row = 1; row < rowNodes.length; row++) {
        let colNodes = xhtml.find(rowNodes[row], () => true, 'th');

        // 校对列序号
        for (let col = this.__colNum + 1; col < colNodes.length; col++) {
            colNodes[col].setAttribute('col', col + 1);
        }

        // 获取新的数据
        let tempNewItemData = this.$$newItemData();

        // 追加数据
        this.__contentArray[this.__tableIndex].content[row - 1].splice(this.__colNum, 0, tempNewItemData);

        // 追加结点
        let newItemNode = xhtml.after(colNodes[this.__colNum],
            `<th row="${row}" col="${this.__colNum + 1}" contenteditable="true" class="item" colspan="1" rowspan="1" style="${this.$$styleToString(tempNewItemData.style)}" open-web-excel></th>`
        );

        // 绑定事件
        xhtml.bind(newItemNode, 'click', event => {
            this.$$itemClickHandler(event);
        });
    }

};
