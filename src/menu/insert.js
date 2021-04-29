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

    for (let col = 1; col <= this.__contentArray[this.__tableIndex].content[this.__rowNum == 1 ? 1 : 0].length; col++) {

        // 获取新的数据
        let tempNewItemData = this.$$newItemData();

        /**
         * 嗅探当前单元格情况，
         * 由于会出现合并单元格情况，所以需要对一些特殊情况，进行特殊校对
         */

        let currentItemData = this.__contentArray[this.__tableIndex].content[this.__rowNum][col - 1];

        //  如果不是第一行，而且自己不可见
        if (this.__rowNum != 1 && currentItemData.style.display == 'none') {

            // 那么，我们现在需要确定我们当前行是否位于合并单元格的顶部
            // 因为，如果自己位于顶部，即使不可见，依旧应该可以向上新增一行而不是增高自己

            // 如何直接自己是不是顶部？
            // 我们可以不停的嗅探左边第一个显示的单元格，如果他可以囊括自己，那自己应该就是上顶部
            // 否则就是非第一行

            let isFirstLine = false;
            for (let toLeftCol = col - 1; toLeftCol >= 1; toLeftCol--) {
                let leftItemData = this.__contentArray[this.__tableIndex].content[this.__rowNum][toLeftCol - 1];
                if (leftItemData.style.display != 'none') {

                    // 如果找到的第一个显示的可以包含当前条目
                    if (toLeftCol - -leftItemData.colspan > col) isFirstLine = true;

                    break;
                }

            }

            // 如果是第一行我们就可以直接放过
            if (!isFirstLine) {

                // 到目前为止，我们可以确定的是，当前新增的条目需要隐藏
                tempNewItemData.style.display = 'none';

                // 判断是不是最左边的
                let isLeftFirst = col == 1 || this.__contentArray[this.__tableIndex].content[this.__rowNum][col - 2].style.display != 'none';

                // 如果是最坐标的，就需要负责修改左上角格子的值
                if (isLeftFirst) {

                    for (let preRow = this.__rowNum - 1; preRow > 0; preRow--) {

                        // 接着，让我们寻找这个条目合并后单元格的左上角
                        if (this.__contentArray[this.__tableIndex].content[preRow - 1][col - 1].style.display != 'none') {

                            // 数据
                            this.__contentArray[this.__tableIndex].content[preRow - 1][col - 1].rowspan -= -1;

                            // 结点
                            let leftTopNode = xhtml.find(rowNodes[preRow], () => true, 'th')[col];
                            leftTopNode.setAttribute('rowspan', leftTopNode.getAttribute('rowspan') - -1);

                            // 找到以后别忘了停止
                            break;
                        }
                    }

                }

            }

        }

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

        xhtml.bind(newItemNode, 'input', event => {
            this.$$itemInputHandler(event);
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

        /**
         * 对当前单元格合并情况进行嗅探
         */

        //  如果不是最后一行
        if (this.__rowNum != this.__contentArray[this.__tableIndex].content.length - 1) {

            let currentItemData = this.__contentArray[this.__tableIndex].content[this.__rowNum - 1][col - 1];

            // 不可见或行数不为1
            if (currentItemData.style.display == 'none' || currentItemData.rowspan != '1') {

                // 为了可以之前当前插入点的相对位置，我们首先需要找到合并后单元格左上角的数据和位置
                let leftTopData = this.$$getLeftTop(this.__rowNum, col);

                // 如果不是最底部一行
                if (leftTopData.row - -leftTopData.content.rowspan - 1 > this.__rowNum) {

                    // 到此为止，可以确定当前的条目一定隐藏
                    tempNewItemData.style.display = 'none';

                    // 如果是最左边的
                    if (leftTopData.col == col) {

                        // 数据
                        this.__contentArray[this.__tableIndex].content[leftTopData.row - 1][leftTopData.col - 1].rowspan -= -1;

                        // 结点
                        let leftTopNode = xhtml.find(rowNodes[leftTopData.row], () => true, 'th')[leftTopData.col];
                        leftTopNode.setAttribute('rowspan', leftTopNode.getAttribute('rowspan') - -1);

                    }

                }
            }
        }

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
        xhtml.bind(newItemNode, 'input', event => {
            this.$$itemInputHandler(event);
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

        /**
         * 对当前单元格合并情况进行嗅探
         */

        let currentItemData = this.__contentArray[this.__tableIndex].content[row - 1][this.__colNum - 1];

        //  如果不是第一列，而且自己不可见
        if (this.__colNum != 1 && currentItemData.style.display == 'none') {

            let isFirstCol = false;
            for (let toTopRow = row - 1; toTopRow >= 1; toTopRow--) {
                let topItemData = this.__contentArray[this.__tableIndex].content[toTopRow - 1][this.__colNum];
                if (topItemData.style.display != 'none') {

                    // 如果找到的第一个显示的可以包含当前条目
                    if (toTopRow - -topItemData.rowspan > row) isFirstCol = true;

                    break;
                }

            }

            // 如果是第一列我们就可以直接放过
            if (!isFirstCol) {
                tempNewItemData.style.display = 'none';

                // 判断是不是最顶部的
                let isTopFirst = row == 1 || this.__contentArray[this.__tableIndex].content[row - 2][this.__colNum].style.display != 'none';

                // 如果是最坐标的，就需要负责修改左上角格子的值
                if (isTopFirst) {

                    for (let preCol = this.__colNum - 1; preCol > 0; preCol--) {

                        // 接着，让我们寻找这个条目合并后单元格的左上角
                        if (this.__contentArray[this.__tableIndex].content[row - 1][preCol - 1].style.display != 'none') {

                            // 数据
                            this.__contentArray[this.__tableIndex].content[row - 1][preCol - 1].colspan -= -1;

                            // 结点
                            let leftTopNode = xhtml.find(rowNodes[row], () => true, 'th')[preCol];
                            leftTopNode.setAttribute('colspan', leftTopNode.getAttribute('colspan') - -1);

                            // 找到以后别忘了停止
                            break;
                        }
                    }

                }

            }

        }

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
        xhtml.bind(newItemNode, 'input', event => {
            this.$$itemInputHandler(event);
        });

    }

    // 最后标记右移
    this.__colNum += 1;
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

        /**
        * 对当前单元格合并情况进行嗅探
        */

        //  如果不是最后一列
        if (this.__colNum != this.__contentArray[this.__tableIndex].content[0].length - 1) {

            let currentItemData = this.__contentArray[this.__tableIndex].content[row - 1][this.__colNum - 1];

            // 不可见或列数不为1
            if (currentItemData.style.display == 'none' || currentItemData.colspan != '1') {

                // 为了可以之前当前插入点的相对位置，我们首先需要找到合并后单元格左上角的数据和位置
                let leftTopData = this.$$getLeftTop(row, this.__colNum);

                // 如果不是最右边一列
                if (leftTopData.col - -leftTopData.content.colspan - 1 > this.__colNum) {

                    // 到此为止，可以确定当前的条目一定隐藏
                    tempNewItemData.style.display = 'none';

                    // 如果是最顶部的
                    if (leftTopData.row == row) {

                        // 数据
                        this.__contentArray[this.__tableIndex].content[leftTopData.row - 1][leftTopData.col - 1].colspan -= -1;

                        // 结点
                        let leftTopNode = xhtml.find(rowNodes[leftTopData.row], () => true, 'th')[leftTopData.col];
                        leftTopNode.setAttribute('colspan', leftTopNode.getAttribute('colspan') - -1);

                    }

                }
            }

        }

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
        xhtml.bind(newItemNode, 'input', event => {
            this.$$itemInputHandler(event);
        });
    }

};
