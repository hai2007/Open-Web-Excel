import xhtml from '@hai2007/tool/xhtml';

export function deleteRow() {

    let rowNodes = xhtml.find(this.__contentDom[this.__tableIndex], () => true, 'tr');

    // 校对行号
    for (let row = this.__rowNum + 1; row <= this.__contentArray[this.__tableIndex].content.length; row++) {
        let colNodes = xhtml.find(rowNodes[row], () => true, 'th');

        // 修改行数
        colNodes[0].innerText = row - 1;

        // 依次修改记录的行数
        for (let col = 1; col < colNodes.length; col++) {
            colNodes[col].setAttribute('row', row - 1);
        }

    }

    let isLastLine = this.__rowNum == this.__contentArray[this.__tableIndex].content.length;// 是否是最后一行

    let downColNodes;
    if (!isLastLine) downColNodes = xhtml.find(rowNodes[this.__rowNum + 1], () => true, 'th');

    // 校对colspan
    for (let col = 1; col <= this.__contentArray[this.__tableIndex].content[0].length; col++) {

        // 如果当前条目隐藏
        if (this.__contentArray[this.__tableIndex].content[this.__rowNum - 1][col - 1].style.display == 'none') {

            // 隐藏的话，就只需要考虑位于左上角的正下方情况
            for (let preRow = this.__rowNum - 1; preRow >= 1; preRow--) {
                if (this.__contentArray[this.__tableIndex].content[preRow - 1][col - 1].style.display != 'none') {

                    // 如果是左上角
                    if (preRow - -this.__contentArray[this.__tableIndex].content[preRow - 1][col - 1].rowspan > this.__rowNum) {

                        let newRowspan = this.__contentArray[this.__tableIndex].content[preRow - 1][col - 1].rowspan - 1;

                        // 结点
                        xhtml.find(rowNodes[preRow], () => true, 'th')[col].setAttribute('rowspan', newRowspan);

                        // 数据
                        this.__contentArray[this.__tableIndex].content[preRow - 1][col - 1].rowspan = newRowspan;

                    }

                    break;
                }

            }

        }

        // 如果没有隐藏，可是是左上角
        // (如果是一行肯定可以直接无视)
        else if (this.__contentArray[this.__tableIndex].content[this.__rowNum - 1][col - 1].rowspan - 1 > 0) {

            let newRowspan = this.__contentArray[this.__tableIndex].content[this.__rowNum - 1][col - 1].rowspan - 1;
            let colspan = this.__contentArray[this.__tableIndex].content[this.__rowNum - 1][col - 1].colspan;

            // 结点
            downColNodes[col].setAttribute('rowspan', newRowspan);
            downColNodes[col].setAttribute('colspan', colspan);
            downColNodes[col].style.display = 'table-cell';

            // 数据
            this.__contentArray[this.__tableIndex].content[this.__rowNum][col - 1].rowspan = newRowspan;
            this.__contentArray[this.__tableIndex].content[this.__rowNum][col - 1].colspan = colspan;
            this.__contentArray[this.__tableIndex].content[this.__rowNum][col - 1].style.display = 'table-cell';

        }

    }

    // 删除当前行
    rowNodes[this.__rowNum].remove();

    // 删除数据
    this.__contentArray[this.__tableIndex].content.splice(this.__rowNum - 1, 1);

    // 重置光标
    this.__btnDom[this.__tableIndex].click();
};

export function deleteCol() {

    let rowNodes = xhtml.find(this.__contentDom[this.__tableIndex], () => true, 'tr');

    // 先删除列标题
    xhtml.find(rowNodes[0], () => true, 'th')[this.__contentArray[this.__tableIndex].content[0].length].remove();

    for (let row = 1; row < rowNodes.length; row++) {
        let colNodes = xhtml.find(rowNodes[row], () => true, 'th');

        // 校对列序号
        for (let col = this.__colNum + 1; col < colNodes.length; col++) {
            colNodes[col].setAttribute('col', col - 1);
        }

        // 删除当前光标所在列
        colNodes[this.__colNum].remove();

        // 数据也要删除
        this.__contentArray[this.__tableIndex].content[row - 1].splice(this.__colNum - 1, 1);
    }

    // 重置光标
    this.__btnDom[this.__tableIndex].click();
};
