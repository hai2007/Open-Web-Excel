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
