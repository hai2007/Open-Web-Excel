// 移动光标到指定位置
export function moveCursorTo(rowNum, colNum) {

    // 记录当前鼠标的位置

    this._rowNum = rowNum;
    this._colNum = colNum;

    // 先获取对应的原始数据

    let oralItemData = this._contentArray[this._tableIndex].content[rowNum - 1][colNum - 1];

    // 接着更新顶部菜单

    this.$$updateMenu(oralItemData.style);

};
