
// 修改默认输入条目的样式
export function setItemStyle(key, value) {

    // 更新数据内容
    this._contentArray[this._tableIndex].content[this._rowNum - 1][this._colNum - 1].style[key] = value;

    // 更新输入条目
    this._target.style[key] = value;

    // 更新菜单状态
    this.$$updateMenu(this._contentArray[this._tableIndex].content[this._rowNum - 1][this._colNum - 1].style);

};
