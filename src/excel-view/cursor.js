import { isElement } from '@hai2007/tool/type';

// 移动光标到指定位置
export function moveCursorTo(target, rowNum, colNum) {

    // 如果shift被按下，我们认为是在选择区间
    if (this._keyLog.shift) {

        console.log('触发选择区间');

    } else {

        if (isElement(this._target)) this._target.setAttribute('active', 'no');

        // 记录当前鼠标的位置

        this._rowNum = rowNum;
        this._colNum = colNum;
        this._target = target;

        // 先获取对应的原始数据

        let oralItemData = this._contentArray[this._tableIndex].content[rowNum - 1][colNum - 1];

        // 接着更新顶部菜单

        this.$$updateMenu(oralItemData.style);

        target.setAttribute('active', 'yes');
    }

};
