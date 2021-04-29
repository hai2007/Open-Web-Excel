import { isElement } from '@hai2007/tool/type';

// 核心方法和工具方法

import { initDom, initView, initTableView, itemClickHandler, itemInputHandler } from './excel-view/init';
import { formatContent, calcColName, styleToString, newItemData, getLeftTop } from './excel-view/tool';

import style from './tool/style';

import { moveCursorTo } from './excel-view/cursor';
import { setItemStyle } from './excel-view/modify';
import { calcRegionInfo, showRegion, cancelRegion } from './excel-view/region';

// 键盘交互总控

import renderKeyboard from './Keyboard';

// 挂载顶部菜单

import menu from './menu/index';
import updateMenu from './menu/update';

import { insertUp, insertDown, insertLeft, insertRight } from './menu/insert';
import { deleteRow, deleteCol } from './menu/delete';

let owe = function (options) {

    if (!(this instanceof owe)) {
        throw new Error('Open-Web-Excel is a constructor and should be called with the `new` keyword');
    }

    // 编辑器挂载点
    if (isElement(options.el)) {

        this.__el = options.el;

        // 内容
        this.__contentArray = this.$$formatContent(options.content);

        // 用于选择记录区域
        this.__region = null;

        // 用于记录是否按下了格式刷按钮
        this.__format = false;

    } else {

        // 挂载点是必须的，一定要有
        throw new Error('options.el is not a element!');
    }

    // 启动键盘事件监听
    this.$$renderKeyboard();

    // 先初始化DOM
    this.$$initDom();

    // 挂载菜单
    this.$$createdMenu();

    // 初始化视图
    this.$$initView();

    // 获取当前Excel内容
    this.valueOf = () => {
        return {
            version: "v1",
            filename: "Open-Web-Excel",
            contents: this.__contentArray
        };
    };

};

// 挂载辅助方法

owe.prototype.$$formatContent = formatContent;
owe.prototype.$$calcColName = calcColName;
owe.prototype.$$addStyle = style();
owe.prototype.$$styleToString = styleToString;
owe.prototype.$$newItemData = newItemData;
owe.prototype.$$itemClickHandler = itemClickHandler;
owe.prototype.$$itemInputHandler = itemInputHandler;
owe.prototype.$$getLeftTop = getLeftTop;

// 挂载核心方法

owe.prototype.$$initDom = initDom;
owe.prototype.$$initView = initView;
owe.prototype.$$initTableView = initTableView;

owe.prototype.$$createdMenu = menu;
owe.prototype.$$updateMenu = updateMenu;

owe.prototype.$$moveCursorTo = moveCursorTo;
owe.prototype.$$setItemStyle = setItemStyle;

owe.prototype.$$calcRegionInfo = calcRegionInfo;
owe.prototype.$$showRegion = showRegion;
owe.prototype.$$cancelRegion = cancelRegion;

owe.prototype.$$insertUpNewRow = insertUp;
owe.prototype.$$insertDownNewRow = insertDown;
owe.prototype.$$insertLeftNewCol = insertLeft;
owe.prototype.$$insertRightNewCol = insertRight;

owe.prototype.$$deleteCurrentRow = deleteRow;
owe.prototype.$$deleteCurrentCol = deleteCol;

// 挂载键盘交互总控

owe.prototype.$$renderKeyboard = renderKeyboard;

if (typeof module === "object" && typeof module.exports === "object") {
    module.exports = owe;
} else {
    window.OpenWebExcel = owe;
}
