import { isElement } from '@hai2007/tool/type';

// 核心方法和工具方法

import { initDom, initView, initTableView } from './excel-view/init';
import { formatContent, calcColName } from './excel-view/tool';
import style from './tool/style';

// 键盘交互总控

import renderKeyboard from './Keyboard';

// 挂载顶部菜单

import menu from './menu/index';

let owe = function (options) {

    if (!(this instanceof owe)) {
        throw new Error('Open-Web-Excel is a constructor and should be called with the `new` keyword');
    }

    // 编辑器挂载点
    if (isElement(options.el)) {

        this._el = options.el;

        // 内容
        this._contentArray = this.$$formatContent(options.content);

    } else {

        // 挂载点是必须的，一定要有
        throw new Error('options.el is not a element!');
    }

    // 先初始化DOM
    this.$$initDom();

    // 挂载菜单
    this.$$createdMenu();

    // 初始化视图
    this.$$initView();

    // 获取当前Excel内容
    this.valueOf = () => {
        return {
            version: "0.1.0",
            filename: "Open-Web-Excel",
            contents: this._contentArray
        };
    };

    // 启动键盘事件监听
    this.$$renderKeyboard();
};

// 挂载辅助方法

owe.prototype.$$formatContent = formatContent;
owe.prototype.$$calcColName = calcColName;
owe.prototype.$$addStyle = style();

// 挂载核心方法

owe.prototype.$$initDom = initDom;
owe.prototype.$$initView = initView;
owe.prototype.$$initTableView = initTableView;

owe.prototype.$$createdMenu = menu;

// 挂载键盘交互总控

owe.prototype.$$renderKeyboard = renderKeyboard;

if (typeof module === "object" && typeof module.exports === "object") {
    module.exports = owe;
} else {
    window.OpenWebExcel = owe;
}
