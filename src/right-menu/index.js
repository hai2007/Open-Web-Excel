import xhtml from '@hai2007/browser/xhtml';
import { getTargetNode } from '../tool/polyfill';

export default function () {

    let rightMenuFrame = xhtml.append(this.__el, `<div class='right-menu-frame' open-web-excel>
        <span class='item' def-type='merge-all' open-web-excel>
            全部合并
        </span>
        <span class='item' def-type='merge-cancel' open-web-excel>
            取消合并
        </span>
        <span class='line' open-web-excel></span>
        <span class='item' def-type='delete-row' open-web-excel>
            删除当前行
        </span>
        <span class='item' def-type='delete-col' open-web-excel>
            删除当前列
        </span>
    </div>`);

    this.__rightMenuDom = rightMenuFrame;

    //  如果点击的是右键菜单，取消全局控制
    xhtml.bind(rightMenuFrame, 'mousedown', event => {
        xhtml.stopPropagation(event);
    });

    // 对菜单添加点击事件
    let menuClickItems = xhtml.find(rightMenuFrame, node => node.getAttribute('def-type'), 'span');

    xhtml.bind(menuClickItems, 'click', event => {

        let node = getTargetNode(event);

        // 获取按钮类型
        let defType = node.getAttribute('def-type');

        this.$$menuHandler(defType, node);

        // 关闭右键菜单
        this.__isrightmenu = false;
        this.__rightMenuDom.style.display = 'none';
    });

    this.$$addStyle('right-menu-frame', `
    .right-menu-frame{
        position:fixed;
        width:120px;
        background-color: white;
        left: 100px;
        top: 100px;
        box-shadow: 0 0 9px 0px #bab2b2;
        font-size: 14px;
        padding:0 5px;
    }
    .right-menu-frame span{
        display: block;
    }
    .right-menu-frame .item{
        padding: 5px 0;
        cursor: pointer;
    }
    .right-menu-frame .item:hover{
        font-weight: 800;
        text-decoration: underline;
    }
    .right-menu-frame .line{
        height: 1px;
        background-color: black;
    }
    `);
};