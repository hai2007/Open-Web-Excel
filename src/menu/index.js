import xhtml from '@hai2007/browser/xhtml';
import colorTemplate from './color-template';
import { getTargetNode } from '../tool/polyfill';

export default function () {

    // 顶部操作栏
    let topDom = xhtml.append(this.__el, `<div class='top-dom' open-web-excel></div>`);

    this.$$addStyle('top-dom', `

       .top-dom{
            width: 100%;
            height: 62px;
            overflow: hidden;
       }

    `);

    // 菜单
    this.__menuDom = xhtml.append(topDom, `<div class='menu' open-web-excel>
        <span open-web-excel>
            操作
            <div open-web-excel>
                <span class='item more' open-web-excel>
                    插入
                    <div open-web-excel>
                        <span class='item' open-web-excel>
                            <span def-type='insert-up'>向上插入</span>
                            <input value='1' open-web-excel />
                            <span def-type='insert-up'>行</span>
                        </span>
                        <span class='item' open-web-excel>
                            <span def-type='insert-down'>向下插入</span>
                            <input value='1' open-web-excel />
                            <span def-type='insert-down'>行</span>
                        </span>
                        <span class='item' open-web-excel>
                            <span def-type='insert-left'>向左插入</span>
                            <input value='1' open-web-excel />
                            <span def-type='insert-left'>列</span>
                        </span>
                        <span class='item' open-web-excel>
                            <span def-type='insert-right'>向右插入</span>
                            <input value='1' open-web-excel />
                            <span def-type='insert-right'>列</span>
                        </span>
                    </div>
                </span>
                <span class='item more' open-web-excel>
                    删除
                    <div open-web-excel>
                        <span class='item' open-web-excel def-type='delete-row'>
                            删除当前行
                        </span>
                        <span class='item' open-web-excel def-type='delete-col'>
                            删除当前列
                        </span>
                    </div>
                </span>
                <span class='item more' open-web-excel>
                    合并单元格
                    <div open-web-excel>
                        <span class='item' def-type='merge-all' open-web-excel>全部合并</span>
                        <span class='item' def-type='merge-cancel' open-web-excel>取消合并</span>
                    </div>
                </span>
            </div>
        </span>
        <span open-web-excel>
            格式
            <div open-web-excel>
                <span class='item' def-type='bold' open-web-excel>粗体</span>
                <span class='item' def-type='italic' open-web-excel>斜体</span>
                <span class='item' def-type='underline' open-web-excel>下划线</span>
                <span class='item' def-type='line-through' open-web-excel>中划线</span>
                <span class='line' open-web-excel></span>
                <span class='item more' open-web-excel>
                    水平对齐
                    <div open-web-excel>
                        <span class='item' def-type='horizontal-left' open-web-excel>左对齐</span>
                        <span class='item' def-type='horizontal-center' open-web-excel>居中对齐</span>
                        <span class='item' def-type='horizontal-right' open-web-excel>右对齐</span>
                    </div>
                </span>
                <span class='item more' open-web-excel>
                    垂直对齐
                    <div open-web-excel>
                        <span class='item' def-type='vertical-top' open-web-excel>顶部对齐</span>
                        <span class='item' def-type='vertical-middle' open-web-excel>居中对齐</span>
                        <span class='item' def-type='vertical-bottom' open-web-excel>底部对齐</span>
                    </div>
                </span>
            </div>
        </span>
        <span open-web-excel>
            帮助
            <div open-web-excel>
                <span class='item' open-web-excel>
                    <a href='https://github.com/hai2007/Open-Web-Excel/issues' open-web-excel target='_blank'>问题反馈</a>
                </span>
            </div>
        </span>
    </div>`);

    this.$$addStyle('menu', `

        .menu{
            border-bottom: 1px solid #d6cccb;
            padding: 0 20px;
            box-sizing: border-box;
            white-space: nowrap;
        }

        .menu>span{
            display: inline-block;
            line-height: 26px;
            padding: 0 10px;
            font-size: 12px;
            cursor: pointer;
            color: #555555;
        }
        .menu>span:hover{
            background: white;
        }

        .menu>span>div{
            margin-left: -10px;
        }

        .menu>span div{
            position:absolute;
            background: white;
            width: 140px;
            box-shadow: 4px 3px 6px 0 #c9c9e2;
            display:none;
            padding:5px 0;
        }

        .menu>span div span{
            display:block;
            position:relative;
            padding:5px 20px;
        }

        .menu>span div span>div{
            left:140px;
            top:0px;
        }

        .menu .line{
            height:1px;
            background-color:#d6cccb;
            padding:0;
            margin:0 10px;
        }

        .menu input{
            width:20px;
            outline:none;
        }

        .menu span:hover>div{
            display:block;
        }

        .menu span.more:after{
            content:">";
            position: absolute;
            right: 12px;
            font-weight: 800;
        }

        .menu a{
            text-decoration: none;
            color: #555555;
        }

        .menu input{
            width:20px;
            outline:none;
        }

        .menu .item.active::before{
            content: "*";
            color: red;
            position: absolute;
            left: 8px;
        }

        .menu .item{
            text-decoration: none;
        }

        .menu .item:hover{
            text-decoration: underline;
        }

    `);

    // 快捷菜单
    this.__menuQuickDom = xhtml.append(topDom, `<div class='quick-menu' open-web-excel>
        <span class='item' def-type='format' open-web-excel>格式刷</span>
        <span class='line' open-web-excel></span>
        <span class='item color' def-type='font-color' open-web-excel>
            文字颜色：<i class='color' open-web-excel></i>
            ${colorTemplate}
        </span>
        <span class='item color' def-type='background-color' open-web-excel>
            填充色：<i class='color' open-web-excel></i>
            ${colorTemplate}
        </span>
        <span class='line' open-web-excel></span>
        <span class='item' def-type='merge-all' open-web-excel>
            全部合并
        </span>
        <span class='item' def-type='merge-cancel' open-web-excel>
            取消合并
        </span>
        <span class='line' open-web-excel></span>
        <span class='item' def-type='horizontal-left' open-web-excel>
            左对齐
        </span>
        <span class='item' def-type='horizontal-center' open-web-excel>
            居中对齐
        </span>
        <span class='item' def-type='horizontal-right' open-web-excel>
            右对齐
        </span>
        <span class='line' open-web-excel></span>
        <span class='item' def-type='vertical-top' open-web-excel>
            顶部对齐
        </span>
        <span class='item' def-type='vertical-middle' open-web-excel>
            居中对齐
        </span>
        <span class='item' def-type='vertical-bottom' open-web-excel>
            底部对齐
        </span>
    </div>`);

    this.$$addStyle('quick-menu', `

        .quick-menu{
            line-height: 36px;
            font-size: 12px;
            white-space: nowrap;
            width: 100%;
            overflow: auto;
        }

        .quick-menu span{
            display:inline-block;
            vertical-align: top;
        }

        .quick-menu span>i.color{
            display: inline-block;
            height: 14px;
            width: 20px;
            border:1px solid #d6cccb;
            vertical-align: middle;
        }

        .quick-menu .item{
            margin:0 10px;
            cursor: pointer;
        }

        .quick-menu .line{
            background-color:#d6cccb;
            width:1px;
            height:22px;
            margin-top:7px;
        }

        .quick-menu .item:hover{
            font-weight: 800;
        }

        .quick-menu .item.active{
            font-weight: 800;
            color: red;
        }

        /* 选择颜色 */

        .color-view{
            font-size: 0px;
            width: 171px;
            position: absolute;
            padding: 10px;
            box-sizing: content-box;
            background: #fefefe;
            box-shadow: 1px 1px 5px #9e9695;
            line-height:1em;
            display:none;
            margin-top: -5px;
            white-space: normal;
        }

        .color:hover>.color-view, .color-view:hover{
            display:block;
        }

        .color-item{
            display: inline-block;
            width: 19px;
            height: 19px;
        }

        .color-item>span{
            width: 15px;
            height: 15px;
            margin: 2px;
            cursor: pointer;
            box-sizing: border-box;
        }

        .color-item>span:hover{
            outline:1px solid black;
        }

    `);

    // 对菜单添加点击事件
    let menuClickItems = xhtml.find(topDom, node => node.getAttribute('def-type'), 'span');

    xhtml.bind(menuClickItems, 'click', event => {

        let node = getTargetNode(event);

        // 获取按钮类型
        let defType = node.getAttribute('def-type');

        this.$$menuHandler(defType, node);

    });

    // 对选择颜色添加点击事件
    let colorItems = xhtml.find(topDom, node => xhtml.hasClass(node, 'color'), 'span');
    for (let i = 0; i < colorItems.length; i++) {

        let colorClickItems = xhtml.find(colorItems[i], () => true, 'span');
        xhtml.bind(colorClickItems, 'click', event => {

            let defType = colorItems[i].getAttribute('def-type');
            let colorValue = getTargetNode(event).style.background;

            // 设置
            this.$$setItemStyle({
                'background-color': 'background',
                'font-color': 'color'
            }[defType], colorValue);

        });

    }

};
