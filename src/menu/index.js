import xhtml from '@hai2007/tool/xhtml';

export default function () {

    // 顶部操作栏
    let topDom = xhtml.append(this._el, "<div class='top-dom' open-web-excel></div>");

    this.$$addStyle('top-dom', `

       .top-dom{
            width: 100%;
            height: 62px;
            overflow: auto;
       }

    `);

    // 菜单
    this._menuDom = xhtml.append(topDom, `<div class='menu' open-web-excel>
        <span open-web-excel>
            格式
        </span>
        <span open-web-excel>帮助</span>
    </div>`);

    this.$$addStyle('menu', `

        .menu{
            border-bottom: 1px solid #d6cccb;
            padding: 0 20px;
            box-sizing: border-box;
        }

        .menu>span{
            display: inline-block;
            line-height: 26px;
            padding: 0 10px;
            font-size: 12px;
            cursor: pointer;
            color: #555555;
        }

    `);

};
