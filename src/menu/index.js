import xhtml from '@hai2007/tool/xhtml';

export default function () {

    // 顶部操作栏
    let topDom = xhtml.append(this._el, `<div class='top-dom' open-web-excel>

    </div>`);

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
            <div open-web-excel>
                <span class='item' open-web-excel>粗体</span>
                <span class='item' open-web-excel>斜体</span>
                <span class='item' open-web-excel>下划线</span>
                <span class='item' open-web-excel>中划线</span>
                <span class='line' open-web-excel></span>
                <span class='item more' open-web-excel>
                    合并单元格
                    <div open-web-excel>
                        <span class='item' open-web-excel>全部合并</span>
                        <span class='item' open-web-excel>水平合并</span>
                        <span class='item' open-web-excel>垂直合并</span>
                        <span class='item' open-web-excel>取消合并</span>
                    </div>
                </span>
                <span class='item more' open-web-excel>
                    水平对齐
                    <div open-web-excel>
                        <span class='item' open-web-excel>左对齐</span>
                        <span class='item' open-web-excel>居中对齐</span>
                        <span class='item' open-web-excel>右对齐</span>
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
        }

        .menu>span div span{
            display:block;
            position:relative;
            padding:10px 20px;
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

    `);

    // 快捷菜单
    this._menuQuickDom = xhtml.append(topDom, `<div class='quick-menu' open-web-excel>
        <span class='item' open-web-excel>格式刷</span>
        <span class='line' open-web-excel></span>
        <span class='item' open-web-excel>
            文字颜色：<i class='color' open-web-excel></i>
        </span>
        <span class='item' open-web-excel>
            填充色：<i class='color' open-web-excel></i>
        </span>
        <span class='line' open-web-excel></span>
        <span class='item' open-web-excel>
            全部合并
        </span>
        <span class='item' open-web-excel>
            水平合并
        </span>
        <span class='item' open-web-excel>
            垂直合并
        </span>
        <span class='item' open-web-excel>
            取消合并
        </span>
        <span class='line' open-web-excel></span>
        <span class='item' open-web-excel>
            左对齐
        </span>
        <span class='item' open-web-excel>
            居中对齐
        </span>
        <span class='item' open-web-excel>
            右对齐
        </span>
    </div>`);

    this.$$addStyle('quick-menu', `

        .quick-menu{
            line-height: 36px;
            font-size: 12px;
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

    `);

};
