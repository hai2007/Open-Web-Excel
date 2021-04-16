# Open Web Excel - ✍️ Web版本的可扩展Excel编辑器

<p>
  <a href="https://hai2007.gitee.io/npm-downloads?interval=7&packages=open-web-excel"><img src="https://img.shields.io/npm/dm/open-web-excel.svg" alt="downloads"></a>
  <a href="https://packagephobia.now.sh/result?p=open-web-excel"><img src="https://packagephobia.now.sh/badge?p=open-web-excel" alt="install size"></a>
  <a href="https://www.jsdelivr.com/package/npm/open-web-excel"><img src="https://data.jsdelivr.com/v1/package/npm/open-web-excel/badge" alt="CDN"></a>
  <a href="https://www.npmjs.com/package/open-web-excel"><img src="https://img.shields.io/npm/v/open-web-excel.svg" alt="Version"></a>
  <a href="https://github.com/hai2007/Open-Web-Excel/blob/master/LICENSE"><img src="https://img.shields.io/npm/l/open-web-excel.svg" alt="License"></a>
  <a href="https://github.com/hai2007/Open-Web-Excel" target='_blank'>
        <img alt="GitHub repo stars" src="https://img.shields.io/github/stars/hai2007/Open-Web-Excel?style=social">
    </a>
</p>

> 温馨提示：使用过程中，你可以查看 [版本历史](./CHANGELOG) 来了解是否需要升级！

> 兼容Chrome、Safari、Edge、Firefox、Opera和IE(9+)等常见浏览器！

## Issues
使用的时候遇到任何问题或有好的建议，请点击进入[issue](https://github.com/hai2007/Open-Web-Excel/issues)，欢迎参与维护！

- 你可以查看[在线用例](https://hai2007.gitee.io/open-web-excel/test/index.html)来快速体验！

## 如何引入

我们推荐你使用npm的方式安装和使用：

```bash
npm install --save open-web-excel
```

当然，你也可以通过CDN的方式引入：

```html
<script src="https://cdn.jsdelivr.net/npm/open-web-excel"></script>
```

## 如何使用

- 特别注意：当前最后一个可用版本（非beta和alpha版本）请查看master分支的说明！

```js
import OpenWebExcel from 'open-web-excel';

var owe = new OpenWebExcel({

    // 编辑器挂载点(必选)
    el: document.getElementById('owe')

});
```

返回的owe里面挂载着后续可控方法：

- 获取当前Excel内容

```js
owe.valueOf();
```

开源协议
---------------------------------------
[MIT](https://github.com/hai2007/Open-Web-Excel/blob/master/LICENSE)

Copyright (c) 2021 [hai2007](https://hai2007.gitee.io/sweethome/) 走一步，再走一步。
