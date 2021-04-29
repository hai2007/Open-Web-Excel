/*!
* Open Web Excel - ✍️ An Excel Used on the Browser Side.
* git+https://github.com/hai2007/Open-Web-Excel.git
*
* author 你好2007
*
* version 0.2.1
*
* Copyright (c) 2021 hai2007 走一步，再走一步。
* Released under the MIT license
*
* Date:Thu Apr 29 2021 10:36:11 GMT+0800 (GMT+08:00)
*/

"use strict";

function _typeof(obj) { "@babel/helpers - typeof"; if (typeof Symbol === "function" && typeof Symbol.iterator === "symbol") { _typeof = function _typeof(obj) { return typeof obj; }; } else { _typeof = function _typeof(obj) { return obj && typeof Symbol === "function" && obj.constructor === Symbol && obj !== Symbol.prototype ? "symbol" : typeof obj; }; } return _typeof(obj); }

(function () {
  'use strict';

  var toString = Object.prototype.toString;
  /**
   * 获取一个值的类型字符串[object type]
   *
   * @param {*} value 需要返回类型的值
   * @returns {string} 返回类型字符串
   */

  function getType(value) {
    if (value == null) {
      return value === undefined ? '[object Undefined]' : '[object Null]';
    }

    return toString.call(value);
  }
  /**
   * 判断一个值是不是number。
   *
   * @param {*} value 需要判断类型的值
   * @returns {boolean} 如果是number返回true，否则返回false
   */


  function _isNumber(value) {
    return typeof value === 'number' || value !== null && _typeof(value) === 'object' && getType(value) === '[object Number]';
  }
  /**
   * 判断一个值是不是一个朴素的'对象'
   * 所谓"纯粹的对象"，就是该对象是通过"{}"或"new Object"创建的
   *
   * @param {*} value 需要判断类型的值
   * @returns {boolean} 如果是朴素的'对象'返回true，否则返回false
   */


  function _isPlainObject(value) {
    if (value === null || _typeof(value) !== 'object' || getType(value) != '[object Object]') {
      return false;
    } // 如果原型为null


    if (Object.getPrototypeOf(value) === null) {
      return true;
    }

    var proto = value;

    while (Object.getPrototypeOf(proto) !== null) {
      proto = Object.getPrototypeOf(proto);
    }

    return Object.getPrototypeOf(value) === proto;
  }

  var domTypeHelp = function domTypeHelp(types, value) {
    return value !== null && _typeof(value) === 'object' && types.indexOf(value.nodeType) > -1 && !_isPlainObject(value);
  };

  var isNumber = _isNumber; // 结点类型

  var isElement = function isElement(input) {
    return domTypeHelp([1, 9, 11], input);
  };
  /*!
   * 💡 - 提供常用的DOM操作方法
   * https://github.com/hai2007/tool.js/blob/master/xhtml.js
   *
   * author hai2007 < https://hai2007.gitee.io/sweethome >
   *
   * Copyright (c) 2021-present hai2007 走一步，再走一步。
   * Released under the MIT license
   */
  // 命名空间路径


  var namespace = {
    svg: "http://www.w3.org/2000/svg",
    xhtml: "http://www.w3.org/1999/xhtml",
    xlink: "http://www.w3.org/1999/xlink",
    xml: "http://www.w3.org/XML/1998/namespace",
    xmlns: "http://www.w3.org/2000/xmlns/"
  };
  /**
   * 结点操作补充
   */

  var xhtml = {
    // 阻止冒泡
    "stopPropagation": function stopPropagation(event) {
      event = event || window.event;

      if (event.stopPropagation) {
        //这是其他非IE浏览器
        event.stopPropagation();
      } else {
        event.cancelBubble = true;
      }
    },
    // 阻止默认事件
    "preventDefault": function preventDefault(event) {
      event = event || window.event;

      if (event.preventDefault) {
        event.preventDefault();
      } else {
        event.returnValue = false;
      }
    },
    // 判断是否是结点
    "isNode": function isNode(param) {
      return param && (param.nodeType === 1 || param.nodeType === 9 || param.nodeType === 11);
    },
    // 绑定事件
    "bind": function bind(dom, eventType, callback) {
      if (dom.constructor === Array || dom.constructor === NodeList || dom.constructor === HTMLCollection) {
        for (var i = 0; i < dom.length; i++) {
          this.bind(dom[i], eventType, callback);
        }

        return;
      }

      if (window.attachEvent) dom.attachEvent("on" + eventType, callback);else dom.addEventListener(eventType, callback, false);
    },
    // 去掉绑定事件
    "unbind": function unbind(dom, eventType, handler) {
      if (dom.constructor === Array || dom.constructor === NodeList || dom.constructor === HTMLCollection) {
        for (var i = 0; i < dom.length; i++) {
          this.unbind(dom[i], eventType, handler);
        }

        return;
      }

      if (window.detachEvent) dom.detachEvent('on' + eventType, handler);else dom.removeEventListener(eventType, handler, false);
    },
    // 在当前上下文context上查找结点
    // selectFun可选，返回boolean用以判断当前面对的结点是否保留
    "find": function find(context, selectFun, tagName) {
      if (!this.isNode(context)) return [];
      var nodes = context.getElementsByTagName(tagName || '*');
      var result = [];

      for (var i = 0; i < nodes.length; i++) {
        if (this.isNode(nodes[i]) && (typeof selectFun != "function" || selectFun(nodes[i]))) result.push(nodes[i]);
      }

      return result;
    },
    // 寻找当前结点的孩子结点
    // selectFun可选，返回boolean用以判断当前面对的结点是否保留
    "children": function children(dom, selectFun) {
      var nodes = dom.childNodes;
      var result = [];

      for (var i = 0; i < nodes.length; i++) {
        if (this.isNode(nodes[i]) && (typeof selectFun != "function" || selectFun(nodes[i]))) result.push(nodes[i]);
      }

      return result;
    },
    // 判断结点是否有指定class
    // clazzs可以是字符串或数组字符串
    // notStrict可选，boolean值，默认false表示结点必须包含全部class,true表示至少包含一个即可
    "hasClass": function hasClass(dom, clazzs, notStrict) {
      if (clazzs.constructor !== Array) clazzs = [clazzs];
      var class_str = " " + (dom.getAttribute('class') || "") + " ";

      for (var i = 0; i < clazzs.length; i++) {
        if (class_str.indexOf(" " + clazzs[i] + " ") >= 0) {
          if (notStrict) return true;
        } else {
          if (!notStrict) return false;
        }
      }

      return true;
    },
    // 删除指定class
    "removeClass": function removeClass(dom, clazz) {
      var oldClazz = " " + (dom.getAttribute('class') || "") + " ";
      var newClazz = oldClazz.replace(" " + clazz.trim() + " ", " ");
      dom.setAttribute('class', newClazz.trim());
    },
    // 添加指定class
    "addClass": function addClass(dom, clazz) {
      if (this.hasClass(dom, clazz)) return;
      var oldClazz = dom.getAttribute('class') || "";
      dom.setAttribute('class', oldClazz + " " + clazz);
    },
    // 字符串变成结点
    // isSvg可选，boolean值，默认false表示结点是html，为true表示svg类型
    "toNode": function toNode(template, isSvg) {
      var frame; // html和svg上下文不一样

      if (isSvg) frame = document.createElementNS(namespace.svg, 'svg');else {
        var frameTagName = 'div'; // 大部分的标签可以直接使用div作为容器
        // 部分特殊的需要特殊的容器标签

        if (/^<tr[> ]/.test(template)) {
          frameTagName = "tbody";
        } else if (/^<th[> ]/.test(template) || /^<td[> ]/.test(template)) {
          frameTagName = "tr";
        } else if (/^<thead[> ]/.test(template) || /^<tbody[> ]/.test(template)) {
          frameTagName = "table";
        }

        frame = document.createElement(frameTagName);
      } // 低版本浏览器svg没有innerHTML，考虑是vue框架中，没有补充

      frame.innerHTML = template;
      var childNodes = frame.childNodes;

      for (var i = 0; i < childNodes.length; i++) {
        if (this.isNode(childNodes[i])) return childNodes[i];
      }
    },
    // 主动触发事件
    "trigger": function trigger(dom, eventType) {
      //创建event的对象实例。
      if (document.createEventObject) {
        // IE浏览器支持fireEvent方法
        dom.fireEvent('on' + eventType, document.createEventObject());
      } // 其他标准浏览器使用dispatchEvent方法
      else {
          var _event = document.createEvent('HTMLEvents'); // 3个参数：事件类型，是否冒泡，是否阻止浏览器的默认行为


          _event.initEvent(eventType, true, false);

          dom.dispatchEvent(_event);
        }
    },
    // 获取样式
    "getStyle": function getStyle(dom, name) {
      // 获取结点的全部样式
      var allStyle = document.defaultView && document.defaultView.getComputedStyle ? document.defaultView.getComputedStyle(dom, null) : dom.currentStyle; // 如果没有指定属性名称，返回全部样式

      return typeof name === 'string' ? allStyle.getPropertyValue(name) : allStyle;
    },
    // 获取元素位置
    "offsetPosition": function offsetPosition(dom) {
      var left = 0;
      var top = 0;
      top = dom.offsetTop;
      left = dom.offsetLeft;
      dom = dom.offsetParent;

      while (dom) {
        top += dom.offsetTop;
        left += dom.offsetLeft;
        dom = dom.offsetParent;
      }

      return {
        "left": left,
        "top": top
      };
    },
    // 获取鼠标相对元素位置
    "mousePosition": function mousePosition(dom, event) {
      var bounding = dom.getBoundingClientRect();
      if (!event || !event.clientX) throw new Error('Event is necessary!');
      return {
        "x": event.clientX - bounding.left,
        "y": event.clientY - bounding.top
      };
    },
    // 删除结点
    "remove": function remove(dom) {
      dom.parentNode.removeChild(dom);
    },
    // 设置多个样式
    "setStyles": function setStyles(dom, styles) {
      for (var key in styles) {
        dom.style[key] = styles[key];
      }
    },
    // 获取元素大小
    "size": function size(dom, type) {
      var elemHeight, elemWidth;

      if (type == 'content') {
        //内容
        elemWidth = dom.clientWidth - (this.getStyle(dom, 'padding-left') + "").replace('px', '') - (this.getStyle(dom, 'padding-right') + "").replace('px', '');
        elemHeight = dom.clientHeight - (this.getStyle(dom, 'padding-top') + "").replace('px', '') - (this.getStyle(dom, 'padding-bottom') + "").replace('px', '');
      } else if (type == 'padding') {
        //内容+内边距
        elemWidth = dom.clientWidth;
        elemHeight = dom.clientHeight;
      } else if (type == 'border') {
        //内容+内边距+边框
        elemWidth = dom.offsetWidth;
        elemHeight = dom.offsetHeight;
      } else if (type == 'scroll') {
        //滚动的宽（不包括border）
        elemWidth = dom.scrollWidth;
        elemHeight = dom.scrollHeight;
      } else {
        elemWidth = dom.offsetWidth;
        elemHeight = dom.offsetHeight;
      }

      return {
        width: elemWidth,
        height: elemHeight
      };
    },
    // 在被选元素内部的结尾插入内容
    "append": function append(el, template) {
      var node = this.isNode(template) ? template : this.toNode(template);
      el.appendChild(node);
      return node;
    },
    // 在被选元素内部的开头插入内容
    "prepend": function prepend(el, template) {
      var node = this.isNode(template) ? template : this.toNode(template);
      el.insertBefore(node, el.childNodes[0]);
      return node;
    },
    // 在被选元素之后插入内容
    "after": function after(el, template) {
      var node = this.isNode(template) ? template : this.toNode(template);
      el.parentNode.insertBefore(node, el.nextSibling);
      return node;
    },
    // 在被选元素之前插入内容
    "before": function before(el, template) {
      var node = this.isNode(template) ? template : this.toNode(template);
      el.parentNode.insertBefore(node, el);
      return node;
    }
  };

  function getTargetNode(event) {
    var _event = event || window.event;

    return _event.target || _event.srcElement;
  }

  function removeNode(node) {
    var pNode = node.parentNode;

    if (pNode) {
      pNode.removeChild(node);
    }
  } // 初始化结点


  function initDom() {
    this.__el.innerHTML = "";
    xhtml.setStyles(this.__el, {
      "background-color": "#f7f7f7",
      "user-select": "none"
    });
  }

  function itemInputHandler(event) {
    this.__contentArray[this.__tableIndex].content[+getTargetNode(event).getAttribute('row') - 1][+getTargetNode(event).getAttribute('col') - 1].value = getTargetNode(event).innerText;
  }

  function itemClickHandler(event) {
    // 如果格式刷按下了
    if (this.__format == true) {
      var rowNodes = xhtml.find(this.__contentDom[this.__tableIndex], function () {
        return true;
      }, 'tr');
      var targetStyle = this.__contentArray[this.__tableIndex].content[+getTargetNode(event).getAttribute('row') - 1][+getTargetNode(event).getAttribute('col') - 1].style;

      for (var row = this.__region.info.row[0]; row <= this.__region.info.row[1]; row++) {
        var colNodes = xhtml.find(rowNodes[row], function () {
          return true;
        }, 'th');

        for (var col = this.__region.info.col[0]; col <= this.__region.info.col[1]; col++) {
          // 遍历所有的样式
          for (var key in targetStyle) {
            // 修改界面显示
            colNodes[col].style[key] = targetStyle[key]; // 修改数据

            this.__contentArray[this.__tableIndex].content[row - 1][col - 1].style[key] = targetStyle[key];
          }
        }
      } // 取消标记格式刷


      this.__format = false;
      xhtml.removeClass(xhtml.find(this.__menuQuickDom, function (node) {
        return node.getAttribute('def-type') == 'format';
      }, 'span')[0], 'active');
    }

    this.$$moveCursorTo(getTargetNode(event), +getTargetNode(event).getAttribute('row'), +getTargetNode(event).getAttribute('col'));
  } // 初始化视图


  function initTableView(itemTable, index, styleToString) {
    var _this = this;

    var tableTemplate = ""; // 顶部的

    tableTemplate += "<tr><th class='top-left' open-web-excel></th>";

    for (var k = 0; k < itemTable.content[0].length; k++) {
      tableTemplate += "<th class='top-name' open-web-excel>" + this.$$calcColName(k) + "</th>";
    }

    tableTemplate += '</tr>'; // 行

    for (var i = 0; i < itemTable.content.length; i++) {
      tableTemplate += "<tr><th class='line-num' open-web-excel>" + (i + 1) + "</th>"; //  列

      for (var j = 0; j < itemTable.content[i].length; j++) {
        // contenteditable="true" 可编辑状态，则可点击获取焦点，同时内容也是可以编辑的
        // tabindex="0" 点击获取焦点，内容是不可编辑的
        tableTemplate += "<th\n                row='".concat(i + 1, "'\n                col='").concat(j + 1, "'\n                contenteditable=\"true\"\n                class=\"item\"\n                colspan=\"").concat(itemTable.content[i][j].colspan, "\"\n                rowspan=\"").concat(itemTable.content[i][j].rowspan, "\"\n                style=\"").concat(styleToString(itemTable.content[i][j].style), "\"\n            open-web-excel>").concat(itemTable.content[i][j].value, "</th>");
      }

      tableTemplate += '</tr>';
    }

    this.__contentDom[index] = xhtml.append(this.__tableFrame, "<table style='display:none;' class='excel-view' open-web-excel>" + tableTemplate + "</table>"); // 后续动态新增的需要重新绑定

    var items = xhtml.find(this.__contentDom[index], function (node) {
      return xhtml.hasClass(node, 'item');
    }, 'th');
    xhtml.bind(items, 'click', function (event) {
      _this.$$itemClickHandler(event);
    });
    xhtml.bind(items, 'input', function (event) {
      _this.$$itemInputHandler(event);
    });
  }

  var bottomClick = function bottomClick(target, index) {
    for (var i = 0; i < target.__contentDom.length; i++) {
      if (i == index) {
        xhtml.setStyles(target.__contentDom[i], {
          'display': 'table'
        });

        target.__btnDom[i].setAttribute('active', 'yes');
      } else {
        xhtml.setStyles(target.__contentDom[i], {
          'display': 'none'
        });

        target.__btnDom[i].setAttribute('active', 'no');
      }
    }

    target.__tableIndex = index;
    target.$$moveCursorTo(target.__contentDom[index].getElementsByTagName('tr')[1].getElementsByTagName('th')[1], 1, 1);
  };

  function initView() {
    var _this2 = this;

    this.__contentDom = [];
    this.__tableFrame = xhtml.append(this.__el, "<div></div>");
    xhtml.setStyles(this.__tableFrame, {
      "width": "100%",
      "height": "calc(100% - 92px)",
      "overflow": "auto"
    });

    for (var index = 0; index < this.__contentArray.length; index++) {
      this.$$initTableView(this.__contentArray[index], index, this.$$styleToString);
      xhtml.setStyles(this.__contentDom[index], {
        "display": index == 0 ? 'table' : "none"
      });
    }

    this.$$addStyle('excel-view', "\n\n        .excel-view{\n            border-collapse: collapse;\n            width: 100%;\n        }\n\n        .excel-view .top-left{\n            border: 1px solid #d6cccb;\n            border-right:none;\n            background-color:white;\n        }\n\n        .excel-view .top-name{\n            border: 1px solid #d6cccb;\n            border-bottom:none;\n            color:gray;\n            font-size:12px;\n        }\n\n        .excel-view .line-num{\n            padding:0 5px;\n            border: 1px solid #d6cccb;\n            border-right:none;\n            color:gray;\n            font-size:12px;\n        }\n\n        .excel-view .item{\n            min-width:50px;\n            white-space: nowrap;\n            border:0.5px solid rgba(85,85,85,0.5);\n            outline:none;\n            font-size:12px;\n            padding:2px;\n        }\n\n        .excel-view .item[active='yes']{\n            outline: 2px dashed red;\n        }\n\n    "); // 添加底部控制选择显示表格按钮

    var bottomBtns = xhtml.append(this.__el, "<div class='bottom-btn' open-web-excel></div>");
    var addBtn = xhtml.append(bottomBtns, "<span class='add item' open-web-excel>+</span>");
    xhtml.bind(addBtn, 'click', function () {
      // 首先，需要追加数据
      _this2.__contentArray.push(_this2.$$formatContent()[0]);

      var index = _this2.__contentArray.length - 1; // 然后添加table

      _this2.$$initTableView(_this2.__contentArray[index], index, _this2.$$styleToString); // 添加底部按钮


      var bottomBtn = xhtml.append(bottomBtns, "<span class='name item' open-web-excel>" + _this2.__contentArray[index].name + "</span>");

      _this2.__btnDom.push(bottomBtn);

      xhtml.bind(bottomBtn, 'click', function () {
        bottomClick(_this2, index);
      });
    });
    this.__btnDom = [];

    var _loop = function _loop(_index2) {
      var bottomBtn = xhtml.append(bottomBtns, "<span class='name item' open-web-excel>" + _this2.__contentArray[_index2].name + "</span>"); // 点击切换显示的视图

      xhtml.bind(bottomBtn, 'click', function () {
        bottomClick(_this2, _index2);
      }); // 双击可以修改名字

      xhtml.bind(bottomBtn, 'dblclick', function () {
        _this2.__btnDom[_index2].setAttribute('contenteditable', 'true');
      });
      xhtml.bind(bottomBtn, 'blur', function () {
        _this2.__contentArray[_index2].name = bottomBtn.innerText;
      }); // 登记起来所有的按钮

      _this2.__btnDom.push(bottomBtn);
    };

    for (var _index2 = 0; _index2 < this.__contentArray.length; _index2++) {
      _loop(_index2);
    }

    this.$$addStyle('bottom-btn', "\n\n        .bottom-btn{\n            width: 100%;\n            height: 30px;\n            overflow: auto;\n            border-top: 1px solid #d6cccb;\n            box-sizing: border-box;\n        }\n\n        .bottom-btn .item{\n            line-height: 30px;\n            box-sizing: border-box;\n            vertical-align: top;\n            display: inline-block;\n            cursor: pointer;\n        }\n\n        .bottom-btn .add{\n            width: 30px;\n            text-align: center;\n            font-size: 18px;\n        }\n\n        .bottom-btn .name{\n            font-size: 12px;\n            padding: 0 10px;\n        }\n        .bottom-btn .name:focus{\n            outline:none;\n        }\n\n        .bottom-btn .name:hover{\n            background-color:#efe9e9;\n        }\n\n        .bottom-btn .name[active='yes']{\n            background-color:white;\n        }\n\n    "); // 初始化点击第一个

    this.__btnDom[0].click();
  }

  function styleToString(style) {
    var styleString = "";

    for (var key in style) {
      styleString += key + ":" + style[key] + ';';
    }

    return styleString;
  }

  function newItemData() {
    return {
      value: " ",
      colspan: "1",
      rowspan: "1",
      style: {
        display: "table-cell",
        color: 'black',
        background: 'white',
        'vertical-align': 'top',
        'text-align': 'left',
        'font-weight': "normal",
        // bold粗体
        'font-style': 'normal',
        // italic斜体
        'text-decoration': 'none' // line-through中划线 underline下划线

      }
    };
  }

  function formatContent(file) {
    // 如果传递了内容
    if (file && 'version' in file && file.filename == 'Open-Web-Excel') {
      // 后续如果格式进行了升级，可以格式兼容转换成最新版本
      return file.contents;
    } // 否则，自动初始化
    else {
        var content = [];

        for (var i = 0; i < 100; i++) {
          var rowArray = [];

          for (var j = 0; j < 30; j++) {
            rowArray.push(this.$$newItemData());
          }

          content.push(rowArray);
        }

        return [{
          name: "未命名",
          content: content
        }];
      }
  }

  function calcColName(index) {
    if (!isNumber(index)) return index;
    var codes = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z'];
    var result = "";

    while (true) {
      // 求解当前坐标
      var _index = index % 26; // 拼接


      result = codes[_index] + result; // 求解余下的数

      index = Math.floor(index / 26);
      if (index == 0) break;
      index -= 1;
    }

    return result;
  }

  function getLeftTop(rowIndex, colIndex) {
    var content = this.__contentArray[this.__tableIndex].content; // 从下到上

    for (var row = rowIndex; row >= 1; row--) {
      // 从右到左
      for (var col = colIndex; col >= 1; col--) {
        // 同一行如果遇到第一个显示的，只有两种可能：
        // 1.这个就是所求
        // 2.本行都不会有结果
        if (content[row - 1][col - 1].style.display != 'none') {
          // 如果目标可以包含自己，那就找到了
          if (content[row - 1][col - 1].rowspan - -row > rowIndex && content[row - 1][col - 1].colspan - -col > colIndex) {
            return {
              row: row,
              col: col,
              content: content[row - 1][col - 1]
            };
          } else {
            break;
          }
        } // 不加else的原因是，理论上一定会存在唯一的一个

      }
    }
  }

  var addUniqueNamespace = function addUniqueNamespace(style) {
    var uniqueNameSpace = 'open-web-excel';
    style = style.replace(/( {0,}){/g, "{");
    style = style.replace(/( {0,}),/g, ",");
    var temp = ""; // 分别表示：是否处于注释中、是否处于内容中、是否由于特殊情况在遇到{前完成了hash

    var isSpecial = false,
        isContent = false,
        hadComplete = false;

    for (var i = 0; i < style.length; i++) {
      if (style[i] == ':' && !isSpecial && !hadComplete && !isContent) {
        hadComplete = true;
        temp += "[" + uniqueNameSpace + "]";
      } else if (style[i] == '{' && !isSpecial) {
        isContent = true;
        if (!hadComplete) temp += "[" + uniqueNameSpace + "]";
      } else if (style[i] == '}' && !isSpecial) {
        isContent = false;
        hadComplete = false;
      } else if (style[i] == '/' && style[i + 1] == '*') {
        isSpecial = true;
      } else if (style[i] == '*' && style[i + 1] == '/') {
        isSpecial = false;
      } else if (style[i] == ',' && !isSpecial && !isContent) {
        if (!hadComplete) temp += "[" + uniqueNameSpace + "]";
        hadComplete = false;
      }

      temp += style[i];
    }

    return temp;
  };

  function style() {
    if ('open-web-excel@style' in window) ;else {
      window['open-web-excel@style'] = {};
    }
    var head = document.head || document.getElementsByTagName('head')[0];
    return function (keyName, styleString) {
      if (window['open-web-excel@style'][keyName]) ;else {
        window['open-web-excel@style'][keyName] = true; // 创建style标签

        var styleElement = document.createElement('style');
        styleElement.setAttribute('type', 'text/css'); // 写入样式内容
        // 添加统一的后缀是防止污染

        styleElement.innerHTML = addUniqueNamespace("/*\n    Style[".concat(keyName, "] for Open-Web-Excel\n    https://www.npmjs.com/package/open-web-excel\n*/\n            ") + styleString); // 添加到页面

        head.appendChild(styleElement);
      }
    };
  } // 移动光标到指定位置


  function moveCursorTo(target, rowNum, colNum) {
    // 如果本来存在区域，应该取消
    if (this.__region != null) {
      this.$$cancelRegion();
      this.__region = null;
    } // 如果shift被按下，我们认为是在选择区间


    if (this.__keyLog.shift) {
      // 记录下来区域信息
      this.__region = this.$$calcRegionInfo({
        row: this.__rowNum,
        col: this.__colNum,
        rowNum: +this.__target.getAttribute('rowspan'),
        colNum: +this.__target.getAttribute('colspan')
      }, {
        row: rowNum,
        col: colNum,
        rowNum: +target.getAttribute('rowspan'),
        colNum: +target.getAttribute('colspan')
      });
      this.$$showRegion();
    } else {
      if (isElement(this.__target)) this.__target.setAttribute('active', 'no'); // 记录当前鼠标的位置

      this.__rowNum = rowNum;
      this.__colNum = colNum;
      this.__target = target; // 先获取对应的原始数据

      var oralItemData = this.__contentArray[this.__tableIndex].content[rowNum - 1][colNum - 1]; // 接着更新顶部菜单

      this.$$updateMenu(oralItemData.style);
      target.setAttribute('active', 'yes');
    }
  } // 修改默认输入条目的样式


  function setItemStyle(key, value) {
    // 更新数据内容
    this.__contentArray[this.__tableIndex].content[this.__rowNum - 1][this.__colNum - 1].style[key] = value; // 更新输入条目

    this.__target.style[key] = value; // 更新菜单状态

    this.$$updateMenu(this.__contentArray[this.__tableIndex].content[this.__rowNum - 1][this.__colNum - 1].style);
  } // 计算出区域的必要信息


  function calcRegionInfo(target1, target2) {
    var region = {
      // 区域的边界信息
      info: {},
      // 区域范围内的所有结点，第一个结点一定是左上角的那个
      nodes: []
    }; // 先计算出行边界

    var row1_min = target1.row;
    var row1_max = target1.row + target1.rowNum - 1;
    var row2_min = target2.row;
    var row2_max = target2.row + target2.rowNum - 1;
    var row_min = row1_min > row2_min ? row2_min : row1_min;
    var row_max = row1_max > row2_max ? row1_max : row2_max; // 再计算出列边界

    var col1_min = target1.col;
    var col1_max = target1.col + target1.colNum - 1;
    var col2_min = target2.col;
    var col2_max = target2.col + target2.colNum - 1;
    var col_min = col1_min > col2_min ? col2_min : col1_min;
    var col_max = col1_max > col2_max ? col1_max : col2_max; // 然后就可以标记区域的边界了

    region.info = {
      row: [row_min, row_max],
      col: [col_min, col_max]
    }; // 最后我们需要计算出此区域里面所有的结点

    var trs = this.__contentDom[this.__tableIndex].getElementsByTagName('tr');

    for (var i = row_min; i <= row_max; i++) {
      var ths = trs[i].getElementsByTagName('th');

      for (var j = 1; j < ths.length; j++) {
        var colValue = ths[j].getAttribute('col');

        if (colValue >= col_min && colValue <= col_max) {
          region.nodes.push(ths[j]);
        } else {
          // 判断是否可以提前结束
          if (colValue > col_max) {
            break;
          }
        }
      }
    }

    return region;
  } // 在页面中标记当前选择的区域


  function showRegion() {
    for (var i = 0; i < this.__region.nodes.length; i++) {
      this.__region.nodes[i].style.background = '#e5e0e0';
    }
  } // 取消在页面中标记的区域效果


  function cancelRegion() {
    for (var i = 0; i < this.__region.nodes.length; i++) {
      this.__region.nodes[i].style.background = this.__contentArray[this.__tableIndex].content[this.__region.nodes[i].getAttribute('row') - 1][this.__region.nodes[i].getAttribute('col') - 1].style.background;
    }
  }
  /*!
   * 💡 - 获取键盘此时按下的键的组合结果
   * https://github.com/hai2007/tool.js/blob/master/getKeyString.js
   *
   * author hai2007 < https://hai2007.gitee.io/sweethome >
   *
   * Copyright (c) 2021-present hai2007 走一步，再走一步。
   * Released under the MIT license
   */
  // 字典表


  var dictionary = {
    // 数字
    48: [0, ')'],
    49: [1, '!'],
    50: [2, '@'],
    51: [3, '#'],
    52: [4, '$'],
    53: [5, '%'],
    54: [6, '^'],
    55: [7, '&'],
    56: [8, '*'],
    57: [9, '('],
    96: [0, 0],
    97: 1,
    98: 2,
    99: 3,
    100: 4,
    101: 5,
    102: 6,
    103: 7,
    104: 8,
    105: 9,
    106: "*",
    107: "+",
    109: "-",
    110: ".",
    111: "/",
    // 字母
    65: ["a", "A"],
    66: ["b", "B"],
    67: ["c", "C"],
    68: ["d", "D"],
    69: ["e", "E"],
    70: ["f", "F"],
    71: ["g", "G"],
    72: ["h", "H"],
    73: ["i", "I"],
    74: ["j", "J"],
    75: ["k", "K"],
    76: ["l", "L"],
    77: ["m", "M"],
    78: ["n", "N"],
    79: ["o", "O"],
    80: ["p", "P"],
    81: ["q", "Q"],
    82: ["r", "R"],
    83: ["s", "S"],
    84: ["t", "T"],
    85: ["u", "U"],
    86: ["v", "V"],
    87: ["w", "W"],
    88: ["x", "X"],
    89: ["y", "Y"],
    90: ["z", "Z"],
    // 方向
    37: "left",
    38: "up",
    39: "right",
    40: "down",
    33: "page up",
    34: "page down",
    35: "end",
    36: "home",
    // 控制键
    16: "shift",
    17: "ctrl",
    18: "alt",
    91: "command",
    92: "command",
    93: "command",
    9: "tab",
    20: "caps lock",
    32: "spacebar",
    8: "backspace",
    13: "enter",
    27: "esc",
    46: "delete",
    45: "insert",
    144: "number lock",
    145: "scroll lock",
    12: "clear",
    19: "pause",
    // 功能键
    112: "f1",
    113: "f2",
    114: "f3",
    115: "f4",
    116: "f5",
    117: "f6",
    118: "f7",
    119: "f8",
    120: "f9",
    121: "f10",
    122: "f11",
    123: "f12",
    // 余下键
    189: ["-", "_"],
    187: ["=", "+"],
    219: ["[", "{"],
    221: ["]", "}"],
    220: ["\\", "|"],
    186: [";", ":"],
    222: ["'", '"'],
    188: [",", "<"],
    190: [".", ">"],
    191: ["/", "?"],
    192: ["`", "~"]
  }; // 非独立键字典

  var help_key = ["shift", "ctrl", "alt"];
  /**
   * 键盘按键
   * 返回键盘此时按下的键的组合结果
   */

  function getKeyString(event) {
    event = event || window.event;
    var keycode = event.keyCode || event.which;
    var key = dictionary[keycode] || keycode;
    if (!key) return;
    if (key.constructor !== Array) key = [key, key];
    var _key = key[0];
    var shift = event.shiftKey ? "shift+" : "",
        alt = event.altKey ? "alt+" : "",
        ctrl = event.ctrlKey ? "ctrl+" : "";
    var resultKey = "",
        preKey = ctrl + shift + alt;

    if (help_key.indexOf(key[0]) >= 0) {
      key[0] = key[1] = "";
    } // 判断是否按下了caps lock


    var lockPress = event.code == "Key" + event.key && !shift; // 只有字母（且没有按下功能Ctrl、shift或alt）区分大小写

    resultKey = preKey + (preKey == '' && lockPress ? key[1] : key[0]);

    if (key[0] == "") {
      resultKey = resultKey.replace(/\+$/, '');
    }

    return resultKey == '' ? _key : resultKey;
  } // 键盘总控


  function renderKeyboard() {
    var _this3 = this;

    if ('__keyLog' in this) {
      console.error('Keyboard has been initialized');
      return;
    } else {
      this.__keyLog = {
        'shift': false
      };
      xhtml.bind(document.body, 'keydown', function (event) {
        var keyString = getKeyString(event); // 标记shift按下

        if (keyString == 'shift') _this3.__keyLog.shift = true;
      });
      xhtml.bind(document.body, 'keyup', function (event) {
        var keyString = getKeyString(event); // 标记shift放开

        if (keyString == 'shift') _this3.__keyLog.shift = false;
      });
    }
  }

  var colors = [['白', 'rgb(255, 255, 255)'], ['漆黑', 'rgb(13, 0, 21)'], ['红', 'rgb(254, 44, 35)'], ['橙', 'rgb(255, 153, 0)'], ['黄', 'rgb(255, 217, 0)'], ['葱绿', 'rgb(163, 224, 67)'], ['湖蓝', 'rgb(55, 217, 240)'], ['天色', 'rgb(77, 168, 238)'], ['藤紫', 'rgb(149, 111, 231)'], ['白练', 'rgb(243, 243, 244)'], ['白鼠', 'rgb(204, 204, 204)'], ['樱', 'rgb(254, 242, 240)'], ['镐', 'rgb(254, 245, 231)'], ['练', 'rgb(254, 252, 217)'], ['芽', 'rgb(237, 246, 232)'], ['水', 'rgb(230, 250, 250)'], ['缥', 'rgb(235, 244, 252)'], ['丁香', 'rgb(240, 237, 246)'], ['灰青', 'rgb(215, 216, 217)'], ['鼠', 'rgb(165, 165, 165)'], ['虹', 'rgb(251, 212, 208)'], ['落柿', 'rgb(255, 215, 185)'], ['花叶', 'rgb(249, 237, 166)'], ['白绿', 'rgb(212, 233, 214)'], ['天青', 'rgb(199, 230, 234)'], ['天空', 'rgb(204, 224, 241)'], ['水晶', 'rgb(218, 213, 233)'], ['薄纯', 'rgb(123, 127, 131)'], ['墨', 'rgb(73, 73, 73)'], ['甚三红', 'rgb(238, 121, 118)'], ['珊瑚', 'rgb(250, 165, 115)'], ['金', 'rgb(230, 179, 34)'], ['薄青', 'rgb(152, 192, 145)'], ['白群', 'rgb(121, 198, 205)'], ['薄花', 'rgb(110, 170, 215)'], ['紫苑', 'rgb(156, 142, 193)'], ['石墨', 'rgb(65, 70, 75)'], ['黑', 'rgb(51, 51, 51)'], ['绯红', 'rgb(190, 26, 29)'], ['棕黄', 'rgb(185, 85, 20)'], ['土黄', 'rgb(173, 114, 14)'], ['苍翠', 'rgb(28, 114, 49)'], ['孔雀', 'rgb(28, 120, 146)'], ['琉璃', 'rgb(25, 67, 156)'], ['青莲', 'rgb(81, 27, 120)']];
  var template = "<div class='color-view' open-web-excel>";

  for (var i = 0; i < colors.length; i++) {
    template += "<div class='color-item' open-web-excel><span title='" + colors[i][0] + "' open-web-excel style='background:" + colors[i][1] + "'> </span></div>";
  }

  template += "</div>";
  var colorTemplate = template;

  function menu() {
    var _this4 = this;

    // 顶部操作栏
    var topDom = xhtml.append(this.__el, "<div class='top-dom' open-web-excel></div>");
    this.$$addStyle('top-dom', "\n\n       .top-dom{\n            width: 100%;\n            height: 62px;\n            overflow: hidden;\n       }\n\n    "); // 菜单

    this.__menuDom = xhtml.append(topDom, "<div class='menu' open-web-excel>\n        <span open-web-excel>\n            \u64CD\u4F5C\n            <div open-web-excel>\n                <span class='item more' open-web-excel>\n                    \u63D2\u5165\n                    <div open-web-excel>\n                        <span class='item' open-web-excel>\n                            <span def-type='insert-up'>\u5411\u4E0A\u63D2\u5165</span>\n                            <input value='1' open-web-excel />\n                            <span def-type='insert-up'>\u884C</span>\n                        </span>\n                        <span class='item' open-web-excel>\n                            <span def-type='insert-down'>\u5411\u4E0B\u63D2\u5165</span>\n                            <input value='1' open-web-excel />\n                            <span def-type='insert-down'>\u884C</span>\n                        </span>\n                        <span class='item' open-web-excel>\n                            <span def-type='insert-left'>\u5411\u5DE6\u63D2\u5165</span>\n                            <input value='1' open-web-excel />\n                            <span def-type='insert-left'>\u5217</span>\n                        </span>\n                        <span class='item' open-web-excel>\n                            <span def-type='insert-right'>\u5411\u53F3\u63D2\u5165</span>\n                            <input value='1' open-web-excel />\n                            <span def-type='insert-right'>\u5217</span>\n                        </span>\n                    </div>\n                </span>\n                <span class='item more' open-web-excel>\n                    \u5220\u9664\n                    <div open-web-excel>\n                        <span class='item' open-web-excel def-type='delete-row'>\n                            \u5220\u9664\u5F53\u524D\u884C\n                        </span>\n                        <span class='item' open-web-excel def-type='delete-col'>\n                            \u5220\u9664\u5F53\u524D\u5217\n                        </span>\n                    </div>\n                </span>\n                <span class='item more' open-web-excel>\n                    \u5408\u5E76\u5355\u5143\u683C\n                    <div open-web-excel>\n                        <span class='item' def-type='merge-all' open-web-excel>\u5168\u90E8\u5408\u5E76</span>\n                        <span class='item' def-type='merge-cancel' open-web-excel>\u53D6\u6D88\u5408\u5E76</span>\n                    </div>\n                </span>\n            </div>\n        </span>\n        <span open-web-excel>\n            \u683C\u5F0F\n            <div open-web-excel>\n                <span class='item' def-type='bold' open-web-excel>\u7C97\u4F53</span>\n                <span class='item' def-type='italic' open-web-excel>\u659C\u4F53</span>\n                <span class='item' def-type='underline' open-web-excel>\u4E0B\u5212\u7EBF</span>\n                <span class='item' def-type='line-through' open-web-excel>\u4E2D\u5212\u7EBF</span>\n                <span class='line' open-web-excel></span>\n                <span class='item more' open-web-excel>\n                    \u6C34\u5E73\u5BF9\u9F50\n                    <div open-web-excel>\n                        <span class='item' def-type='horizontal-left' open-web-excel>\u5DE6\u5BF9\u9F50</span>\n                        <span class='item' def-type='horizontal-center' open-web-excel>\u5C45\u4E2D\u5BF9\u9F50</span>\n                        <span class='item' def-type='horizontal-right' open-web-excel>\u53F3\u5BF9\u9F50</span>\n                    </div>\n                </span>\n                <span class='item more' open-web-excel>\n                    \u5782\u76F4\u5BF9\u9F50\n                    <div open-web-excel>\n                        <span class='item' def-type='vertical-top' open-web-excel>\u9876\u90E8\u5BF9\u9F50</span>\n                        <span class='item' def-type='vertical-middle' open-web-excel>\u5C45\u4E2D\u5BF9\u9F50</span>\n                        <span class='item' def-type='vertical-bottom' open-web-excel>\u5E95\u90E8\u5BF9\u9F50</span>\n                    </div>\n                </span>\n            </div>\n        </span>\n        <span open-web-excel>\n            \u5E2E\u52A9\n            <div open-web-excel>\n                <span class='item' open-web-excel>\n                    <a href='https://github.com/hai2007/Open-Web-Excel/issues' open-web-excel target='_blank'>\u95EE\u9898\u53CD\u9988</a>\n                </span>\n            </div>\n        </span>\n    </div>");
    this.$$addStyle('menu', "\n\n        .menu{\n            border-bottom: 1px solid #d6cccb;\n            padding: 0 20px;\n            box-sizing: border-box;\n            white-space: nowrap;\n        }\n\n        .menu>span{\n            display: inline-block;\n            line-height: 26px;\n            padding: 0 10px;\n            font-size: 12px;\n            cursor: pointer;\n            color: #555555;\n        }\n        .menu>span:hover{\n            background: white;\n        }\n\n        .menu>span>div{\n            margin-left: -10px;\n        }\n\n        .menu>span div{\n            position:absolute;\n            background: white;\n            width: 140px;\n            box-shadow: 4px 3px 6px 0 #c9c9e2;\n            display:none;\n            padding:5px 0;\n        }\n\n        .menu>span div span{\n            display:block;\n            position:relative;\n            padding:5px 20px;\n        }\n\n        .menu>span div span>div{\n            left:140px;\n            top:0px;\n        }\n\n        .menu .line{\n            height:1px;\n            background-color:#d6cccb;\n            padding:0;\n            margin:0 10px;\n        }\n\n        .menu input{\n            width:20px;\n            outline:none;\n        }\n\n        .menu span:hover>div{\n            display:block;\n        }\n\n        .menu span.more:after{\n            content:\">\";\n            position: absolute;\n            right: 12px;\n            font-weight: 800;\n        }\n\n        .menu a{\n            text-decoration: none;\n            color: #555555;\n        }\n\n        .menu input{\n            width:20px;\n            outline:none;\n        }\n\n        .menu .item.active::before{\n            content: \"*\";\n            color: red;\n            position: absolute;\n            left: 8px;\n        }\n\n        .menu .item{\n            text-decoration: none;\n        }\n\n        .menu .item:hover{\n            text-decoration: underline;\n        }\n\n    "); // 快捷菜单

    this.__menuQuickDom = xhtml.append(topDom, "<div class='quick-menu' open-web-excel>\n        <span class='item' def-type='format' open-web-excel>\u683C\u5F0F\u5316</span>\n        <span class='line' open-web-excel></span>\n        <span class='item color' def-type='font-color' open-web-excel>\n            \u6587\u5B57\u989C\u8272\uFF1A<i class='color' open-web-excel></i>\n            ".concat(colorTemplate, "\n        </span>\n        <span class='item color' def-type='background-color' open-web-excel>\n            \u586B\u5145\u8272\uFF1A<i class='color' open-web-excel></i>\n            ").concat(colorTemplate, "\n        </span>\n        <span class='line' open-web-excel></span>\n        <span class='item' def-type='merge-all' open-web-excel>\n            \u5168\u90E8\u5408\u5E76\n        </span>\n        <span class='item' def-type='merge-cancel' open-web-excel>\n            \u53D6\u6D88\u5408\u5E76\n        </span>\n        <span class='line' open-web-excel></span>\n        <span class='item' def-type='horizontal-left' open-web-excel>\n            \u5DE6\u5BF9\u9F50\n        </span>\n        <span class='item' def-type='horizontal-center' open-web-excel>\n            \u5C45\u4E2D\u5BF9\u9F50\n        </span>\n        <span class='item' def-type='horizontal-right' open-web-excel>\n            \u53F3\u5BF9\u9F50\n        </span>\n        <span class='line' open-web-excel></span>\n        <span class='item' def-type='vertical-top' open-web-excel>\n            \u9876\u90E8\u5BF9\u9F50\n        </span>\n        <span class='item' def-type='vertical-middle' open-web-excel>\n            \u5C45\u4E2D\u5BF9\u9F50\n        </span>\n        <span class='item' def-type='vertical-bottom' open-web-excel>\n            \u5E95\u90E8\u5BF9\u9F50\n        </span>\n    </div>"));
    this.$$addStyle('quick-menu', "\n\n        .quick-menu{\n            line-height: 36px;\n            font-size: 12px;\n            white-space: nowrap;\n            width: 100%;\n            overflow: auto;\n        }\n\n        .quick-menu span{\n            display:inline-block;\n            vertical-align: top;\n        }\n\n        .quick-menu span>i.color{\n            display: inline-block;\n            height: 14px;\n            width: 20px;\n            border:1px solid #d6cccb;\n            vertical-align: middle;\n        }\n\n        .quick-menu .item{\n            margin:0 10px;\n            cursor: pointer;\n        }\n\n        .quick-menu .line{\n            background-color:#d6cccb;\n            width:1px;\n            height:22px;\n            margin-top:7px;\n        }\n\n        .quick-menu .item:hover{\n            font-weight: 800;\n        }\n\n        .quick-menu .item.active{\n            font-weight: 800;\n            color: red;\n        }\n\n        /* \u9009\u62E9\u989C\u8272 */\n\n        .color-view{\n            font-size: 0px;\n            width: 171px;\n            position: absolute;\n            padding: 10px;\n            box-sizing: content-box;\n            background: #fefefe;\n            box-shadow: 1px 1px 5px #9e9695;\n            line-height:1em;\n            display:none;\n            margin-top: -5px;\n            white-space: normal;\n        }\n\n        .color:hover>.color-view, .color-view:hover{\n            display:block;\n        }\n\n        .color-item{\n            display: inline-block;\n            width: 19px;\n            height: 19px;\n        }\n\n        .color-item>span{\n            width: 15px;\n            height: 15px;\n            margin: 2px;\n            cursor: pointer;\n            box-sizing: border-box;\n        }\n\n        .color-item>span:hover{\n            outline:1px solid black;\n        }\n\n    "); // 对菜单添加点击事件

    var menuClickItems = xhtml.find(topDom, function (node) {
      return node.getAttribute('def-type');
    }, 'span');
    xhtml.bind(menuClickItems, 'click', function (event) {
      var node = getTargetNode(event); // 获取按钮类型

      var defType = node.getAttribute('def-type'); // 格式化

      if (defType == 'format') {
        // 首先需要确定选择区域，然后点击格式刷来同步格式
        if (_this4.__region != null) {
          // 标记格式刷
          _this4.__format = true;
          xhtml.addClass(xhtml.find(_this4.__menuQuickDom, function (node) {
            return node.getAttribute('def-type') == 'format';
          }, 'span')[0], 'active');
        }
      } // 粗体
      else if (defType == 'bold') {
          _this4.$$setItemStyle('font-weight', xhtml.hasClass(node, 'active') ? 'normal' : 'bold');
        } // 斜体
        else if (defType == 'italic') {
            _this4.$$setItemStyle('font-style', xhtml.hasClass(node, 'active') ? 'normal' : 'italic');
          } // 中划线
          else if (defType == 'line-through') {
              _this4.$$setItemStyle('text-decoration', xhtml.hasClass(node, 'active') ? 'none' : 'line-through');
            } // 下划线
            else if (defType == 'underline') {
                _this4.$$setItemStyle('text-decoration', xhtml.hasClass(node, 'active') ? 'none' : 'underline');
              } // 水平对齐方式
              else if (/^horizontal\-/.test(defType)) {
                  _this4.$$setItemStyle('text-align', defType.replace('horizontal-', ''));
                } // 垂直对齐方式
                else if (/^vertical\-/.test(defType)) {
                    _this4.$$setItemStyle('vertical-align', defType.replace('vertical-', ''));
                  } // 合并单元格
                  else if (/^merge\-/.test(defType)) {
                      // 无选择区域，直接结束
                      if (_this4.__region == null) return; // 全部合并

                      if (defType == 'merge-all') {
                        // 如果选择的区域就一个结点，不用额外的操作了
                        if (_this4.__region.nodes.length <= 1) return; // 删除多余的结点并修改数据

                        for (var _i = 1; _i < _this4.__region.nodes.length; _i++) {
                          _this4.__contentArray[_this4.__tableIndex].content[_this4.__region.nodes[_i].getAttribute('row') - 1][_this4.__region.nodes[_i].getAttribute('col') - 1].style.display = 'none';
                          _this4.__contentArray[_this4.__tableIndex].content[_this4.__region.nodes[_i].getAttribute('row') - 1][_this4.__region.nodes[_i].getAttribute('col') - 1].value = ' ';
                          _this4.__region.nodes[_i].style.display = 'none';
                        }

                        _this4.__region.nodes = [_this4.__region.nodes[0]]; // 修改第一个结点的数据和占位

                        _this4.__contentArray[_this4.__tableIndex].content[_this4.__region.nodes[0].getAttribute('row') - 1][_this4.__region.nodes[0].getAttribute('col') - 1].colspan = _this4.__region.info.col[1] - _this4.__region.info.col[0] + 1 + "";
                        _this4.__contentArray[_this4.__tableIndex].content[_this4.__region.nodes[0].getAttribute('row') - 1][_this4.__region.nodes[0].getAttribute('col') - 1].rowspan = _this4.__region.info.row[1] - _this4.__region.info.row[0] + 1 + "";

                        _this4.__region.nodes[0].setAttribute('colspan', _this4.__region.info.col[1] - _this4.__region.info.col[0] + 1 + "");

                        _this4.__region.nodes[0].setAttribute('rowspan', _this4.__region.info.row[1] - _this4.__region.info.row[0] + 1 + "");

                        _this4.__region.nodes[0].click();
                      } // 取消合并
                      else if (defType == 'merge-cancel') {
                          var rowNodes = xhtml.find(_this4.__contentDom[_this4.__tableIndex], function () {
                            return true;
                          }, 'tr'); // 确保所有的格子都是 1*1 的

                          for (var row = _this4.__region.info.row[0]; row <= _this4.__region.info.row[1]; row++) {
                            var colNodes = xhtml.find(rowNodes[row], function () {
                              return true;
                            }, 'th');

                            for (var col = _this4.__region.info.col[0]; col <= _this4.__region.info.col[1]; col++) {
                              // 修改界面显示
                              colNodes[col].style.display = 'table-cell';
                              colNodes[col].setAttribute('colspan', '1');
                              colNodes[col].setAttribute('rowspan', '1'); // 修改数据

                              _this4.__contentArray[_this4.__tableIndex].content[row - 1][col - 1].style.display = 'table-cell';
                              _this4.__contentArray[_this4.__tableIndex].content[row - 1][col - 1].colspan = '1';
                              _this4.__contentArray[_this4.__tableIndex].content[row - 1][col - 1].rowspan = '1';
                            }
                          }

                          _this4.$$cancelRegion();

                          _this4.__region = null;
                        }
                    } // 插入
                    else if (/^insert\-/.test(defType)) {
                        var num = +xhtml.find(node.parentNode, function () {
                          return true;
                        }, 'input')[0].value; // 向上插入行

                        if (defType == 'insert-up') {
                          for (var _i2 = 0; _i2 < num; _i2++) {
                            _this4.$$insertUpNewRow();
                          }
                        } // 向下插入行
                        else if (defType == 'insert-down') {
                            for (var _i3 = 0; _i3 < num; _i3++) {
                              _this4.$$insertDownNewRow();
                            }
                          } // 向左插入列
                          else if (defType == 'insert-left') {
                              for (var _i4 = 0; _i4 < num; _i4++) {
                                _this4.$$insertLeftNewCol();
                              }
                            } // 向右插入列
                            else if (defType == 'insert-right') {
                                for (var _i5 = 0; _i5 < num; _i5++) {
                                  _this4.$$insertRightNewCol();
                                }
                              }
                      } // 删除
                      else if (/^delete\-/.test(defType)) {
                          // 删除当前行
                          if (defType == 'delete-row') {
                            _this4.$$deleteCurrentRow();
                          } // 删除当前列
                          else if (defType == 'delete-col') {
                              _this4.$$deleteCurrentCol();
                            }
                        }
    }); // 对选择颜色添加点击事件

    var colorItems = xhtml.find(topDom, function (node) {
      return xhtml.hasClass(node, 'color');
    }, 'span');

    var _loop2 = function _loop2(_i6) {
      var colorClickItems = xhtml.find(colorItems[_i6], function () {
        return true;
      }, 'span');
      xhtml.bind(colorClickItems, 'click', function (event) {
        var defType = colorItems[_i6].getAttribute('def-type');

        var colorValue = getTargetNode(event).style.background; // 设置

        _this4.$$setItemStyle({
          'background-color': 'background',
          'font-color': 'color'
        }[defType], colorValue);
      });
    };

    for (var _i6 = 0; _i6 < colorItems.length; _i6++) {
      _loop2(_i6);
    }
  }

  function updateMenu(style) {
    // 更新顶部菜单
    var menuItems = xhtml.find(this.__menuDom, function (node) {
      return node.getAttribute('def-type');
    }, 'span');

    for (var _i7 = 0; _i7 < menuItems.length; _i7++) {
      // 获取按钮类型
      var defType = menuItems[_i7].getAttribute('def-type'); // 粗体


      if (defType == 'bold') {
        if (style['font-weight'] == 'bold') {
          xhtml.addClass(menuItems[_i7], 'active');
        } else {
          xhtml.removeClass(menuItems[_i7], 'active');
        }
      } // 粗体
      else if (defType == 'italic') {
          if (style['font-style'] == 'italic') {
            xhtml.addClass(menuItems[_i7], 'active');
          } else {
            xhtml.removeClass(menuItems[_i7], 'active');
          }
        } // 中划线
        else if (defType == 'underline') {
            if (style['text-decoration'] == 'underline') {
              xhtml.addClass(menuItems[_i7], 'active');
            } else {
              xhtml.removeClass(menuItems[_i7], 'active');
            }
          } // 下划线
          else if (defType == 'line-through') {
              if (style['text-decoration'] == 'line-through') {
                xhtml.addClass(menuItems[_i7], 'active');
              } else {
                xhtml.removeClass(menuItems[_i7], 'active');
              }
            } // 水平对齐方式
            else if (/^horizontal\-/.test(defType)) {
                if (defType == 'horizontal-' + style['text-align']) {
                  xhtml.addClass(menuItems[_i7], 'active');
                } else {
                  xhtml.removeClass(menuItems[_i7], 'active');
                }
              } // 垂直对齐方式
              else if (/^vertical\-/.test(defType)) {
                  if (defType == 'vertical-' + style['vertical-align']) {
                    xhtml.addClass(menuItems[_i7], 'active');
                  } else {
                    xhtml.removeClass(menuItems[_i7], 'active');
                  }
                }
    } // 更新快速使用菜单


    var quickItems = xhtml.find(this.__menuQuickDom, function (node) {
      return node.getAttribute('def-type');
    }, 'span');

    for (var _i8 = 0; _i8 < quickItems.length; _i8++) {
      // 获取按钮类型
      var _defType = quickItems[_i8].getAttribute('def-type'); // 文字颜色


      if (_defType == 'font-color') {
        quickItems[_i8].getElementsByTagName('i')[0].style.backgroundColor = style.color;
      } // 填充色
      else if (_defType == 'background-color') {
          quickItems[_i8].getElementsByTagName('i')[0].style.backgroundColor = style.background;
        } // 水平对齐方式
        else if (/^horizontal\-/.test(_defType)) {
            if (_defType == 'horizontal-' + style['text-align']) {
              xhtml.addClass(quickItems[_i8], 'active');
            } else {
              xhtml.removeClass(quickItems[_i8], 'active');
            }
          } // 垂直对齐方式
          else if (/^vertical\-/.test(_defType)) {
              if (_defType == 'vertical-' + style['vertical-align']) {
                xhtml.addClass(quickItems[_i8], 'active');
              } else {
                xhtml.removeClass(quickItems[_i8], 'active');
              }
            }
    }
  }

  function insertUp() {
    var _this5 = this;

    var rowNodes = xhtml.find(this.__contentDom[this.__tableIndex], function () {
      return true;
    }, 'tr'); // 首先，直接在插入行前面插入一行

    var newRowNode = xhtml.before(rowNodes[this.__rowNum], '<tr><th class="line-num" open-web-excel>' + this.__rowNum + '</th></tr>');
    rowNodes.splice(this.__rowNum, 0, newRowNode);

    this.__contentArray[this.__tableIndex].content.splice(this.__rowNum - 1, 0, []); // 然后，校对数据


    for (var row = this.__rowNum + 1; row <= rowNodes.length - 1; row++) {
      var colNodes = xhtml.find(rowNodes[row], function () {
        return true;
      }, 'th'); // 修改行数

      colNodes[0].innerText = row; // 依次修改记录的行数

      for (var col = 1; col < colNodes.length; col++) {
        colNodes[col].setAttribute('row', row);
      }
    }

    for (var _col = 1; _col <= this.__contentArray[this.__tableIndex].content[this.__rowNum == 1 ? 1 : 0].length; _col++) {
      // 获取新的数据
      var tempNewItemData = this.$$newItemData();
      /**
       * 嗅探当前单元格情况，
       * 由于会出现合并单元格情况，所以需要对一些特殊情况，进行特殊校对
       */

      var currentItemData = this.__contentArray[this.__tableIndex].content[this.__rowNum][_col - 1]; //  如果不是第一行，而且自己不可见

      if (this.__rowNum != 1 && currentItemData.style.display == 'none') {
        // 那么，我们现在需要确定我们当前行是否位于合并单元格的顶部
        // 因为，如果自己位于顶部，即使不可见，依旧应该可以向上新增一行而不是增高自己
        // 如何直接自己是不是顶部？
        // 我们可以不停的嗅探左边第一个显示的单元格，如果他可以囊括自己，那自己应该就是上顶部
        // 否则就是非第一行
        var isFirstLine = false;

        for (var toLeftCol = _col - 1; toLeftCol >= 1; toLeftCol--) {
          var leftItemData = this.__contentArray[this.__tableIndex].content[this.__rowNum][toLeftCol - 1];

          if (leftItemData.style.display != 'none') {
            // 如果找到的第一个显示的可以包含当前条目
            if (toLeftCol - -leftItemData.colspan > _col) isFirstLine = true;
            break;
          }
        } // 如果是第一行我们就可以直接放过


        if (!isFirstLine) {
          // 到目前为止，我们可以确定的是，当前新增的条目需要隐藏
          tempNewItemData.style.display = 'none'; // 判断是不是最左边的

          var isLeftFirst = _col == 1 || this.__contentArray[this.__tableIndex].content[this.__rowNum][_col - 2].style.display != 'none'; // 如果是最坐标的，就需要负责修改左上角格子的值

          if (isLeftFirst) {
            for (var preRow = this.__rowNum - 1; preRow > 0; preRow--) {
              // 接着，让我们寻找这个条目合并后单元格的左上角
              if (this.__contentArray[this.__tableIndex].content[preRow - 1][_col - 1].style.display != 'none') {
                // 数据
                this.__contentArray[this.__tableIndex].content[preRow - 1][_col - 1].rowspan -= -1; // 结点

                var leftTopNode = xhtml.find(rowNodes[preRow], function () {
                  return true;
                }, 'th')[_col];

                leftTopNode.setAttribute('rowspan', leftTopNode.getAttribute('rowspan') - -1); // 找到以后别忘了停止

                break;
              }
            }
          }
        }
      } // 追加数据


      this.__contentArray[this.__tableIndex].content[this.__rowNum - 1].push(tempNewItemData); // 追加结点


      var newItemNode = xhtml.append(newRowNode, "<th row=\"".concat(this.__rowNum, "\" col=\"").concat(_col, "\" contenteditable=\"true\" class=\"item\" colspan=\"1\" rowspan=\"1\" style=\"").concat(this.$$styleToString(tempNewItemData.style), "\" open-web-excel></th>")); // 绑定事件

      xhtml.bind(newItemNode, 'click', function (event) {
        _this5.$$itemClickHandler(event);
      });
      xhtml.bind(newItemNode, 'input', function (event) {
        _this5.$$itemInputHandler(event);
      });
    } // 最后标记下沉


    this.__rowNum += 1;
  }

  function insertDown() {
    var _this6 = this;

    var rowNodes = xhtml.find(this.__contentDom[this.__tableIndex], function () {
      return true;
    }, 'tr'); // 首先，直接在插入行前面插入一行

    var newRowNode = xhtml.after(rowNodes[this.__rowNum], '<tr><th class="line-num" open-web-excel>' + (this.__rowNum + 1) + '</th></tr>');
    rowNodes.splice(this.__rowNum + 1, 0, newRowNode);

    this.__contentArray[this.__tableIndex].content.splice(this.__rowNum, 0, []); // 然后，校对数据


    for (var row = this.__rowNum + 2; row <= rowNodes.length - 1; row++) {
      var colNodes = xhtml.find(rowNodes[row], function () {
        return true;
      }, 'th'); // 修改行数

      colNodes[0].innerText = row; // 依次修改记录的行数

      for (var col = 1; col < colNodes.length; col++) {
        colNodes[col].setAttribute('row', row);
      }
    }

    for (var _col2 = 1; _col2 <= this.__contentArray[this.__tableIndex].content[0].length; _col2++) {
      // 获取新的数据
      var tempNewItemData = this.$$newItemData();
      /**
       * 对当前单元格合并情况进行嗅探
       */
      //  如果不是最后一行

      if (this.__rowNum != this.__contentArray[this.__tableIndex].content.length - 1) {
        var currentItemData = this.__contentArray[this.__tableIndex].content[this.__rowNum - 1][_col2 - 1]; // 不可见或行数不为1

        if (currentItemData.style.display == 'none' || currentItemData.rowspan != '1') {
          // 为了可以之前当前插入点的相对位置，我们首先需要找到合并后单元格左上角的数据和位置
          var leftTopData = this.$$getLeftTop(this.__rowNum, _col2); // 如果不是最底部一行

          if (leftTopData.row - -leftTopData.content.rowspan - 1 > this.__rowNum) {
            // 到此为止，可以确定当前的条目一定隐藏
            tempNewItemData.style.display = 'none'; // 如果是最左边的

            if (leftTopData.col == _col2) {
              // 数据
              this.__contentArray[this.__tableIndex].content[leftTopData.row - 1][leftTopData.col - 1].rowspan -= -1; // 结点

              var leftTopNode = xhtml.find(rowNodes[leftTopData.row], function () {
                return true;
              }, 'th')[leftTopData.col];
              leftTopNode.setAttribute('rowspan', leftTopNode.getAttribute('rowspan') - -1);
            }
          }
        }
      } // 追加数据


      this.__contentArray[this.__tableIndex].content[this.__rowNum].push(tempNewItemData); // 追加结点


      var newItemNode = xhtml.append(newRowNode, "<th row=\"".concat(this.__rowNum, "\" col=\"").concat(_col2, "\" contenteditable=\"true\" class=\"item\" colspan=\"1\" rowspan=\"1\" style=\"").concat(this.$$styleToString(tempNewItemData.style), "\" open-web-excel></th>")); // 绑定事件

      xhtml.bind(newItemNode, 'click', function (event) {
        _this6.$$itemClickHandler(event);
      });
      xhtml.bind(newItemNode, 'input', function (event) {
        _this6.$$itemInputHandler(event);
      });
    }
  }

  function insertLeft() {
    var _this7 = this;

    var rowNodes = xhtml.find(this.__contentDom[this.__tableIndex], function () {
      return true;
    }, 'tr'); // 先修改顶部的位置提示

    xhtml.append(rowNodes[0], "<th class='top-name' open-web-excel>" + this.$$calcColName(this.__contentArray[this.__tableIndex].content[0].length) + "</th>");

    for (var row = 1; row < rowNodes.length; row++) {
      var colNodes = xhtml.find(rowNodes[row], function () {
        return true;
      }, 'th'); // 校对列序号

      for (var col = this.__colNum; col < colNodes.length; col++) {
        colNodes[col].setAttribute('col', col + 1);
      } // 获取新的数据


      var tempNewItemData = this.$$newItemData();
      /**
       * 对当前单元格合并情况进行嗅探
       */

      var currentItemData = this.__contentArray[this.__tableIndex].content[row - 1][this.__colNum - 1]; //  如果不是第一列，而且自己不可见

      if (this.__colNum != 1 && currentItemData.style.display == 'none') {
        var isFirstCol = false;

        for (var toTopRow = row - 1; toTopRow >= 1; toTopRow--) {
          var topItemData = this.__contentArray[this.__tableIndex].content[toTopRow - 1][this.__colNum];

          if (topItemData.style.display != 'none') {
            // 如果找到的第一个显示的可以包含当前条目
            if (toTopRow - -topItemData.rowspan > row) isFirstCol = true;
            break;
          }
        } // 如果是第一列我们就可以直接放过


        if (!isFirstCol) {
          tempNewItemData.style.display = 'none'; // 判断是不是最顶部的

          var isTopFirst = row == 1 || this.__contentArray[this.__tableIndex].content[row - 2][this.__colNum].style.display != 'none'; // 如果是最坐标的，就需要负责修改左上角格子的值

          if (isTopFirst) {
            for (var preCol = this.__colNum - 1; preCol > 0; preCol--) {
              // 接着，让我们寻找这个条目合并后单元格的左上角
              if (this.__contentArray[this.__tableIndex].content[row - 1][preCol - 1].style.display != 'none') {
                // 数据
                this.__contentArray[this.__tableIndex].content[row - 1][preCol - 1].colspan -= -1; // 结点

                var leftTopNode = xhtml.find(rowNodes[row], function () {
                  return true;
                }, 'th')[preCol];
                leftTopNode.setAttribute('colspan', leftTopNode.getAttribute('colspan') - -1); // 找到以后别忘了停止

                break;
              }
            }
          }
        }
      } // 追加数据


      this.__contentArray[this.__tableIndex].content[row - 1].splice(this.__colNum - 1, 0, tempNewItemData); // 追加结点


      var newItemNode = xhtml.before(colNodes[this.__colNum], "<th row=\"".concat(row, "\" col=\"").concat(this.__colNum, "\" contenteditable=\"true\" class=\"item\" colspan=\"1\" rowspan=\"1\" style=\"").concat(this.$$styleToString(tempNewItemData.style), "\" open-web-excel></th>")); // 绑定事件

      xhtml.bind(newItemNode, 'click', function (event) {
        _this7.$$itemClickHandler(event);
      });
      xhtml.bind(newItemNode, 'input', function (event) {
        _this7.$$itemInputHandler(event);
      });
    } // 最后标记右移


    this.__colNum += 1;
  }

  function insertRight() {
    var _this8 = this;

    var rowNodes = xhtml.find(this.__contentDom[this.__tableIndex], function () {
      return true;
    }, 'tr'); // 先修改顶部的位置提示

    xhtml.append(rowNodes[0], "<th class='top-name' open-web-excel>" + this.$$calcColName(this.__contentArray[this.__tableIndex].content[0].length) + "</th>");

    for (var row = 1; row < rowNodes.length; row++) {
      var colNodes = xhtml.find(rowNodes[row], function () {
        return true;
      }, 'th'); // 校对列序号

      for (var col = this.__colNum + 1; col < colNodes.length; col++) {
        colNodes[col].setAttribute('col', col + 1);
      } // 获取新的数据


      var tempNewItemData = this.$$newItemData();
      /**
      * 对当前单元格合并情况进行嗅探
      */
      //  如果不是最后一列

      if (this.__colNum != this.__contentArray[this.__tableIndex].content[0].length - 1) {
        var currentItemData = this.__contentArray[this.__tableIndex].content[row - 1][this.__colNum - 1]; // 不可见或列数不为1

        if (currentItemData.style.display == 'none' || currentItemData.colspan != '1') {
          // 为了可以之前当前插入点的相对位置，我们首先需要找到合并后单元格左上角的数据和位置
          var leftTopData = this.$$getLeftTop(row, this.__colNum); // 如果不是最右边一列

          if (leftTopData.col - -leftTopData.content.colspan - 1 > this.__colNum) {
            // 到此为止，可以确定当前的条目一定隐藏
            tempNewItemData.style.display = 'none'; // 如果是最顶部的

            if (leftTopData.row == row) {
              // 数据
              this.__contentArray[this.__tableIndex].content[leftTopData.row - 1][leftTopData.col - 1].colspan -= -1; // 结点

              var leftTopNode = xhtml.find(rowNodes[leftTopData.row], function () {
                return true;
              }, 'th')[leftTopData.col];
              leftTopNode.setAttribute('colspan', leftTopNode.getAttribute('colspan') - -1);
            }
          }
        }
      } // 追加数据


      this.__contentArray[this.__tableIndex].content[row - 1].splice(this.__colNum, 0, tempNewItemData); // 追加结点


      var newItemNode = xhtml.after(colNodes[this.__colNum], "<th row=\"".concat(row, "\" col=\"").concat(this.__colNum + 1, "\" contenteditable=\"true\" class=\"item\" colspan=\"1\" rowspan=\"1\" style=\"").concat(this.$$styleToString(tempNewItemData.style), "\" open-web-excel></th>")); // 绑定事件

      xhtml.bind(newItemNode, 'click', function (event) {
        _this8.$$itemClickHandler(event);
      });
      xhtml.bind(newItemNode, 'input', function (event) {
        _this8.$$itemInputHandler(event);
      });
    }
  }

  function deleteRow() {
    var rowNodes = xhtml.find(this.__contentDom[this.__tableIndex], function () {
      return true;
    }, 'tr'); // 校对行号

    for (var row = this.__rowNum + 1; row <= this.__contentArray[this.__tableIndex].content.length; row++) {
      var colNodes = xhtml.find(rowNodes[row], function () {
        return true;
      }, 'th'); // 修改行数

      colNodes[0].innerText = row - 1; // 依次修改记录的行数

      for (var col = 1; col < colNodes.length; col++) {
        colNodes[col].setAttribute('row', row - 1);
      }
    }

    var isLastLine = this.__rowNum == this.__contentArray[this.__tableIndex].content.length; // 是否是最后一行

    var downColNodes;
    if (!isLastLine) downColNodes = xhtml.find(rowNodes[this.__rowNum + 1], function () {
      return true;
    }, 'th'); // 校对colspan

    for (var _col3 = 1; _col3 <= this.__contentArray[this.__tableIndex].content[0].length; _col3++) {
      // 如果当前条目隐藏
      if (this.__contentArray[this.__tableIndex].content[this.__rowNum - 1][_col3 - 1].style.display == 'none') {
        // 隐藏的话，就只需要考虑位于左上角的正下方情况
        for (var preRow = this.__rowNum - 1; preRow >= 1; preRow--) {
          if (this.__contentArray[this.__tableIndex].content[preRow - 1][_col3 - 1].style.display != 'none') {
            // 如果是左上角
            if (preRow - -this.__contentArray[this.__tableIndex].content[preRow - 1][_col3 - 1].rowspan > this.__rowNum) {
              var newRowspan = this.__contentArray[this.__tableIndex].content[preRow - 1][_col3 - 1].rowspan - 1; // 结点

              xhtml.find(rowNodes[preRow], function () {
                return true;
              }, 'th')[_col3].setAttribute('rowspan', newRowspan); // 数据


              this.__contentArray[this.__tableIndex].content[preRow - 1][_col3 - 1].rowspan = newRowspan;
            }

            break;
          }
        }
      } // 如果没有隐藏，可是是左上角
      // (如果是一行肯定可以直接无视)
      else if (this.__contentArray[this.__tableIndex].content[this.__rowNum - 1][_col3 - 1].rowspan - 1 > 0) {
          var _newRowspan = this.__contentArray[this.__tableIndex].content[this.__rowNum - 1][_col3 - 1].rowspan - 1;

          var colspan = this.__contentArray[this.__tableIndex].content[this.__rowNum - 1][_col3 - 1].colspan; // 结点

          downColNodes[_col3].setAttribute('rowspan', _newRowspan);

          downColNodes[_col3].setAttribute('colspan', colspan);

          downColNodes[_col3].style.display = 'table-cell'; // 数据

          this.__contentArray[this.__tableIndex].content[this.__rowNum][_col3 - 1].rowspan = _newRowspan;
          this.__contentArray[this.__tableIndex].content[this.__rowNum][_col3 - 1].colspan = colspan;
          this.__contentArray[this.__tableIndex].content[this.__rowNum][_col3 - 1].style.display = 'table-cell';
        }
    } // 删除当前行


    removeNode(rowNodes[this.__rowNum]); // 删除数据

    this.__contentArray[this.__tableIndex].content.splice(this.__rowNum - 1, 1); // 重置光标


    this.__btnDom[this.__tableIndex].click();
  }

  function deleteCol() {
    var rowNodes = xhtml.find(this.__contentDom[this.__tableIndex], function () {
      return true;
    }, 'tr'); // 校对rowspan

    for (var row = 1; row <= this.__contentArray[this.__tableIndex].content.length; row++) {
      // 如果当前条目隐藏
      if (this.__contentArray[this.__tableIndex].content[row - 1][this.__colNum - 1].style.display == 'none') {
        for (var preCol = this.__colNum - 1; preCol >= 1; preCol--) {
          if (this.__contentArray[this.__tableIndex].content[row - 1][preCol - 1].style.display != 'none') {
            // 如果是左上角
            if (preCol - -this.__contentArray[this.__tableIndex].content[row - 1][preCol - 1].colspan > this.__colNum) {
              var newColspan = this.__contentArray[this.__tableIndex].content[row - 1][preCol - 1].colspan - 1; // 结点

              xhtml.find(rowNodes[row], function () {
                return true;
              }, 'th')[preCol].setAttribute('colspan', newColspan); // 数据

              this.__contentArray[this.__tableIndex].content[row - 1][preCol - 1].colspan = newColspan;
            }

            break;
          }
        }
      } //  左上角
      else if (this.__contentArray[this.__tableIndex].content[row - 1][this.__colNum - 1].colspan - 1 > 0) {
          var nextColNode = xhtml.find(rowNodes[row], function () {
            return true;
          }, 'th')[this.__colNum + 1];

          var _newColspan = this.__contentArray[this.__tableIndex].content[row - 1][this.__colNum - 1].colspan - 1;

          var rowspan = this.__contentArray[this.__tableIndex].content[row - 1][this.__colNum - 1].rowspan; // 结点

          nextColNode.setAttribute('colspan', _newColspan);
          nextColNode.setAttribute('rowspan', rowspan);
          nextColNode.style.display = 'table-cell'; // 数据

          this.__contentArray[this.__tableIndex].content[row - 1][this.__colNum].colspan = _newColspan;
          this.__contentArray[this.__tableIndex].content[row - 1][this.__colNum].rowspan = rowspan;
          this.__contentArray[this.__tableIndex].content[row - 1][this.__colNum].style.display = 'table-cell';
        }
    } // 先删除列标题


    xhtml.find(rowNodes[0], function () {
      return true;
    }, 'th')[this.__contentArray[this.__tableIndex].content[0].length].remove();

    for (var _row = 1; _row < rowNodes.length; _row++) {
      var colNodes = xhtml.find(rowNodes[_row], function () {
        return true;
      }, 'th'); // 校对列序号

      for (var col = this.__colNum + 1; col < colNodes.length; col++) {
        colNodes[col].setAttribute('col', col - 1);
      } // 删除当前光标所在列


      removeNode(colNodes[this.__colNum]); // 数据也要删除

      this.__contentArray[this.__tableIndex].content[_row - 1].splice(this.__colNum - 1, 1);
    } // 重置光标


    this.__btnDom[this.__tableIndex].click();
  }

  var owe = function owe(options) {
    var _this9 = this;

    if (!(this instanceof owe)) {
      throw new Error('Open-Web-Excel is a constructor and should be called with the `new` keyword');
    } // 编辑器挂载点


    if (isElement(options.el)) {
      this.__el = options.el; // 内容

      this.__contentArray = this.$$formatContent(options.content); // 用于选择记录区域

      this.__region = null; // 用于记录是否按下了格式刷按钮

      this.__format = false;
    } else {
      // 挂载点是必须的，一定要有
      throw new Error('options.el is not a element!');
    } // 启动键盘事件监听


    this.$$renderKeyboard(); // 先初始化DOM

    this.$$initDom(); // 挂载菜单

    this.$$createdMenu(); // 初始化视图

    this.$$initView(); // 获取当前Excel内容

    this.valueOf = function () {
      return {
        version: "v1",
        filename: "Open-Web-Excel",
        contents: _this9.__contentArray
      };
    };
  }; // 挂载辅助方法


  owe.prototype.$$formatContent = formatContent;
  owe.prototype.$$calcColName = calcColName;
  owe.prototype.$$addStyle = style();
  owe.prototype.$$styleToString = styleToString;
  owe.prototype.$$newItemData = newItemData;
  owe.prototype.$$itemClickHandler = itemClickHandler;
  owe.prototype.$$itemInputHandler = itemInputHandler;
  owe.prototype.$$getLeftTop = getLeftTop; // 挂载核心方法

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
  owe.prototype.$$deleteCurrentCol = deleteCol; // 挂载键盘交互总控

  owe.prototype.$$renderKeyboard = renderKeyboard;

  if ((typeof module === "undefined" ? "undefined" : _typeof(module)) === "object" && _typeof(module.exports) === "object") {
    module.exports = owe;
  } else {
    window.OpenWebExcel = owe;
  }
})();