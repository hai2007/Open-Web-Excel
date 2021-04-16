/*!
* Open Web Excel - ✍️ An Excel Used on the Browser Side.
* git+https://github.com/hai2007/Open-Web-Excel.git
*
* author 你好2007
*
* version 0.1.0-alpha.0
*
* Copyright (c) 2021 hai2007 走一步，再走一步。
* Released under the MIT license
*
* Date:Fri Apr 16 2021 14:19:00 GMT+0800 (GMT+08:00)
*/

"use strict";

function _typeof(obj) { "@babel/helpers - typeof"; if (typeof Symbol === "function" && typeof Symbol.iterator === "symbol") { _typeof = function _typeof(obj) { return typeof obj; }; } else { _typeof = function _typeof(obj) { return obj && typeof Symbol === "function" && obj.constructor === Symbol && obj !== Symbol.prototype ? "symbol" : typeof obj; }; } return _typeof(obj); }

(function () {
  'use strict';

  var owe = function owe(options) {};

  if ((typeof module === "undefined" ? "undefined" : _typeof(module)) === "object" && _typeof(module.exports) === "object") {
    module.exports = owe;
  } else {
    window.OpenWebExcel = owe;
  }
})();