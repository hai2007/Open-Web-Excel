// 键盘总控

import getKeyString from '@hai2007/tool/getKeyString.js';
import xhtml from '@hai2007/tool/xhtml';

export default function () {

    if ('_keyLog' in this) {
        console.error('Keyboard has been initialized');
        return;
    } else {

        this._keyLog = {
            'shift': false
        };

        xhtml.bind(document.body, 'keydown', event => {
            var keyString = getKeyString(event);

            // 标记shift按下
            if (keyString == 'shift') this._keyLog.shift = true;
        });

        xhtml.bind(document.body, 'keyup', event => {
            var keyString = getKeyString(event);

            // 标记shift放开
            if (keyString == 'shift') this._keyLog.shift = false;
        });

    }

};
