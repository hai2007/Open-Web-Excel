import xhtml from '@hai2007/tool/xhtml';

export function insertUp() {

    let rowNodes = xhtml.find(this.__contentDom[this.__tableIndex], () => true, 'tr');

    let newRowNode = xhtml.after(rowNodes[rowNodes.length - 1], '<tr><th class="line-num" open-web-excel>' + (rowNodes.length) + '</th></tr>');

    rowNodes.push(newRowNode);
    this.__contentArray[this.__tableIndex].content.push([]);

    // 一列列插入
    let colNum = this.__contentArray[this.__tableIndex].content[0].length;
    for (let index = 0; index < colNum; index++) {

        // 寻找到第一个可以在其前面追加的位置
        // 如果没有，第一个一定总是合适的
        let row;
        for (row = this.__rowNum; row > 1; row--) {

            if (
                // 如果当前结点显示，肯定可以
                this.__contentArray[this.__tableIndex].content[row - 1][index].style.display != 'none'
                ||

                // 如果前一个显示，而且行为1，当前应该也可以
                (
                    this.__contentArray[this.__tableIndex].content[row - 2][index].style.display != 'none'
                    &&
                    this.__contentArray[this.__tableIndex].content[row - 2][index].rowspan == '1'
                )
            ) break;
        }


        // 然后一个个向下迁移
        let index_move;
        for (index_move = rowNodes.length - 1; index_move > row; index_move--) {

            let needMoveNode = xhtml.find(rowNodes[index_move - 1], 'th')[index + 1];

            // 修改行号
            needMoveNode.setAttribute('row', 1 - -needMoveNode.getAttribute('row'));

            // 结点
            xhtml.after(
                xhtml.find(rowNodes[index_move], 'th')[index]
                ,
                needMoveNode
            );

            // 数据
            this.__contentArray[this.__tableIndex].content[index_move - 1][index] = this.__contentArray[this.__tableIndex].content[index_move - 2][index];
        }

        // 然后对于空白的，进行补充

        let tempNewItemData = this.$$newItemData();

        // 数据
        this.__contentArray[this.__tableIndex].content[index_move - 1][index] = tempNewItemData;

        // 结点
        let newItemNode = xhtml.after(
            xhtml.find(rowNodes[index_move], 'th')[index]
            ,
            `<th row="${index_move}" col="${index + 1}" contenteditable="true" class="item" colspan="1" rowspan="1" style="${this.$$styleToString(tempNewItemData.style)}" open-web-excel></th>`
        );

        // 绑定事件
        xhtml.bind(newItemNode, 'click', event => {

            this.$$itemClickHandler(event);

        });

    }

    this.__rowNum += 1;

};

export function insertDown() {

};

export function insertLeft() {

};

export function insertRight() {

};
