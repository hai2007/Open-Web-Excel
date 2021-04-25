
// 计算出区域的必要信息
export function calcRegionInfo(target1, target2) {

    let region = {

        // 区域的边界信息
        info: {},

        // 区域范围内的所有结点，第一个结点一定是左上角的那个
        nodes: []
    };

    // 先计算出行边界

    let row1_min = target1.row;
    let row1_max = target1.row + target1.rowNum - 1;

    let row2_min = target2.row;
    let row2_max = target2.row + target2.rowNum - 1;

    let row_min = row1_min > row2_min ? row2_min : row1_min;
    let row_max = row1_max > row2_max ? row1_max : row2_max;

    // 再计算出列边界

    let col1_min = target1.col;
    let col1_max = target1.col + target1.colNum - 1;

    let col2_min = target2.col;
    let col2_max = target2.col + target2.colNum - 1;

    let col_min = col1_min > col2_min ? col2_min : col1_min;
    let col_max = col1_max > col2_max ? col1_max : col2_max;

    // 然后就可以标记区域的边界了

    region.info = {
        row: [row_min, row_max],
        col: [col_min, col_max]
    };

    // 最后我们需要计算出此区域里面所有的结点

    let trs = this.__contentDom[this.__tableIndex].getElementsByTagName('tr');
    for (let i = row_min; i <= row_max; i++) {
        let ths = trs[i].getElementsByTagName('th');
        for (let j = 1; j < ths.length; j++) {

            let colValue = ths[j].getAttribute('col');

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
};

// 在页面中标记当前选择的区域
export function showRegion() {

    for (let i = 0; i < this.__region.nodes.length; i++) {
        this.__region.nodes[i].style.background = '#e5e0e0';
    }

};

// 取消在页面中标记的区域效果
export function cancelRegion() {

    for (let i = 0; i < this.__region.nodes.length; i++) {
        this.__region.nodes[i].style.background = this.__contentArray[this.__tableIndex].content[this.__region.nodes[i].getAttribute('row') - 1][this.__region.nodes[i].getAttribute('col') - 1].style.background;
    }

};
