var calcIndex = function (keyStr) {
    var indexExec = /([A-Z]+)(\d+)/.exec(keyStr);

    var col = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"].indexOf(indexExec[1]);

    return [indexExec[2] - 1, col];
};

function xlsxToJson(binaryData, type) {
    //利用XLSX解析二进制文件为xlsx对象
    var xlslObj = XLSX.read(binaryData, { type: type || 'binary' });

    // 把excel整理变成数据
    var excelJson = {
        filename: "Open-Web-Excel",
        version: "v1",
        contents: []
    };

    for (var index = 0; index < xlslObj.SheetNames.length; index++) {

        var name = xlslObj.SheetNames[index];
        var value = xlslObj.Sheets[name];

        var content = [];

        for (var key in value) {

            if (/^([A-Z]+)(\d+)$/.test(key)) {

                var indexArray = calcIndex(key);
                content[indexArray[0]] = content[indexArray[0]] || [];
                content[indexArray[0]][indexArray[1]] = {
                    colspan: "1",
                    rowspan: "1",
                    value: value[key].v,
                    style: {
                        display: "table-cell",
                        color: 'black',
                        background: 'white',
                        'vertical-align': 'top',
                        'text-align': 'left',
                        'font-weight': "normal",// bold粗体
                        'font-style': 'normal',// italic斜体
                        'text-decoration': 'none'// line-through中划线 underline下划线
                    }
                };

            }
        }

        excelJson.contents.push({
            content,
            name
        });
    }

    return excelJson;
}
