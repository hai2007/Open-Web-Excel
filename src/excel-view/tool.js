export function formatContent(content) {

    return content;

};

export function calcColName(index) {
    let codes = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z'];
    let result = "";
    while (true) {
        let _index = index % 26;

        if (_index == 0) {
            result = 'A' + result;
        } else {
            result = codes[_index] + result;
        }

        index = Math.floor(index / 26);

        if (index == 0) break;

        index -= 1;
    }
    return result;
};
