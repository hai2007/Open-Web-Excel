export function getTargetNode(event) {
    let _event = event || window.event;
    return _event.target || _event.srcElement;
};

export function removeNode(node) {
    let pNode = node.parentNode;
    if (pNode) {
        pNode.removeChild(node);
    }
};
