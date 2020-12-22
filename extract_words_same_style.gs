function onOpen() {
    let ui = DocumentApp.getUi()
    ui.createMenu('Custom Menu').addItem('Con lo stesso stile', 'extractWords').addToUi()
}

function sameStyle(a, b) {
    for (let k of [
        DocumentApp.Attribute.UNDERLINE, DocumentApp.Attribute.STRIKETHROUGH, DocumentApp.Attribute.FOREGROUND_COLOR,
        DocumentApp.Attribute.BOLD, DocumentApp.Attribute.ITALIC, DocumentApp.Attribute.BACKGROUND_COLOR,
        DocumentApp.Attribute.FONT_SIZE, DocumentApp.Attribute.FONT_FAMILY
    ]) {
        if (!a.hasOwnProperty(k) || !b.hasOwnProperty(k))
            continue
        if (a[k] !== b[k])
            return false
    }
    return true
}

function extractWords() {
    // cursor === null if there is a selection
    let cursor = DocumentApp.getActiveDocument().getCursor()
    let el = cursor ? cursor.getElement() : DocumentApp.getActiveDocument().getSelection().getRangeElements()[0].getElement()
    let style = el.getAttributes()
    // heading is stored in PARAGRAPH or TEXT parent
    let heading = el.getType() == DocumentApp.ElementType.PARAGRAPH ? el.getHeading() : el.getParent().getHeading()
    Logger.log(style, heading)
    let b = DocumentApp.getActiveDocument().getBody()
    let searchResult = null
    let results = []
    while (searchResult = b.findElement(DocumentApp.ElementType.TEXT, searchResult)) {
        let el = searchResult.getElement()
        let currentHeading = el.getParent().getHeading()
        let text = el.getText()
        let curr = ''
        for (let i = 0; i < text.length; i++) {
            let s = el.getAttributes(i)
            let same = sameStyle(s, style) && (!heading || heading === currentHeading)
            if (same)
                curr += text[i]
            if (!same && curr.length) {
                results.push(curr)
                curr = ''
            }
        }
        if (curr.length)
            results.push(curr)
    }

    if (results.length === 0)
        return

    b.appendPageBreak()

    for (let r of results) {
        let p = b.appendParagraph(r)
        p.setHeading(heading)
        p.setAttributes(style)
    }
}
