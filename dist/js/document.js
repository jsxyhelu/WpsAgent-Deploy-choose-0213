function getDocumentContent() {
    try {
        const app = window.Application;
        const doc = app.ActiveDocument;
        
        if (!doc) {
            warn('获取文档内容失败：没有打开的文档');
            return { success: false, error: '当前没有打开任何文档' };
        }

        const selection = app.Selection;
        
        if (selection && selection.Start !== selection.End) {
            const range = selection.Range;
            debug('获取选中内容', { length: range.Text.length });
            return { 
                success: true, 
                content: range.Text, 
                name: doc.Name, 
                isSelection: true,
                start: selection.Start,
                end: selection.End
            };
        }

        const fullRange = doc.Content;
        debug('获取全文内容', { length: fullRange.Text.length });
        return { 
            success: true, 
            content: fullRange.Text, 
            name: doc.Name, 
            isSelection: false,
            start: 0,
            end: fullRange.End
        };
    } catch (e) {
        error('获取文档内容异常', { error: e.message });
        return { success: false, error: e.message };
    }
}

function addCommentToDocument(range, text) {
    try {
        const doc = window.Application.ActiveDocument;
        if (!doc) {
            warn('添加批注失败：没有打开的文档');
            return false;
        }
        
        doc.Comments.Add(range, text);
        debug('添加批注成功', { textLength: text.length });
        return true;
    } catch (e) {
        error('添加批注异常', { error: e.message });
        return false;
    }
}

function clearAllComments() {
    try {
        const doc = window.Application.ActiveDocument;
        if (!doc) {
            warn('清除批注失败：没有打开的文档');
            return { success: false, error: '当前没有打开任何文档' };
        }
        
        const count = doc.Comments.Count;
        doc.DeleteAllComments();
        info('清除所有批注成功', { count });
        return { success: true, count };
    } catch (e) {
        error('清除批注异常', { error: e.message });
        return { success: false, error: e.message };
    }
}

function insertTextToDocument(text) {
    try {
        const app = window.Application;
        const doc = app.ActiveDocument;
        
        if (!doc) {
            warn('插入文本失败：没有打开的文档');
            return { success: false, error: '当前没有打开任何文档' };
        }
        
        const selection = app.Selection;
        selection.TypeText(text);
        selection.TypeParagraph();
        
        info('插入文本成功', { length: text.length });
        return { success: true };
    } catch (e) {
        error('插入文本异常', { error: e.message });
        return { success: false, error: e.message };
    }
}

function applyFormatting(config, docContext) {
    try {
        const app = window.Application;
        const doc = app.ActiveDocument;
        
        if (!doc) {
            warn('应用格式失败：没有打开的文档');
            return { success: false, error: '当前没有打开任何文档' };
        }
        
        const setup = doc.PageSetup;
        setup.TopMargin = app.CentimetersToPoints(config.marginTop);
        setup.BottomMargin = app.CentimetersToPoints(config.marginBottom);
        setup.LeftMargin = app.CentimetersToPoints(config.marginLeft);
        setup.RightMargin = app.CentimetersToPoints(config.marginRight);
        
        let range;
        if (docContext.isSelection) {
            range = doc.Range(docContext.start, docContext.end);
        } else {
            range = doc.Content;
        }
        
        range.Font.NameFarEast = config.titleFont;
        range.Font.Size = config.titleFontSize;
        
        const pf = range.ParagraphFormat;
        pf.LineSpacingRule = 4;
        pf.LineSpacing = config.lineSpacing;
        pf.CharacterUnitFirstLineIndent = 2;
        
        info('应用格式成功', { config });
        return { success: true };
    } catch (e) {
        error('应用格式异常', { error: e.message });
        return { success: false, error: e.message };
    }
}