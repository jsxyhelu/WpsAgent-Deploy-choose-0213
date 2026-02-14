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
        
        const pf = range.ParagraphFormat;
        pf.LineSpacingRule = 4;
        pf.LineSpacing = config.lineSpacing;
        pf.CharacterUnitFirstLineIndent = 2;
        
        const paragraphs = range.Paragraphs;
        debug('开始应用格式', { paragraphCount: paragraphs.Count });
        
        for (let i = 1; i <= paragraphs.Count; i++) {
            const para = paragraphs.Item(i);
            const paraRange = para.Range;
            const styleName = para.Style ? para.Style.NameLocal.toLowerCase() : '';
            
            debug(`段落 ${i} 样式`, { styleName, text: paraRange.Text.substring(0, 20) });
            
            if (styleName.includes('heading 1') || styleName.includes('标题 1')) {
                paraRange.Font.NameFarEast = config.h1Font;
                paraRange.Font.Size = config.h1FontSize;
                paraRange.Font.Bold = true;
                debug(`段落 ${i} 应用1级标题样式`, { font: config.h1Font, size: config.h1FontSize });
            } else if (styleName.includes('heading 2') || styleName.includes('标题 2')) {
                paraRange.Font.NameFarEast = config.h2Font;
                paraRange.Font.Size = config.h2FontSize;
                paraRange.Font.Bold = true;
                debug(`段落 ${i} 应用2级标题样式`, { font: config.h2Font, size: config.h2FontSize });
            } else if (styleName.includes('heading 3') || styleName.includes('标题 3')) {
                paraRange.Font.NameFarEast = config.h3Font;
                paraRange.Font.Size = config.h3FontSize;
                paraRange.Font.Bold = true;
                debug(`段落 ${i} 应用3级标题样式`, { font: config.h3Font, size: config.h3FontSize });
            } else {
                paraRange.Font.NameFarEast = config.titleFont;
                paraRange.Font.Size = config.titleFontSize;
                paraRange.Font.Bold = false;
                debug(`段落 ${i} 应用正文样式`, { font: config.titleFont, size: config.titleFontSize });
            }
        }
        
        info('应用格式成功', { config });
        return { success: true };
    } catch (e) {
        error('应用格式异常', { error: e.message });
        return { success: false, error: e.message };
    }
}