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
        
        const paragraphs = range.Paragraphs;
        
        for (let i = 1; i <= paragraphs.Count; i++) {
            const para = paragraphs.Item(i);
            const text = para.Range.Text.trim();
            
            if (!text) continue;
            
            const style = para.Style;
            const outlineLevel = para.OutlineLevel;
            
            if (outlineLevel === 1) {
                para.Range.Font.NameFarEast = config.h1Font;
                para.Range.Font.Size = config.h1FontSize;
                para.Range.Font.Bold = true;
                debug('设置1级标题', { text: text.substring(0, 20) });
            } else if (outlineLevel === 2) {
                para.Range.Font.NameFarEast = config.h2Font;
                para.Range.Font.Size = config.h2FontSize;
                para.Range.Font.Bold = true;
                debug('设置2级标题', { text: text.substring(0, 20) });
            } else if (outlineLevel === 3) {
                para.Range.Font.NameFarEast = config.h3Font;
                para.Range.Font.Size = config.h3FontSize;
                para.Range.Font.Bold = true;
                debug('设置3级标题', { text: text.substring(0, 20) });
            } else if (outlineLevel === 0) {
                para.Range.Font.NameFarEast = config.titleFont;
                para.Range.Font.Size = config.titleFontSize;
                para.Range.Font.Bold = true;
                debug('设置标题', { text: text.substring(0, 20) });
            } else {
                const pf = para.ParagraphFormat;
                pf.LineSpacingRule = 4;
                pf.LineSpacing = config.lineSpacing;
                pf.CharacterUnitFirstLineIndent = 2;
                
                para.Range.Font.NameFarEast = '宋体';
                para.Range.Font.Size = 12;
                para.Range.Font.Bold = false;
                
                if (text.includes('（') && text.includes('）')) {
                    const bracketStart = text.indexOf('（');
                    const bracketEnd = text.indexOf('）');
                    if (bracketStart !== -1 && bracketEnd !== -1) {
                        const bracketRange = doc.Range(para.Range.Start + bracketStart, para.Range.Start + bracketEnd + 1);
                        bracketRange.Font.NameFarEast = config.bracketFont;
                        bracketRange.Font.Size = config.bracketFontSize;
                        debug('设置括号内文字', { text: text.substring(bracketStart, Math.min(bracketEnd + 1, 20)) });
                    }
                }
            }
        }
        
        info('应用格式成功', { config });
        return { success: true };
    } catch (e) {
        error('应用格式异常', { error: e.message });
        return { success: false, error: e.message };
    }
}