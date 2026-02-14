// 润色文档内容
async function polishDocument() {
    const result = getDocumentContent();
    
    if (!result.success) {
        appendMessage('assistant', '无法获取文档内容：' + result.error);
        return;
    }
    
    if (!result.content || result.content.trim() === '') {
        appendMessage('assistant', '文档内容为空，无法进行润色。');
        return;
    }
    
    // 显示开始润色的提示
    const scopeText = result.isSelection ? '选中内容' : `文档「${result.name}」`;
    appendMessage('assistant', `正在对${scopeText}进行润色处理...`);
    
    // 显示加载动画
    const loadingId = showLoading();
    
    try {
        // 如果是全文润色，按自然段分别处理
        if (!result.isSelection) {
            await polishByParagraphs(result, loadingId);
        } else {
            // 选中内容润色，整体处理
            await polishSingleContent(result, loadingId);
        }
        
    } catch (error) {
        removeLoading(loadingId);
        appendMessage('assistant', '润色过程中发生错误：' + error.message);
    }
}

// 整体润色单个内容
async function polishSingleContent(docContext, loadingId) {
    // 构建润色请求的提示词
    const polishMessages = [
        {
            role: 'system',
            content: `你是一个专业的中文文案润色助手。请对用户提供的文本进行润色，使其更加流畅、准确、专业。

润色要求：
1. 保持原文的核心意思和结构不变
2. 优化表达方式，使语句更加通顺自然
3. 纠正语法错误和不当用词
4. 提升文字的专业性和可读性
5. 不要添加原文中没有的内容
6. 不要改变原文的语气和风格定位

请直接输出润色后的文本，不要添加任何说明或注释。`
        },
        {
            role: 'user',
            content: `请润色以下内容：\n\n${docContext.content}`
        }
    ];
    
    const response = await callAIAPI(polishMessages);
    
    // 移除加载动画
    removeLoading(loadingId);
    
    // 显示润色结果
    appendMessage('assistant', `✨ 润色完成！正在添加批注到选中部分...`);
    
    // 在文档中添加润色批注
    const commentAdded = addPolishComment(docContext, response);
    
    if (commentAdded) {
        appendMessage('assistant', '✅ 已将润色结果以批注形式添加到原文对应位置。');
    } else {
        appendMessage('assistant', '⚠️ 添加批注失败，以下是润色后的内容：\n\n' + response);
    }
}

// 按自然段分别润色全文
async function polishByParagraphs(docContext, loadingId) {
    const doc = window.Application.ActiveDocument;
    if (!doc) {
        removeLoading(loadingId);
        appendMessage('assistant', '无法访问文档。');
        return;
    }
    
    // 按自然段分割内容（以换行符分割）
    const paragraphs = docContext.content.split(/\n+/).filter(p => p.trim() !== '');
    
    if (paragraphs.length === 0) {
        removeLoading(loadingId);
        appendMessage('assistant', '文档中没有可润色的段落。');
        return;
    }
    
    removeLoading(loadingId);
    appendMessage('assistant', `📝 检测到 ${paragraphs.length} 个自然段，将逐段进行润色...`);
    
    let successCount = 0;
    let currentOffset = docContext.start;
    
    for (let i = 0; i < paragraphs.length; i++) {
        const paragraph = paragraphs[i];
        const paragraphLength = paragraph.length;
        
        // 显示当前进度
        const progressLoadingId = showLoading();
        appendMessage('assistant', `正在润色第 ${i + 1}/${paragraphs.length} 段...`);
        
        try {
            // 为当前段落请求润色
            const polishMessages = [
                {
                    role: 'system',
                    content: `你是一个专业的中文文案润色助手。请对用户提供的文本进行润色，使其更加流畅、准确、专业。

润色要求：
1. 保持原文的核心意思和结构不变
2. 优化表达方式，使语句更加通顺自然
3. 纠正语法错误和不当用词
4. 提升文字的专业性和可读性
5. 不要添加原文中没有的内容
6. 不要改变原文的语气和风格定位

请直接输出润色后的文本，不要添加任何说明或注释。`
                },
                {
                    role: 'user',
                    content: `请润色以下内容：\n\n${paragraph}`
                }
            ];
            
            const polishedParagraph = await callAIAPI(polishMessages);
            
            // 为当前段落添加批注
            const paragraphContext = {
                content: paragraph,
                start: currentOffset,
                end: currentOffset + paragraphLength
            };
            
            const commentAdded = addPolishComment(paragraphContext, polishedParagraph);
            
            removeLoading(progressLoadingId);
            
            if (commentAdded) {
                successCount++;
            }
            
        } catch (error) {
            removeLoading(progressLoadingId);
            appendMessage('assistant', `⚠️ 第 ${i + 1} 段润色失败：${error.message}`);
        }
        
        // 更新偏移量（包括段落长度和换行符）
        // 在文档文本中找到下一个段落的起始位置
        const nextParagraphIndex = docContext.content.indexOf(paragraphs[i + 1], currentOffset + paragraphLength - docContext.start);
        if (nextParagraphIndex !== -1) {
            currentOffset = docContext.start + nextParagraphIndex;
        } else {
            currentOffset += paragraphLength + 1; // 加上换行符
        }
    }
    
    appendMessage('assistant', `✅ 全文润色完成！成功为 ${successCount}/${paragraphs.length} 个段落添加了润色批注。`);
}
