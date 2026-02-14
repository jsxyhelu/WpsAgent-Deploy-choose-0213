let conversationHistory = [
    {
        role: 'system',
        content: '你是WpsAgent，一个友好的AI助手，专门帮助用户处理文档相关的问题。请用简洁、专业的方式回答问题。'
    }
];

async function sendMessage() {
    const input = document.getElementById('userInput');
    const sendBtn = document.getElementById('sendBtn');
    const message = input.value.trim();
    
    if (!message) return;
    
    input.value = '';
    input.style.height = 'auto';
    sendBtn.disabled = true;
    
    appendMessage('user', message);
    
    conversationHistory.push({
        role: 'user',
        content: message
    });
    
    const loadingId = showLoading();
    
    try {
        const currentConfig = getConfig();
        const response = await callAIWithRetry(conversationHistory, currentConfig);
        
        removeLoading(loadingId);
        
        appendMessage('assistant', response);
        
        conversationHistory.push({
            role: 'assistant',
            content: response
        });
        
    } catch (error) {
        removeLoading(loadingId);
        appendMessage('assistant', '抱歉，发生了错误：' + error.message);
        error('发送消息失败', { error: error.message });
    }
    
    sendBtn.disabled = false;
    input.focus();
}

function parseMarkdown(text) {
    if (!text) return '';
    
    let html = text;
    
    html = html.replace(/```(\w*)\n([\s\S]*?)```/g, '<pre><code class="language-$1">$2</code></pre>');
    html = html.replace(/`([^`]+)`/g, '<code>$1</code>');
    html = html.replace(/^### (.*$)/gim, '<h3>$1</h3>');
    html = html.replace(/^## (.*$)/gim, '<h2>$1</h2>');
    html = html.replace(/^# (.*$)/gim, '<h1>$1</h1>');
    html = html.replace(/\*\*([^\*]+)\*\*/g, '<strong>$1</strong>');
    html = html.replace(/__([^_]+)__/g, '<strong>$1</strong>');
    html = html.replace(/\*([^\*]+)\*/g, '<em>$1</em>');
    html = html.replace(/_([^_]+)_/g, '<em>$1</em>');
    html = html.replace(/^\- (.+)$/gim, '<li>$1</li>');
    html = html.replace(/(<li>.*<\/li>)/s, '<ul>$1</ul>');
    html = html.replace(/^\d+\.\s+(.+)$/gim, '<li>$1</li>');
    html = html.replace(/\[([^\]]+)\]\(([^\)]+)\)/g, '<a href="$2" target="_blank">$1</a>');
    html = html.replace(/\n\n/g, '</p><p>');
    html = html.replace(/\n/g, '<br>');
    
    if (!html.startsWith('<')) {
        html = '<p>' + html + '</p>';
    }
    
    return html;
}

function appendMessage(role, content, targetId = 'chatMessages') {
    const messagesContainer = document.getElementById(targetId);
    if (!messagesContainer) return;

    const messageDiv = document.createElement('div');
    messageDiv.className = `message ${role}`;
    
    if (targetId !== 'chatMessages') {
        messageDiv.style.marginBottom = '8px';
        messageDiv.style.maxWidth = '100%';
    }

    const contentDiv = document.createElement('div');
    contentDiv.className = 'message-content';
    contentDiv.innerHTML = parseMarkdown(content);
    
    messageDiv.appendChild(contentDiv);
    messagesContainer.appendChild(messageDiv);
    
    messagesContainer.scrollTop = messagesContainer.scrollHeight;
}

function showLoading(targetId = 'chatMessages') {
    const messagesContainer = document.getElementById(targetId);
    if (!messagesContainer) return null;

    const loadingDiv = document.createElement('div');
    const loadingId = 'loading-' + Date.now();
    loadingDiv.id = loadingId;
    loadingDiv.className = 'message assistant loading';
    loadingDiv.innerHTML = `
        <div class="message-content">
            <div class="loading-dot"></div>
            <div class="loading-dot"></div>
            <div class="loading-dot"></div>
        </div>
    `;
    messagesContainer.appendChild(loadingDiv);
    messagesContainer.scrollTop = messagesContainer.scrollHeight;
    return loadingId;
}

function removeLoading(loadingId) {
    const loadingDiv = document.getElementById(loadingId);
    if (loadingDiv) {
        loadingDiv.remove();
    }
}

function autoResizeTextarea() {
    const textarea = document.getElementById('userInput');
    textarea.style.height = 'auto';
    textarea.style.height = Math.min(textarea.scrollHeight, 100) + 'px';
}

function getActiveTargetId() {
    const activeTab = document.querySelector('.tab-item.active');
    if (!activeTab) return 'chatMessages';
    const tabId = activeTab.getAttribute('data-tab-id');
    const mapping = {
        'chat': 'chatMessages',
        'proofread': 'proofreadResult',
        'polish': 'polishResult',
        'format': 'formatResult',
        'continue': 'continueResult',
        'translate': 'translateResult',
        'summarize': 'summarizeResult',
        'navigate': 'navigateResult',
        'write': 'writeResult'
    };
    return mapping[tabId] || 'chatMessages';
}

function clearMarkers() {
    const result = clearAllComments();
    if (result.success) {
        alert(`✅ 已成功清除文档中的所有标记（共 ${result.count} 条）。`);
    } else {
        alert('清除标记失败：' + result.error);
    }
}

function clearTabResult() {
    const activeTab = document.querySelector('.tab-item.active');
    if (!activeTab) return;
    
    const tabId = activeTab.getAttribute('data-tab-id');
    
    if (tabId === 'chat') {
        clearChat();
        return;
    }
    
    const targetId = getActiveTargetId();
    const container = document.getElementById(targetId);
    if (container) {
        container.innerHTML = '';
    }
}

function clearChat() {
    conversationHistory = [
        {
            role: 'system',
            content: '你是WpsAgent，一个友好的AI助手，专门帮助用户处理文档相关的问题。请用简洁、专业的方式回答问题。'
        }
    ];
    
    const messagesContainer = document.getElementById('chatMessages');
    messagesContainer.innerHTML = `
        <div class="message assistant">
            <div class="message-content">你好！我是 WpsAgent，有什么可以帮助你的吗？</div>
        </div>
    `;
    
    const input = document.getElementById('userInput');
    input.value = '';
    input.style.height = 'auto';
    input.focus();
    
    info('清空对话历史');
}

async function reviewDocument() {
    const targetId = getActiveTargetId();
    const result = getDocumentContent();
    
    if (!result.success) {
        appendMessage('assistant', '无法获取文档内容：' + result.error, targetId);
        return;
    }
    
    if (!result.content || result.content.trim() === '') {
        appendMessage('assistant', '文档内容为空，无法进行检查。', targetId);
        return;
    }
    
    if (targetId !== 'chatMessages') {
        const container = document.getElementById(targetId);
        if (container) container.innerHTML = '';
    }
    
    const scopeText = result.isSelection ? '选中内容' : `文档「${result.name}」`;
    appendMessage('assistant', `正在对${scopeText}进行错别字检查...`, targetId);
    
    const loadingId = showLoading(targetId);
    
    try {
        const currentConfig = getConfig();
        const checkMessages = [
            {
                role: 'system',
                content: `你是一个专业的中文文档校对助手。请仔细检查用户提供的文本中的错别字、同音字误用、形近字误用等问题。

请严格按照以下JSON格式输出检查结果，不要输出其他内容：
{
  "hasErrors": true或false,
  "errors": [
    {
      "wrong": "错误的词",
      "correct": "正确的词",
      "reason": "错误原因说明（如：同音字误用、形近字混淆等）",
      "context": "包含错误的句子片段"
    }
  ],
  "summary": "总结说明"
}

注意：
1. 只检查错别字问题，不要修改语法或文风
2. wrong字段必须是文档中实际出现的原文
3. reason字段简要说明错误类型和原因
4. 如果没有错别字，errors数组为空，hasErrors为false`
            },
            {
                role: 'user',
                content: `请检查以下文档内容的错别字：\n\n${result.content}`
            }
        ];
        
        const response = await callAIWithRetry(checkMessages, currentConfig);
        
        removeLoading(loadingId);
        
        let checkResult;
        try {
            let jsonStr = response;
            const jsonMatch = response.match(/```(?:json)?\s*([\s\S]*?)\s*```/);
            if (jsonMatch) {
                jsonStr = jsonMatch[1];
            }
            checkResult = JSON.parse(jsonStr);
        } catch (e) {
            appendMessage('assistant', `📝 错别字检查结果：\n\n${response}`, targetId);
            return;
        }
        
        if (!checkResult.hasErrors || !checkResult.errors || checkResult.errors.length === 0) {
            appendMessage('assistant', `✅ 检查完成！\n\n${checkResult.summary || '未发现错别字，内容正确。'}`, targetId);
            return;
        }
        
        let resultText = `📝 发现 ${checkResult.errors.length} 处错别字：\n\n`;
        checkResult.errors.forEach((err, index) => {
            resultText += `${index + 1}. 「${err.wrong}」→「${err.correct}」\n`;
            resultText += `   原因：${err.reason || '疑似错别字'}\n\n`;
        });
        resultText += `正在添加批注到${result.isSelection ? '选中部分' : '文档'}...`;
        appendMessage('assistant', resultText, targetId);
        
        const commentCount = addCommentsToDocument(checkResult.errors, result);
        
        appendMessage('assistant', `✅ 已成功添加 ${commentCount} 条批注。`, targetId);
        
    } catch (error) {
        removeLoading(loadingId);
        appendMessage('assistant', '检查过程中发生错误：' + error.message, targetId);
        error('校对文档失败', { error: error.message });
    }
}

function addCommentsToDocument(errors, docContext) {
    const doc = window.Application.ActiveDocument;
    if (!doc) return 0;
    
    let addedCount = 0;
    const searchText = docContext.content;
    const baseOffset = docContext.start;
    
    const errorOccurrences = [];
    for (const err of errors) {
        let searchStart = 0;
        let pos = searchText.indexOf(err.wrong, searchStart);
        
        while (pos !== -1) {
            errorOccurrences.push({
                pos: pos,
                globalStart: baseOffset + pos,
                err: err
            });
            searchStart = pos + err.wrong.length;
            pos = searchText.indexOf(err.wrong, searchStart);
        }
    }
    
    errorOccurrences.sort((a, b) => b.pos - a.pos);
    
    for (const item of errorOccurrences) {
        try {
            const err = item.err;
            const globalStart = item.globalStart;
            const wrongLength = err.wrong.length;
            
            const wrongRange = doc.Range(globalStart, globalStart + wrongLength);
            
            wrongRange.HighlightColorIndex = 6;
            wrongRange.Font.StrikeThrough = true;
            
            const commentText = `「${err.wrong}」→「${err.correct}」\n原因：${err.reason || '疑似错别字'}`;
            doc.Comments.Add(wrongRange, commentText);
            
            wrongRange.InsertBefore(err.correct);
            
            const correctRange = doc.Range(globalStart, globalStart + err.correct.length);
            correctRange.HighlightColorIndex = 4;
            correctRange.Font.Bold = true;
            
            addedCount++;
        } catch (e) {
            error('应用审阅格式失败', { error: e.message, wrong: item.err.wrong });
        }
    }
    
    return addedCount;
}

async function polishDocument() {
    const targetId = getActiveTargetId();
    const result = getDocumentContent();
    
    if (!result.success) {
        appendMessage('assistant', '无法获取文档内容：' + result.error, targetId);
        return;
    }
    
    if (!result.content || result.content.trim() === '') {
        appendMessage('assistant', '文档内容为空，无法进行润色。', targetId);
        return;
    }
    
    if (targetId !== 'chatMessages') {
        const container = document.getElementById(targetId);
        if (container) container.innerHTML = '';
    }
    
    const scopeText = result.isSelection ? '选中内容' : `文档「${result.name}」`;
    appendMessage('assistant', `正在对${scopeText}进行润色处理...`, targetId);
    
    const loadingId = showLoading(targetId);
    
    try {
        if (!result.isSelection) {
            await polishByParagraphs(result, loadingId, targetId);
        } else {
            await polishSingleContent(result, loadingId, targetId);
        }
        
    } catch (error) {
        removeLoading(loadingId);
        appendMessage('assistant', '润色过程中发生错误：' + error.message, targetId);
        error('润色文档失败', { error: error.message });
    }
}

async function polishSingleContent(docContext, loadingId, targetId = 'chatMessages') {
    const currentConfig = getConfig();
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
    
    const response = await callAIWithRetry(polishMessages, currentConfig);
    
    removeLoading(loadingId);
    
    appendMessage('assistant', `✨ 润色完成！正在添加批注到选中部分...`, targetId);
    
    const commentAdded = addPolishComment(docContext, response);
    
    if (commentAdded) {
        appendMessage('assistant', '✅ 已将润色结果以批注形式添加到原文对应位置。', targetId);
    } else {
        appendMessage('assistant', '⚠️ 添加批注失败，以下是润色后的内容：\n\n' + response, targetId);
    }
}

async function polishByParagraphs(docContext, loadingId, targetId = 'chatMessages') {
    const doc = window.Application.ActiveDocument;
    if (!doc) {
        removeLoading(loadingId);
        appendMessage('assistant', '无法访问文档。', targetId);
        return;
    }
    
    const paragraphs = docContext.content.split('\n');
    
    if (paragraphs.length === 0) {
        removeLoading(loadingId);
        appendMessage('assistant', '文档中没有可润色的段落。', targetId);
        return;
    }
    
    removeLoading(loadingId);
    appendMessage('assistant', `📝 检测到 ${paragraphs.length} 个自然段，将逐段进行润色...`, targetId);
    
    let successCount = 0;
    let currentOffset = docContext.start;
    const currentConfig = getConfig();
    
    for (let i = 0; i < paragraphs.length; i++) {
        const paragraph = paragraphs[i];
        const paragraphLength = paragraph.length;
        
        if (paragraph.trim() === '') {
            currentOffset += paragraphLength + 1;
            continue;
        }
        
        const progressLoadingId = showLoading(targetId);
        appendMessage('assistant', `正在润色第 ${i + 1}/${paragraphs.length} 段...`, targetId);
        
        try {
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
            
            const polishedParagraph = await callAIWithRetry(polishMessages, currentConfig);
            
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
            appendMessage('assistant', `⚠️ 第 ${i + 1} 段润色失败：${error.message}`, targetId);
            error('润色段落失败', { index: i, error: error.message });
        }
        
        currentOffset += paragraphLength + 1;
    }
    
    appendMessage('assistant', `✅ 全文润色完成！成功为 ${successCount}/${paragraphs.length} 个段落添加了润色批注。`, targetId);
}

function addPolishComment(docContext, polishedText) {
    try {
        const doc = window.Application.ActiveDocument;
        if (!doc) return false;
        
        const range = doc.Range(docContext.start, docContext.end);
        
        const commentText = `✨ 润色建议：\n\n${polishedText}`;
        
        doc.Comments.Add(range, commentText);
        
        return true;
    } catch (e) {
        error('添加润色批注失败', { error: e.message });
        return false;
    }
}

async function formatDocument() {
    const targetId = getActiveTargetId();
    const currentConfig = getConfig();
    const result = getDocumentContent();
    
    if (!result.success) {
        appendMessage('assistant', '错误：' + result.error, targetId);
        return;
    }

    if (targetId !== 'chatMessages') {
        document.getElementById(targetId).innerHTML = '';
    }
    appendMessage('assistant', '正在按照配置的标准进行排版...', targetId);
    const loadingId = showLoading(targetId);

    try {
        const systemPrompt = `你是一个专业的文档排版助手。请参考以下排版要求，对提供的文档内容进行规范化建议：

排版标准要求：
1. 标题字体：${currentConfig.titleFont}，字号：${currentConfig.titleFontSize}pt
2. 1级标题：${currentConfig.h1Font}，字号：${currentConfig.h1FontSize}pt
3. 2级标题：${currentConfig.h2Font}，字号：${currentConfig.h2FontSize}pt
4. 3级标题：${currentConfig.h3Font}，字号：${currentConfig.h3FontSize}pt
5. 行距设置：${currentConfig.lineSpacing}磅
6. 页边距配置：上${currentConfig.marginTop}cm, 下${currentConfig.marginBottom}cm, 左${currentConfig.marginLeft}cm, 右${currentConfig.marginRight}cm
7. 括号内文字建议字体：${currentConfig.bracketFont}，字号：${currentConfig.bracketFontSize}pt

请返回格式化后的文档内容建议，保持原文意思不变。`;

        const messages = [
            { role: 'system', content: systemPrompt },
            { role: 'user', content: result.content }
        ];

        const response = await callAIWithRetry(messages, currentConfig);
        removeLoading(loadingId);

        appendMessage('assistant', '📝 排版结果建议：\n\n' + response, targetId);

        try {
            const formatResult = applyFormatting(currentConfig, result);
            
            if (formatResult.success) {
                appendMessage('assistant', `✅ 已自动应用您的自定义排版配置：
- 页边距：${currentConfig.marginTop}/${currentConfig.marginBottom}/${currentConfig.marginLeft}/${currentConfig.marginRight} cm
- 行距：${currentConfig.lineSpacing} 磅
- 标题字体：${currentConfig.titleFont} (${currentConfig.titleFontSize}pt)
- 1级标题：${currentConfig.h1Font} (${currentConfig.h1FontSize}pt)
- 2级标题：${currentConfig.h2Font} (${currentConfig.h2FontSize}pt)
- 3级标题：${currentConfig.h3Font} (${currentConfig.h3FontSize}pt)
- 括号内容：${currentConfig.bracketFont} (${currentConfig.bracketFontSize}pt)`, targetId);
            } else {
                appendMessage('assistant', '⚠️ 自动格式化应用时遇到问题：' + formatResult.error + '\n您可以根据上面的 AI 建议手动调整。', targetId);
            }
        } catch (formatError) {
            error('排版应用失败', { error: formatError.message });
            appendMessage('assistant', '⚠️ 自动格式化应用时遇到问题：' + formatError.message + '\n您可以根据上面的 AI 建议手动调整。', targetId);
        }
    } catch (e) {
        removeLoading(loadingId);
        appendMessage('assistant', '排版失败：' + e.message, targetId);
        error('排版文档失败', { error: e.message });
    }
}

async function continueDocument() {
    const targetId = getActiveTargetId();
    const result = getDocumentContent();
    
    if (!result.success) { 
        appendMessage('assistant', '错误：' + result.error, targetId); 
        return; 
    }
    
    if (targetId !== 'chatMessages') document.getElementById(targetId).innerHTML = '';
    appendMessage('assistant', '正在构思续写内容...', targetId);
    const loadingId = showLoading(targetId);
    
    try {
        const currentConfig = getConfig();
        const messages = [
            { role: 'system', content: '你是一个擅长续写的作家。请根据用户提供的上文，续写一段文字（约100-200字），保持文风一致。直接输出续写内容。' },
            { role: 'user', content: result.content }
        ];
        const response = await callAIWithRetry(messages, currentConfig);
        removeLoading(loadingId);
        appendMessage('assistant', '✨ 续写建议：\n\n' + response, targetId);
        appendMessage('assistant', '您可以手动将上述内容复制到文档中。', targetId);
    } catch (e) {
        removeLoading(loadingId);
        appendMessage('assistant', '续写失败：' + e.message, targetId);
        error('续写文档失败', { error: e.message });
    }
}

async function translateDocument() {
    const targetId = getActiveTargetId();
    const result = getDocumentContent();
    
    if (!result.success) { 
        appendMessage('assistant', '错误：' + result.error, targetId); 
        return; 
    }
    
    if (targetId !== 'chatMessages') document.getElementById(targetId).innerHTML = '';
    appendMessage('assistant', '正在进行翻译...', targetId);
    const loadingId = showLoading(targetId);
    
    try {
        const currentConfig = getConfig();
        const messages = [
            { role: 'system', content: '你是一个精通中英互译的翻译官。请将用户提供的文本翻译成英文（如果原文是英文则翻译成中文）。保持语气地道。直接输出译文。' },
            { role: 'user', content: result.content }
        ];
        const response = await callAIWithRetry(messages, currentConfig);
        removeLoading(loadingId);
        appendMessage('assistant', '🌍 翻译结果：\n\n' + response, targetId);
    } catch (e) {
        removeLoading(loadingId);
        appendMessage('assistant', '翻译失败：' + e.message, targetId);
        error('翻译文档失败', { error: e.message });
    }
}

async function summarizeDocument() {
    const targetId = getActiveTargetId();
    const result = getDocumentContent();
    
    if (!result.success) { 
        appendMessage('assistant', '错误：' + result.error, targetId); 
        return; 
    }
    
    if (targetId !== 'chatMessages') document.getElementById(targetId).innerHTML = '';
    appendMessage('assistant', '正在提取摘要...', targetId);
    const loadingId = showLoading(targetId);
    
    try {
        const currentConfig = getConfig();
        const messages = [
            { role: 'system', content: '你是一个擅长提炼重点的助手。请总结用户提供的内容，列出核心要点（使用列表格式）。直接输出总结内容。' },
            { role: 'user', content: result.content }
        ];
        const response = await callAIWithRetry(messages, currentConfig);
        removeLoading(loadingId);
        appendMessage('assistant', '💡 内容摘要：\n\n' + response, targetId);
    } catch (e) {
        removeLoading(loadingId);
        appendMessage('assistant', '总结失败：' + e.message, targetId);
        error('总结文档失败', { error: e.message });
    }
}

const DOCUMENT_TYPES = {
    'notice': {
        name: '通知',
        prompt: '你是一位资深的党政机关公文写作专家。请根据GB/T 9704-2012《党政机关公文格式》标准，生成一份规范的通知文件。\n\n通知要求：\n1. 标题：关于XXX的通知\n2. 主送机关\n3. 正文：包括通知事由、通知事项、执行要求\n4. 落款：发文机关、成文日期\n\n请生成一份完整的通知范文（约300-500字），内容具体、条理清晰。'
    },
    'report': {
        name: '报告',
        prompt: '你是一位资深的党政机关公文写作专家。请根据GB/T 9704-2012《党政机关公文格式》标准，生成一份规范的报告文件。\n\n报告要求：\n1. 标题：关于XXX的报告\n2. 主送机关\n3. 正文：包括工作背景、主要内容、取得成效、存在问题、下步打算\n4. 结尾：特此报告\n5. 落款：报告机关、成文日期\n\n请生成一份完整的报告范文（约500-800字），内容详实、数据准确。'
    },
    'request': {
        name: '请示',
        prompt: '你是一位资深的党政机关公文写作专家。请根据GB/T 9704-2012《党政机关公文格式》标准，生成一份规范的请示文件。\n\n请示要求：\n1. 标题：关于XXX的请示\n2. 主送机关\n3. 正文：包括请示缘由、请示事项、请示理由\n4. 结尾：妥否，请批示或当否，请批复\n5. 落款：请示机关、成文日期\n\n请生成一份完整的请示范文（约300-500字），理由充分、请求明确。'
    },
    'reply': {
        name: '批复',
        prompt: '你是一位资深的党政机关公文写作专家。请根据GB/T 9704-2012《党政机关公文格式》标准，生成一份规范的批复文件。\n\n批复要求：\n1. 标题：关于XXX的批复\n2. 主送机关\n3. 正文：包括引述来文、明确意见、提出要求\n4. 结尾：此复\n5. 落款：批复机关、成文日期\n\n请生成一份完整的批复范文（约200-400字），态度明确、简洁规范。'
    },
    'letter': {
        name: '函',
        prompt: '你是一位资深的党政机关公文写作专家。请根据GB/T 9704-2012《党政机关公文格式》标准，生成一份规范的函件。\n\n函要求：\n1. 标题：关于XXX的函\n2. 主送机关\n3. 正文：包括发函缘由、发函事项、结尾用语（如"请予支持"、"盼复"等）\n4. 落款：发文机关、成文日期\n\n请生成一份完整的函件范文（约300-500字），语气平和、内容具体。'
    }
};

let currentGeneratedContent = '';

async function generateDocument(type) {
    const targetId = getActiveTargetId();
    const docConfig = DOCUMENT_TYPES[type];
    
    if (!docConfig) {
        appendMessage('assistant', '错误：不支持的公文类型', targetId);
        return;
    }

    if (targetId !== 'chatMessages') {
        document.getElementById(targetId).innerHTML = '';
    }
    
    appendMessage('assistant', `正在生成${docConfig.name}...`, targetId);
    const loadingId = showLoading(targetId);

    try {
        const currentConfig = getConfig();
        const messages = [
            { role: 'system', content: docConfig.prompt },
            { role: 'user', content: '请生成一份规范的' + docConfig.name + '范文' }
        ];

        const response = await callAIWithRetry(messages, currentConfig);
        removeLoading(loadingId);
        
        currentGeneratedContent = response;
        
        appendMessage('assistant', `📄 **${docConfig.name}范文**\n\n${response}`, targetId);
        
        const resultContainer = document.getElementById(targetId);
        const buttonDiv = document.createElement('div');
        buttonDiv.style.marginTop = '12px';
        buttonDiv.style.textAlign = 'center';
        buttonDiv.innerHTML = `
            <button onclick="insertToDocument()" style="padding: 10px 20px; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; border: none; border-radius: 6px; cursor: pointer; font-weight: 500; font-size: 14px; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">
                ✨ 添加到正文
            </button>
        `;
        resultContainer.appendChild(buttonDiv);
        
    } catch (e) {
        removeLoading(loadingId);
        appendMessage('assistant', `生成${docConfig.name}失败：` + e.message, targetId);
        error('生成公文失败', { type, error: e.message });
    }
}

function insertToDocument() {
    try {
        const app = window.Application;
        const doc = app.ActiveDocument;
        
        if (!doc) {
            alert('❌ 当前没有打开任何文档');
            return;
        }
        
        if (!currentGeneratedContent) {
            alert('❌ 没有可插入的内容');
            return;
        }
        
        const selection = app.Selection;
        selection.TypeText(currentGeneratedContent);
        selection.TypeParagraph();
        
        alert('✅ 内容已成功添加到文档中！');
        
        currentGeneratedContent = '';
        
    } catch (e) {
        error('插入文档失败', { error: e.message });
        alert('❌ 插入失败：' + e.message);
    }
}

window.onload = function() {
    const input = document.getElementById('userInput');
    
    input.addEventListener('input', autoResizeTextarea);
    
    input.addEventListener('keydown', function(e) {
        if (e.key === 'Enter' && !e.shiftKey) {
            e.preventDefault();
            sendMessage();
        }
    });

    window.tabManager = tabManager;

    setInterval(() => {
        try {
            const pendingTabStr = window.Application.PluginStorage.getItem('pending_tab');
            if (pendingTabStr) {
                const pendingTab = JSON.parse(pendingTabStr);
                if (pendingTab && pendingTab.id && pendingTab.label) {
                    tabManager.createOrActivateTab(pendingTab.id, pendingTab.label);
                    window.Application.PluginStorage.setItem('pending_tab', '');
                }
            }
        } catch (e) {
            debug('检查待创建 Tab 失败', { error: e.message });
        }
    }, 500);

    loadConfigFromServer().then(loadedConfig => {
        info('配置加载成功', { apiUrl: loadedConfig.apiUrl, model: loadedConfig.model });
    }).catch(err => {
        warn('从服务器加载配置失败，使用本地配置', { error: err.message });
    });
};