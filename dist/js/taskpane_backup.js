// AI API 配置（默认值）
const DEFAULT_CONFIG = {
    apiUrl: 'https://api.deepseek.com/chat/completions',
    apiKey: 'sk-db6aab7933a2427497761018834fe5b1',
    model: 'deepseek-chat'
};

// 从PluginStorage加载配置
function loadConfig() {
    try {
        const saved = window.Application.PluginStorage.getItem('wpsagent_config');
        if (saved) {
            return { ...DEFAULT_CONFIG, ...JSON.parse(saved) };
        }
    } catch (e) {
        console.error('加载配置失败:', e);
    }
    return { ...DEFAULT_CONFIG };
}

// 保存配置到PluginStorage
function saveConfig(config) {
    try {
        window.Application.PluginStorage.setItem('wpsagent_config', JSON.stringify(config));
    } catch (e) {
        console.error('保存配置失败:', e);
    }
}

// 当前配置
let currentConfig = loadConfig();

// 对话历史
let conversationHistory = [
    {
        role: 'system',
        content: '你是WpsAgent，一个友好的AI助手，专门帮助用户处理文档相关的问题。请用简洁、专业的方式回答问题。'
    }
];

// 发送消息
async function sendMessage() {
    const input = document.getElementById('userInput');
    const sendBtn = document.getElementById('sendBtn');
    const message = input.value.trim();
    
    if (!message) return;
    
    // 禁用输入
    input.value = '';
    input.style.height = 'auto';
    sendBtn.disabled = true;
    
    // 添加用户消息到界面
    appendMessage('user', message);
    
    // 添加到对话历史
    conversationHistory.push({
        role: 'user',
        content: message
    });
    
    // 显示加载动画
    const loadingId = showLoading();
    
    try {
        const response = await callAIAPI(conversationHistory);
        
        // 移除加载动画
        removeLoading(loadingId);
        
        // 添加AI回复到界面
        appendMessage('assistant', response);
        
        // 添加到对话历史
        conversationHistory.push({
            role: 'assistant',
            content: response
        });
        
    } catch (error) {
        removeLoading(loadingId);
        appendMessage('assistant', '抱歉，发生了错误：' + error.message);
    }
    
    sendBtn.disabled = false;
    input.focus();
}

// 调用 AI API
async function callAIAPI(messages) {
    const response = await fetch(currentConfig.apiUrl, {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${currentConfig.apiKey}`
        },
        body: JSON.stringify({
            model: currentConfig.model,
            messages: messages,
            temperature: 0.7,
            max_tokens: 2048
        })
    });
    
    if (!response.ok) {
        const errorData = await response.json().catch(() => ({}));
        throw new Error(errorData.error?.message || `HTTP ${response.status}`);
    }
    
    const data = await response.json();
    return data.choices[0].message.content;
}

// 添加消息到聊天界面
function appendMessage(role, content) {
    const messagesContainer = document.getElementById('chatMessages');
    const messageDiv = document.createElement('div');
    messageDiv.className = `message ${role}`;
    
    const contentDiv = document.createElement('div');
    contentDiv.className = 'message-content';
    contentDiv.textContent = content;
    
    messageDiv.appendChild(contentDiv);
    messagesContainer.appendChild(messageDiv);
    
    // 滚动到底部
    messagesContainer.scrollTop = messagesContainer.scrollHeight;
}

// 显示加载动画
function showLoading() {
    const messagesContainer = document.getElementById('chatMessages');
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

// 移除加载动画
function removeLoading(loadingId) {
    const loadingDiv = document.getElementById(loadingId);
    if (loadingDiv) {
        loadingDiv.remove();
    }
}

// 自动调整输入框高度
function autoResizeTextarea() {
    const textarea = document.getElementById('userInput');
    textarea.style.height = 'auto';
    textarea.style.height = Math.min(textarea.scrollHeight, 100) + 'px';
}

// 初始化
window.onload = function() {
    const input = document.getElementById('userInput');
    
    // 监听输入调整高度
    input.addEventListener('input', autoResizeTextarea);
    
    // 监听回车发送（Shift+Enter换行）
    input.addEventListener('keydown', function(e) {
        if (e.key === 'Enter' && !e.shiftKey) {
            e.preventDefault();
            sendMessage();
        }
    });
}

// 清除文档所有标记（批注）
function clearMarkers() {
    try {
        const doc = window.Application.ActiveDocument;
        if (!doc) {
            appendMessage('assistant', '当前没有打开任何文档，无法清除标记。');
            return;
        }
        
        const count = doc.Comments.Count;
        if (count === 0) {
            appendMessage('assistant', '文档中没有发现任何标记（批注）。');
            return;
        }
        
        // 删除所有批注
        doc.DeleteAllComments();
        appendMessage('assistant', `✅ 已成功清除文档中的所有标记（共 ${count} 条）。`);
    } catch (e) {
        appendMessage('assistant', '清除标记失败：' + e.message);
    }
}

// 清空对话
function clearChat() {
    // 重置对话历史，只保留系统提示词
    conversationHistory = [
        {
            role: 'system',
            content: '你是WpsAgent，一个友好的AI助手，专门帮助用户处理文档相关的问题。请用简洁、专业的方式回答问题。'
        }
    ];
    
    // 清空界面消息
    const messagesContainer = document.getElementById('chatMessages');
    messagesContainer.innerHTML = `
        <div class="message assistant">
            <div class="message-content">你好！我是 WpsAgent，有什么可以帮助你的吗？</div>
        </div>
    `;
    
    // 重置输入框
    const input = document.getElementById('userInput');
    input.value = '';
    input.style.height = 'auto';
    input.focus();
}

// 获取WPS文档内容（支持获取选中内容）
function getDocumentContent() {
    try {
        const app = window.Application;
        const doc = app.ActiveDocument;
        if (!doc) {
            return { success: false, error: '当前没有打开任何文档' };
        }

        const selection = app.Selection;
        // 如果有选中内容且不是光标闪烁状态（即选中长度 > 0）
        if (selection && selection.Start !== selection.End) {
            const range = selection.Range;
            return { 
                success: true, 
                content: range.Text, 
                name: doc.Name, 
                isSelection: true,
                start: selection.Start,
                end: selection.End
            };
        }

        // 否则返回全文
        const fullRange = doc.Content;
        return { 
            success: true, 
            content: fullRange.Text, 
            name: doc.Name, 
            isSelection: false,
            start: 0,
            end: fullRange.End
        };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

// 审阅文档 - 错别字检查
async function reviewDocument() {
    const result = getDocumentContent();
    
    if (!result.success) {
        appendMessage('assistant', '无法获取文档内容：' + result.error);
        return;
    }
    
    if (!result.content || result.content.trim() === '') {
        appendMessage('assistant', '文档内容为空，无法进行检查。');
        return;
    }
    
    // 显示开始检查的提示
    const scopeText = result.isSelection ? '选中内容' : `文档「${result.name}」`;
    appendMessage('assistant', `正在对${scopeText}进行错别字检查...`);
    
    // 显示加载动画
    const loadingId = showLoading();
    
    try {
        // 构建错别字检查的提示词，要求返回JSON格式
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
        
        const response = await callAIAPI(checkMessages);
        
        // 移除加载动画
        removeLoading(loadingId);
        
        // 解析JSON结果
        let checkResult;
        try {
            // 尝试提取JSON（处理可能的markdown代码块）
            let jsonStr = response;
            const jsonMatch = response.match(/```(?:json)?\s*([\s\S]*?)\s*```/);
            if (jsonMatch) {
                jsonStr = jsonMatch[1];
            }
            checkResult = JSON.parse(jsonStr);
        } catch (e) {
            // JSON解析失败，直接显示原始结果
            appendMessage('assistant', `📝 错别字检查结果：\n\n${response}`);
            return;
        }
        
        // 显示检查结果
        if (!checkResult.hasErrors || !checkResult.errors || checkResult.errors.length === 0) {
            appendMessage('assistant', `✅ 检查完成！\n\n${checkResult.summary || '未发现错别字，内容正确。'}`);
            return;
        }
        
        // 构建结果文本
        let resultText = `📝 发现 ${checkResult.errors.length} 处错别字：\n\n`;
        checkResult.errors.forEach((err, index) => {
            resultText += `${index + 1}. 「${err.wrong}」→「${err.correct}」\n`;
            resultText += `   原因：${err.reason || '疑似错别字'}\n\n`;
        });
        resultText += `正在添加批注到${result.isSelection ? '选中部分' : '文档'}...`;
        appendMessage('assistant', resultText);
        
        // 添加批注到文档
        const commentCount = addCommentsToDocument(checkResult.errors, result);
        
        appendMessage('assistant', `✅ 已成功添加 ${commentCount} 条批注。`);
        
    } catch (error) {
        removeLoading(loadingId);
        appendMessage('assistant', '检查过程中发生错误：' + error.message);
    }
}

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
                content: `请润色以下内容：\n\n${result.content}`
            }
        ];
        
        const response = await callAIAPI(polishMessages);
        
        // 移除加载动画
        removeLoading(loadingId);
        
        // 显示润色结果
        appendMessage('assistant', `✨ 润色完成！正在添加批注到${result.isSelection ? '选中部分' : '文档'}...`);
        
        // 在文档中添加润色批注
        const commentAdded = addPolishComment(result, response);
        
        if (commentAdded) {
            appendMessage('assistant', '✅ 已将润色结果以批注形式添加到原文对应位置。');
        } else {
            appendMessage('assistant', '⚠️ 添加批注失败，以下是润色后的内容：\n\n' + response);
        }
        
    } catch (error) {
        removeLoading(loadingId);
        appendMessage('assistant', '润色过程中发生错误：' + error.message);
    }
}

// 将润色结果添加为批注
function addPolishComment(docContext, polishedText) {
    try {
        const doc = window.Application.ActiveDocument;
        if (!doc) return false;
        
        // 创建Range指向被润色的内容范围
        const range = doc.Range(docContext.start, docContext.end);
        
        // 构建批注内容
        const commentText = `✨ 润色建议：\n\n${polishedText}`;
        
        // 添加批注
        doc.Comments.Add(range, commentText);
        
        return true;
    } catch (e) {
        console.error('添加润色批注失败:', e);
        return false;
    }
}

// 在文档中添加批注
function addCommentsToDocument(errors, docContext) {
    const doc = window.Application.ActiveDocument;
    if (!doc) return 0;
    
    let addedCount = 0;
    // 只在被审阅的内容范围内进行查找
    const searchText = docContext.content;
    const baseOffset = docContext.start;
    
    for (const err of errors) {
        try {
            // 在被审阅的文本中查找错误词的位置
            let searchStart = 0;
            let pos = searchText.indexOf(err.wrong, searchStart);
            
            while (pos !== -1) {
                // 计算全局 Range 位置
                const globalStart = baseOffset + pos;
                const range = doc.Range(globalStart, globalStart + err.wrong.length);
                
                // 构建格式化的批注内容
                const commentText = `「${err.wrong}」→「${err.correct}」\n原因：${err.reason || '疑似错别字'}`;
                
                doc.Comments.Add(range, commentText);
                addedCount++;
                
                // 继续查找下一个
                searchStart = pos + err.wrong.length;
                pos = searchText.indexOf(err.wrong, searchStart);
            }
        } catch (e) {
            console.error('添加批注失败:', err.wrong, e);
        }
    }
    
    return addedCount;
}
