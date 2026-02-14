// AI API 及 文档格式配置（默认值）
const DEFAULT_CONFIG = {
    apiUrl: 'https://api.deepseek.com/chat/completions',
    apiKey: 'sk-db6aab7933a2427497761018834fe5b1',
    model: 'deepseek-chat',
    // 默认排版设置
    marginTop: 2.54,
    marginBottom: 2.54,
    marginLeft: 3.18,
    marginRight: 3.18,
    lineSpacing: 28, // 1.5倍行距对应磅值
    titleFont: '黑体',
    titleFontSize: 16,
    h1Font: '黑体',
    h1FontSize: 16,
    h2Font: '黑体',
    h2FontSize: 14,
    h3Font: '楷体',
    h3FontSize: 12,
    bracketFont: '仿宋',
    bracketFontSize: 12
};

// Tab 管理
const tabManager = {
    tabs: new Map(), // 存储已创建的 tab，key: tabId, value: { id, label, closeable }
    
    // 创建或激活 tab
    createOrActivateTab(tabId, label, closeable = true) {
        // 检查 tab 是否已存在
        if (this.tabs.has(tabId)) {
            this.activateTab(tabId);
            return;
        }
        
        // 创建新 tab
        this.tabs.set(tabId, { id: tabId, label, closeable });
        
        // 创建 tab 项
        const tabItem = this.createTabItem(tabId, label, closeable);
        document.getElementById('tabList').appendChild(tabItem);
        
        // 创建 tab 内容容器
        const tabContent = this.createTabContent(tabId);
        document.querySelector('.tab-contents').appendChild(tabContent);
        
        // 激活新 tab
        this.activateTab(tabId);
    },
    
    // 创建 tab 项 DOM
    createTabItem(tabId, label, closeable) {
        const tabItem = document.createElement('div');
        tabItem.className = 'tab-item';
        tabItem.setAttribute('data-tab-id', tabId);
        tabItem.onclick = () => this.activateTab(tabId);
        
        const tabLabel = document.createElement('span');
        tabLabel.className = 'tab-label';
        tabLabel.textContent = label;
        tabItem.appendChild(tabLabel);
        
        // 添加关闭按钮（如果可关闭）
        if (closeable) {
            const closeBtn = document.createElement('span');
            closeBtn.className = 'tab-close';
            closeBtn.innerHTML = `
                <svg viewBox="0 0 12 12" fill="none" stroke="currentColor" stroke-width="2">
                    <path d="M2 2L10 10M10 2L2 10"/>
                </svg>
            `;
            closeBtn.onclick = (e) => {
                e.stopPropagation();
                this.closeTab(tabId);
            };
            tabItem.appendChild(closeBtn);
        }
        
        return tabItem;
    },
    
    // 创建 tab 内容容器
    createTabContent(tabId) {
        const tabContent = document.createElement('div');
        tabContent.className = 'tab-content';
        tabContent.setAttribute('data-tab-id', tabId);
        
        // 根据 tabId 创建不同的内容
        const contentHtml = this.getTabContentTemplate(tabId);
        tabContent.innerHTML = contentHtml;
        
        return tabContent;
    },
    
    // 获取 tab 内容模板
    getTabContentTemplate(tabId) {
        const templates = {
            'proofread': `
                <div class="tab-functional-content" style="padding: 12px; display: flex; flex-direction: column; gap: 12px;">
                    <div style="background: #f8f9ff; padding: 10px; border-radius: 8px; border: 1px solid #e0e7ff;">
                        <div style="display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 4px;">
                            <h4 style="margin: 0; color: #4338ca;">错别字校对</h4>
                            <div style="display: flex; gap: 6px;">
                                <button onclick="clearMarkers()" style="padding: 3px 8px; background: white; border: 1px solid #d1d5db; border-radius: 4px; cursor: pointer; font-size: 11px; color: #666;">完成校对</button>
                                <button onclick="clearTabResult()" style="padding: 3px 8px; background: white; border: 1px solid #d1d5db; border-radius: 4px; cursor: pointer; font-size: 11px; color: #666;">历史清空</button>
                            </div>
                        </div>
                        <p style="font-size: 12px; color: #666; margin: 0;">AI 将自动扫描文档中的错别字、词语误用，并以批注形式给出修改建议。</p>
                    </div>
                    <button class="primary-btn" onclick="reviewDocument()" style="padding: 10px; background: #667eea; color: white; border: none; border-radius: 6px; cursor: pointer; font-weight: 500;">
                        开始检查
                    </button>
                    <div id="proofreadResult" style="margin-top: 8px; font-size: 13px; line-height: 1.6; color: #444;"></div>
                </div>
            `,
            'polish': `
                <div class="tab-functional-content" style="padding: 12px; display: flex; flex-direction: column; gap: 12px;">
                    <div style="background: #f8f9ff; padding: 10px; border-radius: 8px; border: 1px solid #e0e7ff;">
                        <div style="display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 4px;">
                            <h4 style="margin: 0; color: #4338ca;">文档润色</h4>
                            <div style="display: flex; gap: 6px;">
                                <button onclick="clearMarkers()" style="padding: 3px 8px; background: white; border: 1px solid #d1d5db; border-radius: 4px; cursor: pointer; font-size: 11px; color: #666;">完成润色</button>
                                <button onclick="clearTabResult()" style="padding: 3px 8px; background: white; border: 1px solid #d1d5db; border-radius: 4px; cursor: pointer; font-size: 11px; color: #666;">历史清空</button>
                            </div>
                        </div>
                        <p style="font-size: 12px; color: #666; margin: 0;">提升文章的流畅度、专业感和可读性。支持全文分段处理或对选中内容进行润色。</p>
                    </div>
                    <button class="primary-btn" onclick="polishDocument()" style="padding: 10px; background: #667eea; color: white; border: none; border-radius: 6px; cursor: pointer; font-weight: 500;">
                        立即润色
                    </button>
                    <div id="polishResult" style="margin-top: 8px; font-size: 13px; line-height: 1.6; color: #444;"></div>
                </div>
            `,
            'format': `
                <div class="tab-functional-content" style="padding: 12px; display: flex; flex-direction: column; gap: 12px;">
                    <div style="background: #f8f9ff; padding: 10px; border-radius: 8px; border: 1px solid #e0e7ff;">
                        <div style="display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 4px;">
                            <h4 style="margin: 0; color: #4338ca;">智能排版</h4>
                            <div style="display: flex; gap: 6px;">
                                <button onclick="clearMarkers()" style="padding: 3px 8px; background: white; border: 1px solid #d1d5db; border-radius: 4px; cursor: pointer; font-size: 11px; color: #666;">完成排版</button>
                                <button onclick="clearTabResult()" style="padding: 3px 8px; background: white; border: 1px solid #d1d5db; border-radius: 4px; cursor: pointer; font-size: 11px; color: #666;">历史清空</button>
                            </div>
                        </div>
                        <p style="font-size: 12px; color: #666; margin: 0;">AI 将分析文档结构，针对字体、间距、标题层级等给出专业的排版建议。</p>
                    </div>
                    <button class="primary-btn" onclick="formatDocument()" style="padding: 10px; background: #667eea; color: white; border: none; border-radius: 6px; cursor: pointer; font-weight: 500;">
                        分析排版
                    </button>
                    <div id="formatResult" style="margin-top: 8px; font-size: 13px; line-height: 1.6; color: #444;"></div>
                </div>
            `,
            'continue': `
                <div class="tab-functional-content" style="padding: 12px; display: flex; flex-direction: column; gap: 12px;">
                    <div style="background: #f8f9ff; padding: 10px; border-radius: 8px; border: 1px solid #e0e7ff;">
                        <div style="display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 4px;">
                            <h4 style="margin: 0; color: #4338ca;">内容续写</h4>
                            <div style="display: flex; gap: 6px;">
                                <button onclick="clearMarkers()" style="padding: 3px 8px; background: white; border: 1px solid #d1d5db; border-radius: 4px; cursor: pointer; font-size: 11px; color: #666;">完成续写</button>
                                <button onclick="clearTabResult()" style="padding: 3px 8px; background: white; border: 1px solid #d1d5db; border-radius: 4px; cursor: pointer; font-size: 11px; color: #666;">历史清空</button>
                            </div>
                        </div>
                        <p style="font-size: 12px; color: #666; margin: 0;">根据上文语境自动续写内容，保持文风一致，激发创作灵感。</p>
                    </div>
                    <button class="primary-btn" onclick="continueDocument()" style="padding: 10px; background: #667eea; color: white; border: none; border-radius: 6px; cursor: pointer; font-weight: 500;">
                        开始续写
                    </button>
                    <div id="continueResult" style="margin-top: 8px; font-size: 13px; line-height: 1.6; color: #444;"></div>
                </div>
            `,
            'translate': `
                <div class="tab-functional-content" style="padding: 12px; display: flex; flex-direction: column; gap: 12px;">
                    <div style="background: #f8f9ff; padding: 10px; border-radius: 8px; border: 1px solid #e0e7ff;">
                        <div style="display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 4px;">
                            <h4 style="margin: 0; color: #4338ca;">翻译助手</h4>
                            <div style="display: flex; gap: 6px;">
                                <button onclick="clearMarkers()" style="padding: 3px 8px; background: white; border: 1px solid #d1d5db; border-radius: 4px; cursor: pointer; font-size: 11px; color: #666;">完成翻译</button>
                                <button onclick="clearTabResult()" style="padding: 3px 8px; background: white; border: 1px solid #d1d5db; border-radius: 4px; cursor: pointer; font-size: 11px; color: #666;">历史清空</button>
                            </div>
                        </div>
                        <p style="font-size: 12px; color: #666; margin: 0;">支持多语言互译，保持原意准确性的同时优化译文地道度。</p>
                    </div>
                    <button class="primary-btn" onclick="translateDocument()" style="padding: 10px; background: #667eea; color: white; border: none; border-radius: 6px; cursor: pointer; font-weight: 500;">
                        翻译为英文
                    </button>
                    <div id="translateResult" style="margin-top: 8px; font-size: 13px; line-height: 1.6; color: #444;"></div>
                </div>
            `,
            'summarize': `
                <div class="tab-functional-content" style="padding: 12px; display: flex; flex-direction: column; gap: 12px;">
                    <div style="background: #f8f9ff; padding: 10px; border-radius: 8px; border: 1px solid #e0e7ff;">
                        <div style="display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 4px;">
                            <h4 style="margin: 0; color: #4338ca;">全文总结</h4>
                            <div style="display: flex; gap: 6px;">
                                <button onclick="clearMarkers()" style="padding: 3px 8px; background: white; border: 1px solid #d1d5db; border-radius: 4px; cursor: pointer; font-size: 11px; color: #666;">完成总结</button>
                                <button onclick="clearTabResult()" style="padding: 3px 8px; background: white; border: 1px solid #d1d5db; border-radius: 4px; cursor: pointer; font-size: 11px; color: #666;">历史清空</button>
                            </div>
                        </div>
                        <p style="font-size: 12px; color: #666; margin: 0;">自动提取文档核心观点和摘要，帮助快速把握长文重点。</p>
                    </div>
                    <button class="primary-btn" onclick="summarizeDocument()" style="padding: 10px; background: #667eea; color: white; border: none; border-radius: 6px; cursor: pointer; font-weight: 500;">
                        提取摘要
                    </button>
                    <div id="summarizeResult" style="margin-top: 8px; font-size: 13px; line-height: 1.6; color: #444;"></div>
                </div>
            `,
            'navigate': `
                <div class="tab-functional-content" style="height: 100%; display: flex; flex-direction: column; padding: 0;">
                    <div id="navigateResult" style="flex: 1; width: 100%; height: 100%; overflow: hidden;">
                        <iframe id="navigationFrame" src="${NAVIGATION_URLS.DOUBAO}" style="width: 100%; height: 100%; border: none;" sandbox="allow-same-origin allow-scripts allow-popups allow-forms"></iframe>
                    </div>
                </div>
            `,
            'write': `
                <div class="tab-functional-content" style="padding: 12px; display: flex; flex-direction: column; gap: 12px;">
                    <div style="background: #f8f9ff; padding: 10px; border-radius: 8px; border: 1px solid #e0e7ff;">
                        <div style="display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 4px;">
                            <h4 style="margin: 0; color: #4338ca;">机关公文写作助手</h4>
                            <div style="display: flex; gap: 6px;">
                                <button onclick="clearMarkers()" style="padding: 3px 8px; background: white; border: 1px solid #d1d5db; border-radius: 4px; cursor: pointer; font-size: 11px; color: #666;">完成写作</button>
                                <button onclick="clearTabResult()" style="padding: 3px 8px; background: white; border: 1px solid #d1d5db; border-radius: 4px; cursor: pointer; font-size: 11px; color: #666;">历史清空</button>
                            </div>
                        </div>
                        <p style="font-size: 12px; color: #666; margin: 0;">选择常见公文类型，AI将为您生成规范化文本，可直接插入到文档中。</p>
                    </div>
                    <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 8px;">
                        <button onclick="generateDocument('notice')" style="padding: 12px; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; border: none; border-radius: 6px; cursor: pointer; font-weight: 500; font-size: 13px;">
                            📢 通知
                        </button>
                        <button onclick="generateDocument('report')" style="padding: 12px; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; border: none; border-radius: 6px; cursor: pointer; font-weight: 500; font-size: 13px;">
                            📊 报告
                        </button>
                        <button onclick="generateDocument('request')" style="padding: 12px; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; border: none; border-radius: 6px; cursor: pointer; font-weight: 500; font-size: 13px;">
                            📝 请示
                        </button>
                        <button onclick="generateDocument('reply')" style="padding: 12px; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; border: none; border-radius: 6px; cursor: pointer; font-weight: 500; font-size: 13px;">
                            💬 批复
                        </button>
                        <button onclick="generateDocument('letter')" style="padding: 12px; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; border: none; border-radius: 6px; cursor: pointer; font-weight: 500; font-size: 13px; grid-column: 1 / -1;">
                            ✉️ 函
                        </button>
                    </div>
                    <div id="writeResult" style="margin-top: 8px; font-size: 13px; line-height: 1.6; color: #444;"></div>
                </div>
            `
        };
        
        return templates[tabId] || `<div style="padding: 12px;"><h3>${tabId}</h3><p>内容加载中...</p></div>`;
    },
    
    // 激活 tab
    activateTab(tabId) {
        // 移除所有 active 状态
        document.querySelectorAll('.tab-item').forEach(item => {
            item.classList.remove('active');
        });
        document.querySelectorAll('.tab-content').forEach(content => {
            content.classList.remove('active');
        });
        
        // 添加当前 tab 的 active 状态
        const tabItem = document.querySelector(`.tab-item[data-tab-id="${tabId}"]`);
        const tabContent = document.querySelector(`.tab-content[data-tab-id="${tabId}"]`);
        
        if (tabItem) tabItem.classList.add('active');
        if (tabContent) tabContent.classList.add('active');
    },
    
    // 关闭 tab
    closeTab(tabId) {
        // 不能关闭对话 tab
        if (tabId === 'chat') return;
        
        // 获取要关闭的 tab
        const tabItem = document.querySelector(`.tab-item[data-tab-id="${tabId}"]`);
        const tabContent = document.querySelector(`.tab-content[data-tab-id="${tabId}"]`);
        
        if (!tabItem) return;
        
        const wasActive = tabItem.classList.contains('active');
        
        // 删除 DOM 元素
        tabItem.remove();
        if (tabContent) tabContent.remove();
        
        // 从 Map 中删除
        this.tabs.delete(tabId);
        
        // 如果关闭的是当前激活的 tab，激活对话 tab
        if (wasActive) {
            this.activateTab('chat');
        }
    }
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

// 导航功能相关的URL配置
const NAVIGATION_URLS = {
    DOUBAO: 'https://www.doubao.com/'
};



// 简易 Markdown 转 HTML 解析器
function parseMarkdown(text) {
    if (!text) return '';
    
    let html = text;
    
    // 代码块（需要先处理，避免内部内容被错误解析）
    html = html.replace(/```(\w*)\n([\s\S]*?)```/g, '<pre><code class="language-$1">$2</code></pre>');
    
    // 行内代码
    html = html.replace(/`([^`]+)`/g, '<code>$1</code>');
    
    // 标题（h1-h6）
    html = html.replace(/^### (.*$)/gim, '<h3>$1</h3>');
    html = html.replace(/^## (.*$)/gim, '<h2>$1</h2>');
    html = html.replace(/^# (.*$)/gim, '<h1>$1</h1>');
    
    // 粗体
    html = html.replace(/\*\*([^\*]+)\*\*/g, '<strong>$1</strong>');
    html = html.replace(/__([^_]+)__/g, '<strong>$1</strong>');
    
    // 斜体
    html = html.replace(/\*([^\*]+)\*/g, '<em>$1</em>');
    html = html.replace(/_([^_]+)_/g, '<em>$1</em>');
    
    // 无序列表
    html = html.replace(/^\- (.+)$/gim, '<li>$1</li>');
    html = html.replace(/(<li>.*<\/li>)/s, '<ul>$1</ul>');
    
    // 有序列表
    html = html.replace(/^\d+\.\s+(.+)$/gim, '<li>$1</li>');
    
    // 链接
    html = html.replace(/\[([^\]]+)\]\(([^\)]+)\)/g, '<a href="$2" target="_blank">$1</a>');
    
    // 换行处理
    html = html.replace(/\n\n/g, '</p><p>');
    html = html.replace(/\n/g, '<br>');
    
    // 包裹段落
    if (!html.startsWith('<')) {
        html = '<p>' + html + '</p>';
    }
    
    return html;
}

// 添加消息到界面
function appendMessage(role, content, targetId = 'chatMessages') {
    const messagesContainer = document.getElementById(targetId);
    if (!messagesContainer) return;

    const messageDiv = document.createElement('div');
    messageDiv.className = `message ${role}`;
    
    // 如果是功能 Tab，样式稍作调整
    if (targetId !== 'chatMessages') {
        messageDiv.style.marginBottom = '8px';
        messageDiv.style.maxWidth = '100%';
    }

    const contentDiv = document.createElement('div');
    contentDiv.className = 'message-content';
    
    // 解析 Markdown 并渲染为 HTML
    contentDiv.innerHTML = parseMarkdown(content);
    
    messageDiv.appendChild(contentDiv);
    messagesContainer.appendChild(messageDiv);
    
    // 滚动到底部
    messagesContainer.scrollTop = messagesContainer.scrollHeight;
}

// 显示加载动画
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

    // 暴露 tabManager 到全局
    window.tabManager = tabManager;

    // 定时检查是否有待创建的 tab（用于与 Ribbon 通信）
    setInterval(() => {
        try {
            const pendingTabStr = window.Application.PluginStorage.getItem('pending_tab');
            if (pendingTabStr) {
                const pendingTab = JSON.parse(pendingTabStr);
                // 确保数据有效
                if (pendingTab && pendingTab.id && pendingTab.label) {
                    tabManager.createOrActivateTab(pendingTab.id, pendingTab.label);
                    // 处理完后清除
                    window.Application.PluginStorage.setItem('pending_tab', '');
                }
            }
        } catch (e) {
            // 忽略错误
        }
    }, 500);
}

// 清除文档所有标记（批注）
function clearMarkers() {
    try {
        const doc = window.Application.ActiveDocument;
        if (!doc) {
            alert('当前没有打开任何文档，无法清除标记。');
            return;
        }
        
        const count = doc.Comments.Count;
        if (count === 0) {
            alert('文档中没有发现任何标记（批注）。');
            return;
        }
        
        // 删除所有批注
        doc.DeleteAllComments();
        alert(`✅ 已成功清除文档中的所有标记（共 ${count} 条）。`);
    } catch (e) {
        alert('清除标记失败：' + e.message);
    }
}

// 清空当前 Tab 的结果显示区
function clearTabResult() {
    const activeTab = document.querySelector('.tab-item.active');
    if (!activeTab) return;
    
    const tabId = activeTab.getAttribute('data-tab-id');
    
    // 对话 Tab 特殊处理
    if (tabId === 'chat') {
        clearChat();
        return;
    }
    
    // 功能 Tab：清空结果区
    const targetId = getActiveTargetId();
    const container = document.getElementById(targetId);
    if (container) {
        container.innerHTML = '';
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

// 获取当前激活 Tab 的消息容器 ID
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

// 审阅文档 - 错别字检查
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
    
    // 清空上次结果
    if (targetId !== 'chatMessages') {
        const container = document.getElementById(targetId);
        if (container) container.innerHTML = '';
    }
    
    // 显示开始检查的提示
    const scopeText = result.isSelection ? '选中内容' : `文档「${result.name}」`;
    appendMessage('assistant', `正在对${scopeText}进行错别字检查...`, targetId);
    
    // 显示加载动画
    const loadingId = showLoading(targetId);
    
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
            appendMessage('assistant', `📝 错别字检查结果：\n\n${response}`, targetId);
            return;
        }
        
        // 显示检查结果
        if (!checkResult.hasErrors || !checkResult.errors || checkResult.errors.length === 0) {
            appendMessage('assistant', `✅ 检查完成！\n\n${checkResult.summary || '未发现错别字，内容正确。'}`, targetId);
            return;
        }
        
        // 构建结果文本
        let resultText = `📝 发现 ${checkResult.errors.length} 处错别字：\n\n`;
        checkResult.errors.forEach((err, index) => {
            resultText += `${index + 1}. 「${err.wrong}」→「${err.correct}」\n`;
            resultText += `   原因：${err.reason || '疑似错别字'}\n\n`;
        });
        resultText += `正在添加批注到${result.isSelection ? '选中部分' : '文档'}...`;
        appendMessage('assistant', resultText, targetId);
        
        // 添加批注到文档
        const commentCount = addCommentsToDocument(checkResult.errors, result);
        
        appendMessage('assistant', `✅ 已成功添加 ${commentCount} 条批注。`, targetId);
        
    } catch (error) {
        removeLoading(loadingId);
        appendMessage('assistant', '检查过程中发生错误：' + error.message, targetId);
    }
}

// 润色文档内容
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
    
    // 清空上次结果
    if (targetId !== 'chatMessages') {
        const container = document.getElementById(targetId);
        if (container) container.innerHTML = '';
    }
    
    // 显示开始润色的提示
    const scopeText = result.isSelection ? '选中内容' : `文档「${result.name}」`;
    appendMessage('assistant', `正在对${scopeText}进行润色处理...`, targetId);
    
    // 显示加载动画
    const loadingId = showLoading(targetId);
    
    try {
        // 如果是全文润色，按自然段分别处理
        if (!result.isSelection) {
            await polishByParagraphs(result, loadingId, targetId);
        } else {
            // 选中内容润色，整体处理
            await polishSingleContent(result, loadingId, targetId);
        }
        
    } catch (error) {
        removeLoading(loadingId);
        appendMessage('assistant', '润色过程中发生错误：' + error.message, targetId);
    }
}

// 整体润色单个内容
async function polishSingleContent(docContext, loadingId, targetId = 'chatMessages') {
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
    appendMessage('assistant', `✨ 润色完成！正在添加批注到选中部分...`, targetId);
    
    // 在文档中添加润色批注
    const commentAdded = addPolishComment(docContext, response);
    
    if (commentAdded) {
        appendMessage('assistant', '✅ 已将润色结果以批注形式添加到原文对应位置。', targetId);
    } else {
        appendMessage('assistant', '⚠️ 添加批注失败，以下是润色后的内容：\n\n' + response, targetId);
    }
}

// 按自然段分别润色全文
async function polishByParagraphs(docContext, loadingId, targetId = 'chatMessages') {
    const doc = window.Application.ActiveDocument;
    if (!doc) {
        removeLoading(loadingId);
        appendMessage('assistant', '无法访问文档。', targetId);
        return;
    }
    
    // 按自然段分割内容（每个回车换行就是一个自然段）
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
    
    for (let i = 0; i < paragraphs.length; i++) {
        const paragraph = paragraphs[i];
        const paragraphLength = paragraph.length;
        
        // 跳过空段落（但不影响偏移量计算）
        if (paragraph.trim() === '') {
            currentOffset += paragraphLength + 1; // 包括换行符
            continue;
        }
        
        // 显示当前进度
        const progressLoadingId = showLoading(targetId);
        appendMessage('assistant', `正在润色第 ${i + 1}/${paragraphs.length} 段...`, targetId);
        
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
            appendMessage('assistant', `⚠️ 第 ${i + 1} 段润色失败：${error.message}`, targetId);
        }
        
        // 更新偏移量：当前段落长度 + 换行符(1个字符)
        currentOffset += paragraphLength + 1;
    }
    
    appendMessage('assistant', `✅ 全文润色完成！成功为 ${successCount}/${paragraphs.length} 个段落添加了润色批注。`, targetId);
}
// 智能排版
// 按照国家标准对文档进行排版
async function formatDocument() {
    const targetId = getActiveTargetId();
    const config = loadConfig(); // 加载用户自定义配置
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
        // 构建系统提示词，参考用户配置
        const systemPrompt = `你是一个专业的文档排版助手。请参考以下排版要求，对提供的文档内容进行规范化建议：

排版标准要求：
1. 标题字体：${config.titleFont}，字号：${config.titleFontSize}pt
2. 1级标题：${config.h1Font}，字号：${config.h1FontSize}pt
3. 2级标题：${config.h2Font}，字号：${config.h2FontSize}pt
4. 3级标题：${config.h3Font}，字号：${config.h3FontSize}pt
5. 行距设置：${config.lineSpacing}磅
6. 页边距配置：上${config.marginTop}cm, 下${config.marginBottom}cm, 左${config.marginLeft}cm, 右${config.marginRight}cm
7. 括号内文字建议字体：${config.bracketFont}，字号：${config.bracketFontSize}pt

请返回格式化后的文档内容建议，保持原文意思不变。`;

        const messages = [
            { role: 'system', content: systemPrompt },
            { role: 'user', content: result.content }
        ];

        const response = await callAIAPI(messages);
        removeLoading(loadingId);

        // 显示格式化建议
        appendMessage('assistant', '📝 排版结果建议：\n\n' + response, targetId);

        // 尝试应用格式化到文档
        try {
            const app = window.Application;
            const doc = app.ActiveDocument;
            if (doc) {
                // A. 设置页面边距
                const setup = doc.PageSetup;
                setup.TopMargin = app.CentimetersToPoints(config.marginTop);
                setup.BottomMargin = app.CentimetersToPoints(config.marginBottom);
                setup.LeftMargin = app.CentimetersToPoints(config.marginLeft);
                setup.RightMargin = app.CentimetersToPoints(config.marginRight);

                // B. 应用格式化到选中内容或全文
                let range;
                if (result.isSelection) {
                    range = doc.Range(result.start, result.end);
                } else {
                    range = doc.Content;
                }
                
                // 1. 设置基本字体格式
                range.Font.NameFarEast = config.titleFont; // 默认使用标题字体作为基准
                range.Font.Size = config.titleFontSize;
                
                // 2. 设置段落格式
                const pf = range.ParagraphFormat;
                // 设置固定行距（磅）
                pf.LineSpacingRule = 4; // wdLineSpaceExactly (4)
                pf.LineSpacing = config.lineSpacing;
                
                // 设置首行缩进2字符
                pf.CharacterUnitFirstLineIndent = 2;
                
                // 3. 处理标题和特殊格式
                const paragraphs = range.Paragraphs;
                const count = paragraphs.Count;
                
                for (let i = 1; i <= count; i++) {
                    try {
                        const para = paragraphs.Item(i);
                        const paraRange = para.Range;
                        const text = paraRange.Text.trim();
                        
                        // 1级标题检测
                        if (/^第[一二三四五六七八九十\d]+[章]/.test(text) || /^[一二三四五六七八九十\d]+[\.、]/.test(text)) {
                            paraRange.Font.NameFarEast = config.h1Font;
                            paraRange.Font.Size = config.h1FontSize;
                            paraRange.Font.Bold = true;
                            para.ParagraphFormat.Alignment = 1; // 居中
                        }
                        // 2级标题检测
                        else if (/^第[一二三四五六七八九十\d]+[节]/.test(text) || /^[（\(][一二三四五六七八九十\d]+[）\)]/.test(text)) {
                            paraRange.Font.NameFarEast = config.h2Font;
                            paraRange.Font.Size = config.h2FontSize;
                            paraRange.Font.Bold = true;
                        }
                        // 3级标题检测
                        else if (/^[①②③④⑤⑥⑦⑧⑨⑩\d]+[\.、]/.test(text)) {
                            paraRange.Font.NameFarEast = config.h3Font;
                            paraRange.Font.Size = config.h3FontSize;
                        }
                        
                        // 处理括号中的内容
                        // 这是一个简化处理：如果整个段落被括号包裹，或者查找段落中的括号
                        // 注意：WPS API 处理局部 Range 比较复杂，这里仅作示意
                        if (/^[（\(].+[）\)]$/.test(text)) {
                            paraRange.Font.NameFarEast = config.bracketFont;
                            paraRange.Font.Size = config.bracketFontSize;
                        }
                        
                    } catch (paraErr) {
                        continue;
                    }
                }
                
                appendMessage('assistant', `✅ 已自动应用您的自定义排版配置：
- 页边距：${config.marginTop}/${config.marginBottom}/${config.marginLeft}/${config.marginRight} cm
- 行距：${config.lineSpacing} 磅
- 标题字体：${config.titleFont} (${config.titleFontSize}pt)
- 1级标题：${config.h1Font} (${config.h1FontSize}pt)
- 2级标题：${config.h2Font} (${config.h2FontSize}pt)
- 3级标题：${config.h3Font} (${config.h3FontSize}pt)
- 括号内容：${config.bracketFont} (${config.bracketFontSize}pt)`, targetId);
            }
        } catch (formatError) {
            console.error("排版应用失败:", formatError);
            appendMessage('assistant', '⚠️ 自动格式化应用时遇到问题：' + formatError.message + '\n您可以根据上面的 AI 建议手动调整。', targetId);
        }
    } catch (e) {
        removeLoading(loadingId);
        appendMessage('assistant', '排版失败：' + e.message, targetId);
    }
}

// 内容续写
async function continueDocument() {
    const targetId = getActiveTargetId();
    const result = getDocumentContent();
    if (!result.success) { appendMessage('assistant', '错误：' + result.error, targetId); return; }
    
    if (targetId !== 'chatMessages') document.getElementById(targetId).innerHTML = '';
    appendMessage('assistant', '正在构思续写内容...', targetId);
    const loadingId = showLoading(targetId);
    
    try {
        const messages = [
            { role: 'system', content: '你是一个擅长续写的作家。请根据用户提供的上文，续写一段文字（约100-200字），保持文风一致。直接输出续写内容。' },
            { role: 'user', content: result.content }
        ];
        const response = await callAIAPI(messages);
        removeLoading(loadingId);
        appendMessage('assistant', '✨ 续写建议：\n\n' + response, targetId);
        appendMessage('assistant', '您可以手动将上述内容复制到文档中。', targetId);
    } catch (e) {
        removeLoading(loadingId);
        appendMessage('assistant', '续写失败：' + e.message, targetId);
    }
}

// 翻译助手
async function translateDocument() {
    const targetId = getActiveTargetId();
    const result = getDocumentContent();
    if (!result.success) { appendMessage('assistant', '错误：' + result.error, targetId); return; }
    
    if (targetId !== 'chatMessages') document.getElementById(targetId).innerHTML = '';
    appendMessage('assistant', '正在进行翻译...', targetId);
    const loadingId = showLoading(targetId);
    
    try {
        const messages = [
            { role: 'system', content: '你是一个精通中英互译的翻译官。请将用户提供的文本翻译成英文（如果原文是英文则翻译成中文）。保持语气地道。直接输出译文。' },
            { role: 'user', content: result.content }
        ];
        const response = await callAIAPI(messages);
        removeLoading(loadingId);
        appendMessage('assistant', '🌍 翻译结果：\n\n' + response, targetId);
    } catch (e) {
        removeLoading(loadingId);
        appendMessage('assistant', '翻译失败：' + e.message, targetId);
    }
}

// 全文总结
async function summarizeDocument() {
    const targetId = getActiveTargetId();
    const result = getDocumentContent();
    if (!result.success) { appendMessage('assistant', '错误：' + result.error, targetId); return; }
    
    if (targetId !== 'chatMessages') document.getElementById(targetId).innerHTML = '';
    appendMessage('assistant', '正在提取摘要...', targetId);
    const loadingId = showLoading(targetId);
    
    try {
        const messages = [
            { role: 'system', content: '你是一个擅长提炼重点的助手。请总结用户提供的内容，列出核心要点（使用列表格式）。直接输出总结内容。' },
            { role: 'user', content: result.content }
        ];
        const response = await callAIAPI(messages);
        removeLoading(loadingId);
        appendMessage('assistant', '💡 内容摘要：\n\n' + response, targetId);
    } catch (e) {
        removeLoading(loadingId);
        appendMessage('assistant', '总结失败：' + e.message, targetId);
    }
}

// 机关公文类型配置
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

// 存储当前生成的文档内容，用于插入操作
let currentGeneratedContent = '';

// 生成机关公文
async function generateDocument(type) {
    const targetId = getActiveTargetId();
    const config = DOCUMENT_TYPES[type];
    
    if (!config) {
        appendMessage('assistant', '错误：不支持的公文类型', targetId);
        return;
    }

    if (targetId !== 'chatMessages') {
        document.getElementById(targetId).innerHTML = '';
    }
    
    appendMessage('assistant', `正在生成${config.name}...`, targetId);
    const loadingId = showLoading(targetId);

    try {
        const messages = [
            { role: 'system', content: config.prompt },
            { role: 'user', content: '请生成一份规范的' + config.name + '范文' }
        ];

        const response = await callAIAPI(messages);
        removeLoading(loadingId);
        
        // 保存生成的内容
        currentGeneratedContent = response;
        
        // 显示生成的内容
        appendMessage('assistant', `📄 **${config.name}范文**\n\n${response}`, targetId);
        
        // 添加"添加到正文"按钮
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
        appendMessage('assistant', `生成${config.name}失败：` + e.message, targetId);
    }
}

// 将生成的内容插入到WPS文档
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
        
        // 获取当前光标位置
        const selection = app.Selection;
        
        // 插入生成的内容
        selection.TypeText(currentGeneratedContent);
        
        // 插入后换行
        selection.TypeParagraph();
        
        alert('✅ 内容已成功添加到文档中！');
        
        // 清空缓存的内容
        currentGeneratedContent = '';
        
    } catch (e) {
        alert('❌ 插入失败：' + e.message);
    }
}



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

// 在文档中添加格式化标注（错别字红底删除线，正确文字绿底插入前面，并保留批注）
function addCommentsToDocument(errors, docContext) {
    const doc = window.Application.ActiveDocument;
    if (!doc) return 0;
    
    let addedCount = 0;
    const searchText = docContext.content;
    const baseOffset = docContext.start;
    
    // 1. 收集所有错误发生的位置
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
    
    // 2. 按位置从后往前排序，避免插入字符后影响前面字符的索引偏移
    errorOccurrences.sort((a, b) => b.pos - a.pos);
    
    // 3. 执行格式化和插入操作
    for (const item of errorOccurrences) {
        try {
            const err = item.err;
            const globalStart = item.globalStart;
            const wrongLength = err.wrong.length;
            
            // 获取错别字的 Range
            const wrongRange = doc.Range(globalStart, globalStart + wrongLength);
            
            // A. 设置错别字格式：红底 + 删除线
            wrongRange.HighlightColorIndex = 6; // wdRed (6)
            wrongRange.Font.StrikeThrough = true;
            
            // B. 添加原有的批注功能
            const commentText = `「${err.wrong}」→「${err.correct}」\n原因：${err.reason || '疑似错别字'}`;
            doc.Comments.Add(wrongRange, commentText);
            
            // C. 在错别字前面插入正确文字
            wrongRange.InsertBefore(err.correct);
            
            // D. 设置插入文字的格式：绿底
            // 插入后，正确文字就在 globalStart 位置开始
            const correctRange = doc.Range(globalStart, globalStart + err.correct.length);
            correctRange.HighlightColorIndex = 4; // wdBrightGreen (4)
            correctRange.Font.Bold = true; // 可选：加粗以便更清晰
            
            addedCount++;
        } catch (e) {
            console.error('应用审阅格式失败:', item.err.wrong, e);
        }
    }
    
    return addedCount;
}
