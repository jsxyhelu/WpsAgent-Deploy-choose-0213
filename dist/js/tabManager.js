const tabManager = {
    tabs: new Map(),
    
    createOrActivateTab(tabId, label, closeable = true) {
        if (this.tabs.has(tabId)) {
            this.activateTab(tabId);
            debug('激活已有 Tab', { tabId });
            return;
        }
        
        this.tabs.set(tabId, { id: tabId, label, closeable });
        
        const tabItem = this.createTabItem(tabId, label, closeable);
        document.getElementById('tabList').appendChild(tabItem);
        
        const tabContent = this.createTabContent(tabId);
        document.querySelector('.tab-contents').appendChild(tabContent);
        
        this.activateTab(tabId);
        info('创建新 Tab', { tabId, label });
    },
    
    createTabItem(tabId, label, closeable) {
        const tabItem = document.createElement('div');
        tabItem.className = 'tab-item';
        tabItem.setAttribute('data-tab-id', tabId);
        tabItem.onclick = () => this.activateTab(tabId);
        
        const tabLabel = document.createElement('span');
        tabLabel.className = 'tab-label';
        tabLabel.textContent = label;
        tabItem.appendChild(tabLabel);
        
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
    
    createTabContent(tabId) {
        const tabContent = document.createElement('div');
        tabContent.className = 'tab-content';
        tabContent.setAttribute('data-tab-id', tabId);
        
        const contentHtml = this.getTabContentTemplate(tabId);
        tabContent.innerHTML = contentHtml;
        
        return tabContent;
    },
    
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
                        <iframe id="navigationFrame" src="https://www.doubao.com/" style="width: 100%; height: 100%; border: none;" sandbox="allow-same-origin allow-scripts allow-popups allow-forms"></iframe>
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
    
    activateTab(tabId) {
        document.querySelectorAll('.tab-item').forEach(item => {
            item.classList.remove('active');
        });
        document.querySelectorAll('.tab-content').forEach(content => {
            content.classList.remove('active');
        });
        
        const tabItem = document.querySelector(`.tab-item[data-tab-id="${tabId}"]`);
        const tabContent = document.querySelector(`.tab-content[data-tab-id="${tabId}"]`);
        
        if (tabItem) tabItem.classList.add('active');
        if (tabContent) tabContent.classList.add('active');
        
        debug('激活 Tab', { tabId });
    },
    
    closeTab(tabId) {
        if (tabId === 'chat') {
            warn('尝试关闭对话 Tab，已阻止');
            return;
        }
        
        const tabItem = document.querySelector(`.tab-item[data-tab-id="${tabId}"]`);
        const tabContent = document.querySelector(`.tab-content[data-tab-id="${tabId}"]`);
        
        if (!tabItem) return;
        
        const wasActive = tabItem.classList.contains('active');
        
        tabItem.remove();
        if (tabContent) tabContent.remove();
        
        this.tabs.delete(tabId);
        
        if (wasActive) {
            this.activateTab('chat');
        }
        
        info('关闭 Tab', { tabId });
    }
};