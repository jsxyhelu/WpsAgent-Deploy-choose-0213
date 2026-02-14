const config = require('./config.js');

window.onload = function() {
    loadSettingsToUI();
    setupEventListeners();
}

function setupEventListeners() {
    const modelTypeSelect = document.getElementById('modelType');
    const localNeedAuthCheckbox = document.getElementById('localNeedAuth');
    
    modelTypeSelect.addEventListener('change', onModelTypeChange);
    localNeedAuthCheckbox.addEventListener('change', onLocalAuthChange);
}

function onModelTypeChange() {
    const modelType = document.getElementById('modelType').value;
    const cloudConfig = document.getElementById('cloudConfig');
    const localConfig = document.getElementById('localConfig');
    
    if (modelType === 'cloud') {
        cloudConfig.style.display = 'block';
        localConfig.style.display = 'none';
    } else {
        cloudConfig.style.display = 'none';
        localConfig.style.display = 'block';
    }
}

function onLocalAuthChange() {
    const needAuth = document.getElementById('localNeedAuth').checked;
    const localApiKeyGroup = document.getElementById('localApiKeyGroup');
    
    localApiKeyGroup.style.display = needAuth ? 'block' : 'none';
}

async function loadSettingsToUI() {
    const loadedConfig = await config.loadConfig();
    
    document.getElementById('modelType').value = loadedConfig.modelType || 'cloud';
    
    document.getElementById('apiUrl').value = loadedConfig.apiUrl || '';
    document.getElementById('apiKey').value = loadedConfig.apiKey || '';
    document.getElementById('modelName').value = loadedConfig.model || '';
    
    document.getElementById('localApiUrl').value = loadedConfig.apiUrl || '';
    document.getElementById('localModelName').value = loadedConfig.model || '';
    document.getElementById('localNeedAuth').checked = loadedConfig.localNeedAuth || false;
    document.getElementById('localApiKey').value = loadedConfig.apiKey || '';
    
    document.getElementById('marginTop').value = loadedConfig.marginTop;
    document.getElementById('marginBottom').value = loadedConfig.marginBottom;
    document.getElementById('marginLeft').value = loadedConfig.marginLeft;
    document.getElementById('marginRight').value = loadedConfig.marginRight;
    document.getElementById('lineSpacing').value = loadedConfig.lineSpacing;
    document.getElementById('titleFont').value = loadedConfig.titleFont;
    document.getElementById('titleFontSize').value = loadedConfig.titleFontSize;
    document.getElementById('h1Font').value = loadedConfig.h1Font;
    document.getElementById('h1FontSize').value = loadedConfig.h1FontSize;
    document.getElementById('h2Font').value = loadedConfig.h2Font;
    document.getElementById('h2FontSize').value = loadedConfig.h2FontSize;
    document.getElementById('h3Font').value = loadedConfig.h3Font;
    document.getElementById('h3FontSize').value = loadedConfig.h3FontSize;
    document.getElementById('bracketFont').value = loadedConfig.bracketFont;
    document.getElementById('bracketFontSize').value = loadedConfig.bracketFontSize;
    
    onModelTypeChange();
    onLocalAuthChange();
}

async function saveSettings() {
    const modelType = document.getElementById('modelType').value;
    
    let apiUrl, apiKey, model;
    
    if (modelType === 'cloud') {
        apiUrl = document.getElementById('apiUrl').value.trim();
        apiKey = document.getElementById('apiKey').value.trim();
        model = document.getElementById('modelName').value.trim();
        
        if (!apiUrl) {
            alert('请输入 API URL');
            return;
        }
        
        if (!apiKey) {
            alert('请输入 API Key');
            return;
        }
        
        if (!model) {
            alert('请输入模型名称');
            return;
        }
    } else {
        apiUrl = document.getElementById('localApiUrl').value.trim();
        model = document.getElementById('localModelName').value.trim();
        const needAuth = document.getElementById('localNeedAuth').checked;
        
        if (needAuth) {
            apiKey = document.getElementById('localApiKey').value.trim();
            if (!apiKey) {
                alert('请输入本地 API Key');
                return;
            }
        } else {
            apiKey = '';
        }
        
        if (!apiUrl) {
            alert('请输入本地 API 地址');
            return;
        }
        
        if (!model) {
            alert('请输入模型名称');
            return;
        }
    }
    
    const newConfig = {
        modelType,
        apiUrl,
        apiKey,
        model,
        localNeedAuth: modelType === 'local' ? document.getElementById('localNeedAuth').checked : false,
        marginTop: parseFloat(document.getElementById('marginTop').value) || 2.54,
        marginBottom: parseFloat(document.getElementById('marginBottom').value) || 2.54,
        marginLeft: parseFloat(document.getElementById('marginLeft').value) || 3.18,
        marginRight: parseFloat(document.getElementById('marginRight').value) || 3.18,
        lineSpacing: parseFloat(document.getElementById('lineSpacing').value) || 28,
        titleFont: document.getElementById('titleFont').value.trim() || '黑体',
        titleFontSize: parseFloat(document.getElementById('titleFontSize').value) || 16,
        h1Font: document.getElementById('h1Font').value.trim() || '黑体',
        h1FontSize: parseFloat(document.getElementById('h1FontSize').value) || 16,
        h2Font: document.getElementById('h2Font').value.trim() || '黑体',
        h2FontSize: parseFloat(document.getElementById('h2FontSize').value) || 14,
        h3Font: document.getElementById('h3Font').value.trim() || '楷体',
        h3FontSize: parseFloat(document.getElementById('h3FontSize').value) || 12,
        bracketFont: document.getElementById('bracketFont').value.trim() || '仿宋',
        bracketFontSize: parseFloat(document.getElementById('bracketFontSize').value) || 12
    };
    
    const switchResult = await config.switchModelType(modelType, apiUrl, apiKey, model);
    
    if (!switchResult.success) {
        alert('切换模型失败：' + switchResult.error);
        return;
    }
    
    config.saveConfig(newConfig);
    
    const modelDisplayName = modelType === 'cloud' ? `云端模型：${model}` : `本地模型：${model}`;
    alert('✅ 配置已保存！\n当前模型：' + modelDisplayName);
    closeDialog();
}

async function resetSettings() {
    if (confirm('确定要恢复到默认配置吗？')) {
        config.resetConfig();
        await loadSettingsToUI();
        alert('✅ 已恢复默认配置');
    }
}

function closeDialog() {
    try {
        let settingsId = window.Application.PluginStorage.getItem("settings_pane_id");
        if (settingsId) {
            let settingsPane = window.Application.GetTaskPane(settingsId);
            if (settingsPane) {
                settingsPane.Visible = false;
            }
        }
    } catch (e) {
        console.error('关闭设置面板失败:', e);
    }
}
