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
function saveConfigToStorage(config) {
    try {
        window.Application.PluginStorage.setItem('wpsagent_config', JSON.stringify(config));
    } catch (e) {
        console.error('保存配置失败:', e);
    }
}

// 查看存储数据
function viewStorage() {
    try {
        const key = 'wpsagent_config';
        const data = window.Application.PluginStorage.getItem(key);
        if (data) {
            alert('PluginStorage 数据 (' + key + '):\n\n' + JSON.stringify(JSON.parse(data), null, 4));
        } else {
            alert('PluginStorage 中没有找到键名为 "' + key + '" 的数据。');
        }
    } catch (e) {
        alert('读取 PluginStorage 失败: ' + e.message);
    }
}

// 页面加载时初始化
window.onload = function() {
    loadSettingsToUI();
}

// 加载配置到UI
function loadSettingsToUI() {
    const config = loadConfig();
    document.getElementById('apiUrl').value = config.apiUrl;
    document.getElementById('apiKey').value = config.apiKey;
    document.getElementById('modelName').value = config.model;
    
    // 加载排版设置
    document.getElementById('marginTop').value = config.marginTop;
    document.getElementById('marginBottom').value = config.marginBottom;
    document.getElementById('marginLeft').value = config.marginLeft;
    document.getElementById('marginRight').value = config.marginRight;
    document.getElementById('lineSpacing').value = config.lineSpacing;
    document.getElementById('titleFont').value = config.titleFont;
    document.getElementById('titleFontSize').value = config.titleFontSize;
    document.getElementById('h1Font').value = config.h1Font;
    document.getElementById('h1FontSize').value = config.h1FontSize;
    document.getElementById('h2Font').value = config.h2Font;
    document.getElementById('h2FontSize').value = config.h2FontSize;
    document.getElementById('h3Font').value = config.h3Font;
    document.getElementById('h3FontSize').value = config.h3FontSize;
    document.getElementById('bracketFont').value = config.bracketFont;
    document.getElementById('bracketFontSize').value = config.bracketFontSize;
}

// 保存设置
function saveSettings() {
    const apiUrl = document.getElementById('apiUrl').value.trim();
    const apiKey = document.getElementById('apiKey').value.trim();
    const model = document.getElementById('modelName').value.trim();
    
    // 获取排版设置
    const config = {
        apiUrl, 
        apiKey, 
        model,
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
    
    // 保存配置
    saveConfigToStorage(config);
    
    alert('✅ 配置已保存！\n当前模型：' + model);
    closeDialog();
}

// 恢复默认设置
function resetSettings() {
    if (confirm('确定要恢复到默认配置吗？')) {
        saveConfigToStorage(DEFAULT_CONFIG);
        loadSettingsToUI();
        alert('✅ 已恢复默认配置');
    }
}

// 关闭TaskPane
function closeDialog() {
    try {
        let settingsId = window.Application.PluginStorage.getItem("settings_pane_id")
        if (settingsId) {
            let settingsPane = window.Application.GetTaskPane(settingsId)
            if (settingsPane) {
                settingsPane.Visible = false
            }
        }
    } catch (e) {
        console.error('关闭设置面板失败:', e)
    }
}
