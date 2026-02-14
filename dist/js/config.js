const DEFAULT_CONFIG = {
    modelType: 'cloud',
    apiUrl: 'https://api.deepseek.com/chat/completions',
    apiKey: '',
    model: 'deepseek-chat',
    localNeedAuth: false,
    marginTop: 2.54,
    marginBottom: 2.54,
    marginLeft: 3.18,
    marginRight: 3.18,
    lineSpacing: 28,
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

let currentConfig = { ...DEFAULT_CONFIG };

async function loadConfig() {
    try {
        const saved = window.Application.PluginStorage.getItem('wpsagent_config');
        if (saved) {
            const parsed = JSON.parse(saved);
            currentConfig = { ...DEFAULT_CONFIG, ...parsed };
            info('从本地存储加载配置', { modelType: currentConfig.modelType, model: currentConfig.model });
            return currentConfig;
        }
    } catch (e) {
        error('加载配置失败', { error: e.message });
    }
    return { ...DEFAULT_CONFIG };
}

async function loadConfigFromServer() {
    try {
        const response = await fetch('/api/config');
        if (response.ok) {
            const serverConfig = await response.json();
            currentConfig = { ...DEFAULT_CONFIG, ...serverConfig };
            info('从服务器加载配置', { modelType: currentConfig.modelType, model: currentConfig.model });
            return currentConfig;
        }
    } catch (e) {
        warn('从服务器加载配置失败，使用本地配置', { error: e.message });
    }
    return loadConfig();
}

function saveConfig(config) {
    try {
        currentConfig = { ...DEFAULT_CONFIG, ...config };
        window.Application.PluginStorage.setItem('wpsagent_config', JSON.stringify(currentConfig));
        info('保存配置', { modelType: currentConfig.modelType, model: currentConfig.model });
    } catch (e) {
        error('保存配置失败', { error: e.message });
    }
}

function getConfig() {
    return currentConfig;
}

function resetConfig() {
    currentConfig = { ...DEFAULT_CONFIG };
    saveConfig(currentConfig);
    info('重置配置为默认值');
    return currentConfig;
}

async function switchModelType(modelType, apiUrl, apiKey, model) {
    try {
        const response = await fetch('/api/switch-model', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                modelType,
                apiUrl,
                apiKey,
                model
            })
        });
        
        if (response.ok) {
            const result = await response.json();
            info('切换模型类型成功', { modelType, model });
            return { success: true, message: result.message };
        } else {
            const error = await response.json();
            error('切换模型类型失败', { error: error.error });
            return { success: false, error: error.error };
        }
    } catch (e) {
        error('切换模型类型异常', { error: e.message });
        return { success: false, error: e.message };
    }
}

function getModelDisplayName() {
    if (currentConfig.modelType === 'cloud') {
        return `云端模型: ${currentConfig.model}`;
    } else {
        return `本地模型: ${currentConfig.model}`;
    }
}

function isLocalModel() {
    return currentConfig.modelType === 'local';
}

function isCloudModel() {
    return currentConfig.modelType === 'cloud';
}