const logger = require('./logger.js');

async function callAIAPI(messages, config) {
    logger.info('调用 AI API', { messageCount: messages.length, model: config.model });
    
    try {
        const response = await fetch(config.apiUrl, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${config.apiKey}`
            },
            body: JSON.stringify({
                model: config.model,
                messages: messages,
                temperature: 0.7,
                max_tokens: 2048
            })
        });
        
        if (!response.ok) {
            const errorData = await response.json().catch(() => ({}));
            const errorMsg = errorData.error?.message || `HTTP ${response.status}`;
            logger.error('API 调用失败', { status: response.status, error: errorMsg });
            throw new Error(errorMsg);
        }
        
        const data = await response.json();
        const result = data.choices[0].message.content;
        logger.info('API 调用成功', { responseLength: result.length });
        return result;
        
    } catch (error) {
        logger.error('API 调用异常', { error: error.message });
        throw error;
    }
}

async function callAIWithRetry(messages, config, maxRetries = 3) {
    let lastError;
    
    for (let i = 0; i < maxRetries; i++) {
        try {
            return await callAIAPI(messages, config);
        } catch (error) {
            lastError = error;
            logger.warn(`API 调用失败，重试 ${i + 1}/${maxRetries}`, { error: error.message });
            
            if (i < maxRetries - 1) {
                await new Promise(resolve => setTimeout(resolve, 1000 * (i + 1)));
            }
        }
    }
    
    throw lastError;
}
