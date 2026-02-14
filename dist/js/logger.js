const LOG_LEVELS = {
    ERROR: 'ERROR',
    WARN: 'WARN',
    INFO: 'INFO',
    DEBUG: 'DEBUG'
};

let logLevel = LOG_LEVELS.INFO;
let logHistory = [];

function setLogLevel(level) {
    if (Object.values(LOG_LEVELS).includes(level)) {
        logLevel = level;
    }
}

function formatMessage(level, message, data = null) {
    const timestamp = new Date().toISOString();
    const logEntry = {
        timestamp,
        level,
        message,
        data
    };
    logHistory.push(logEntry);
    
    if (logHistory.length > 100) {
        logHistory = logHistory.slice(-100);
    }
    
    return `[${timestamp}] [${level}] ${message}${data ? ' ' + JSON.stringify(data) : ''}`;
}

function error(message, data = null) {
    if (logLevel === LOG_LEVELS.ERROR || logLevel === LOG_LEVELS.WARN || logLevel === LOG_LEVELS.INFO || logLevel === LOG_LEVELS.DEBUG) {
        console.error(formatMessage(LOG_LEVELS.ERROR, message, data));
    }
}

function warn(message, data = null) {
    if (logLevel === LOG_LEVELS.WARN || logLevel === LOG_LEVELS.INFO || logLevel === LOG_LEVELS.DEBUG) {
        console.warn(formatMessage(LOG_LEVELS.WARN, message, data));
    }
}

function info(message, data = null) {
    if (logLevel === LOG_LEVELS.INFO || logLevel === LOG_LEVELS.DEBUG) {
        console.log(formatMessage(LOG_LEVELS.INFO, message, data));
    }
}

function debug(message, data = null) {
    if (logLevel === LOG_LEVELS.DEBUG) {
        console.log(formatMessage(LOG_LEVELS.DEBUG, message, data));
    }
}

function getLogHistory() {
    return [...logHistory];
}

function clearLogHistory() {
    logHistory = [];
}

function exportLogs() {
    const logs = getLogHistory().map(entry => 
        `[${entry.timestamp}] [${entry.level}] ${entry.message}${entry.data ? ' ' + JSON.stringify(entry.data) : ''}`
    ).join('\n');
    
    return logs;
}