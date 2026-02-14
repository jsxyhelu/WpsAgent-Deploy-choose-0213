
//这个函数在整个wps加载项中是第一个执行的
function OnAddinLoad(ribbonUI){
    if (typeof (window.Application.ribbonUI) != "object"){
		window.Application.ribbonUI = ribbonUI
    }
    
    if (typeof (window.Application.Enum) != "object") { // 如果没有内置枚举值
        window.Application.Enum = WPS_Enum
    }
    return true
}

function OnAction(control) {
    const eleId = control.Id
    switch (eleId) {
        case "btnOpenSidebar":
            {
                let tsId = window.Application.PluginStorage.getItem("taskpane_id")
                if (!tsId) {
                    let tskpane = window.Application.CreateTaskPane(GetUrlPath() + "/ui/taskpane.html")
                    let id = tskpane.ID
                    window.Application.PluginStorage.setItem("taskpane_id", id)
                    tskpane.DockPosition = window.Application.Enum.msoCTPDockPositionRight
                    tskpane.Visible = true
                } else {
                    let tskpane = window.Application.GetTaskPane(tsId)
                    tskpane.DockPosition = window.Application.Enum.msoCTPDockPositionRight
                    tskpane.Visible = !tskpane.Visible
                }
            }
            break
        case "btnSettings":
            {
                // 使用TaskPane显示设置界面
                let settingsId = window.Application.PluginStorage.getItem("settings_pane_id")
                if (!settingsId) {
                    let settingsPane = window.Application.CreateTaskPane(GetUrlPath() + "/ui/dialog.html")
                    let id = settingsPane.ID
                    window.Application.PluginStorage.setItem("settings_pane_id", id)
                    settingsPane.Width = 450
                    settingsPane.DockPosition = window.Application.Enum.msoCTPDockPositionRight
                    settingsPane.Visible = true
                } else {
                    let settingsPane = window.Application.GetTaskPane(settingsId)
                    settingsPane.Width = 450
                    settingsPane.DockPosition = window.Application.Enum.msoCTPDockPositionRight
                    settingsPane.Visible = !settingsPane.Visible
                }
            }
            break
        case "btnProofread":
        case "btnPolish":
        case "btnFormat":
        case "btnContinue":
        case "btnWrite":
        case "btnTranslate":
        case "btnSummarize":
        case "btnNavigate":
            {
                // AI功能按钮：打开侧边栏并创建对应 tab
                const buttonMap = {
                    'btnProofread': { id: 'proofread', label: '校对' },
                    'btnPolish': { id: 'polish', label: '润色' },
                    'btnFormat': { id: 'format', label: '排版' },
                    'btnContinue': { id: 'continue', label: '续写' },
                    'btnWrite': { id: 'write', label: '写作' },
                    'btnTranslate': { id: 'translate', label: '翻译' },
                    'btnSummarize': { id: 'summarize', label: '总结' },
                    'btnNavigate': { id: 'navigate', label: '导航' }
                };
                
                const tabInfo = buttonMap[eleId];
                if (tabInfo) {
                    // 1. 设置待处理的 tab 信息到存储中（可靠的通信方式）
                    window.Application.PluginStorage.setItem('pending_tab', JSON.stringify(tabInfo));

                    // 2. 确保侧边栏已打开
                    let tsId = window.Application.PluginStorage.getItem("taskpane_id")
                    let tskpane;
                    
                    if (!tsId) {
                        tskpane = window.Application.CreateTaskPane(GetUrlPath() + "/ui/taskpane.html")
                        window.Application.PluginStorage.setItem("taskpane_id", tskpane.ID)
                        tskpane.DockPosition = window.Application.Enum.msoCTPDockPositionRight
                        tskpane.Visible = true
                    } else {
                        tskpane = window.Application.GetTaskPane(tsId)
                        tskpane.DockPosition = window.Application.Enum.msoCTPDockPositionRight
                        tskpane.Visible = true
                    }

                    // 3. 尝试直接调用（可选，作为补充）
                    try {
                        if (tskpane && tskpane.ContentWindow && tskpane.ContentWindow.tabManager) {
                            tskpane.ContentWindow.tabManager.createOrActivateTab(tabInfo.id, tabInfo.label);
                            window.Application.PluginStorage.setItem('pending_tab', '');
                        }
                    } catch (e) {
                        // 失败了也没关系，taskpane 会通过轮询机制处理 pending_tab
                    }
                }
            }
            break
        case "btnHelp":
            {
                // 帮助按钮：打开帮助文档
                let helpId = window.Application.PluginStorage.getItem("help_pane_id")
                if (!helpId) {
                    let helpPane = window.Application.CreateTaskPane(GetUrlPath() + "/ui/help.html")
                    let id = helpPane.ID
                    window.Application.PluginStorage.setItem("help_pane_id", id)
                    helpPane.Width = 900
                    helpPane.DockPosition = window.Application.Enum.msoCTPDockPositionRight
                    helpPane.Visible = true
                } else {
                    let helpPane = window.Application.GetTaskPane(helpId)
                    helpPane.Width = 900
                    helpPane.DockPosition = window.Application.Enum.msoCTPDockPositionRight
                    helpPane.Visible = !helpPane.Visible
                }
            }
            break
        case "btnRestart":
            {
                // 关闭 WPS 功能
                let docs = window.Application.Documents;
                let hasUnsaved = false;
                
                for (let i = 1; i <= docs.Count; i++) {
                    let doc = docs.Item(i);
                    if (doc.Path == "" && !doc.Saved) {
                        hasUnsaved = true;
                        break;
                    }
                }

                let msg = "确定要保存所有文档并关闭 WPS 吗？";
                if (hasUnsaved) {
                    msg += "\n注意：有新建文档尚未保存，关闭后可能无法恢复。";
                }

                if (confirm(msg)) {
                    // 1. 保存所有已命名的文档
                    for (let i = 1; i <= docs.Count; i++) {
                        try {
                            let doc = docs.Item(i);
                            if (doc.Path != "" || !doc.Saved) {
                                doc.Save();
                            }
                        } catch (e) {}
                    }

                    // 2. 退出当前进程
                    window.Application.Quit();
                }
            }
            break
        default:
            break
    }
    return true
}

function GetImage(control) {
    const eleId = control.Id
    switch (eleId) {
        case "btnOpenSidebar":
            return "images/robot.svg"
        case "btnSettings":
            return "images/settings.svg"
        case "btnProofread":
            return "images/proofread.svg"
        case "btnPolish":
            return "images/polish.svg"
        case "btnFormat":
            return "images/format.svg"
        case "btnContinue":
            return "images/continue.svg"
        case "btnWrite":
            return "images/write.svg"
        case "btnTranslate":
            return "images/translate.svg"
        case "btnSummarize":
            return "images/summarize.svg"
        case "btnNavigate":
            return "images/navigate.svg"
        case "btnHelp":
            return "images/help.svg"
        case "btnRestart":
            return "images/close.svg"
        default:
            ;
    }
    return "images/newFromTemp.svg"
}
