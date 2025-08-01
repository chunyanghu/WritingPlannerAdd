/* global document, Office, Word, Chart, Notification, $ */

// =================================================================
// App State and Initialization
// =================================================================

const app = {
    data: {
        projects: [],
        activeProjectId: null,
    },
    chartInstance: null,
    isInitialized: false,
};

Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        if (app.isInitialized) return;
        app.isInitialized = true;

        try {
            loadAppData(); // This now loads from document settings
            setupEventListeners();
            renderProjectSelector();
            updateAllDisplaysForActiveProject();
            initChart();
            setInterval(checkReminder, 60000);
        } catch (error) {
            console.error("Initialization error:", error);
            showMessage("初始化失败: " + error.message, "danger");
        }
    }
});

function setupEventListeners() {
    document.getElementById("projectSelector").onchange = switchActiveProject;
    document.getElementById("newProjectBtn").onclick = showNewProjectModal;
    document.getElementById("deleteProjectBtn").onclick = showDeleteConfirmModal;
    document.getElementById("clearDataBtn").onclick = showClearDataConfirmModal; // 新增清除数据按钮的事件
    document.getElementById("savePlan").onclick = saveCurrentProjectPlan;
    document.getElementById("updateProgress").onclick = updateCurrentProjectProgress;

    // 模态框按钮的监听器
    document.getElementById("confirmNewProjectBtn").onclick = handleCreateNewProject;
    document.getElementById("confirmDeleteBtn").onclick = handleDeleteCurrentProject;
    document.getElementById("confirmClearDataBtn").onclick = handleClearAllData; // 新增清除数据确认按钮的事件

    if ("Notification" in window && Notification.permission === "default") {
        Notification.requestPermission().catch(err => console.error('Notification permission error:', err));
    }
}

// =================================================================
// Data Management (Using Office.Settings for document-bound storage)
// =================================================================

function loadAppData() {
    const savedDataString = Office.context.document.settings.get('writingAppMultiTaskData');
    
    if (savedDataString) {
        try {
            const savedData = JSON.parse(savedDataString);
            app.data = {
                projects: savedData.projects || [],
                activeProjectId: savedData.activeProjectId || null,
            };
        } catch (e) {
            console.error("Failed to parse data from document settings.", e);
            initializeEmptyData();
        }
    } else {
        initializeEmptyData();
    }

    if (app.data.projects.length === 0) {
        const defaultProject = createProjectObject("我的第一个项目");
        app.data.projects.push(defaultProject);
        app.data.activeProjectId = defaultProject.id;
        saveAppData(); 
    }

    if (!app.data.activeProjectId && app.data.projects.length > 0) {
        app.data.activeProjectId = app.data.projects[0].id;
    }
}

function saveAppData(callback) {
    try {
        const dataString = JSON.stringify(app.data);
        Office.context.document.settings.set('writingAppMultiTaskData', dataString);

        Office.context.document.settings.saveAsync(function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                console.log('Settings saved successfully to the document.');
                if (callback) callback(true);
            } else {
                console.error('Failed to save settings: ' + asyncResult.error.message);
                showMessage("数据同步到文档失败！", "danger");
                if (callback) callback(false);
            }
        });
    } catch (e) {
        console.error("Failed to stringify app data.", e);
        if (callback) callback(false);
    }
}

function initializeEmptyData() {
    app.data = {
        projects: [],
        activeProjectId: null,
    };
}

// =================================================================
// Project Management & Data Clearing
// =================================================================

function createProjectObject(name) {
    return {
        id: 'proj_' + Date.now() + Math.random(),
        name: name,
        targetWords: 10000,
        deadline: '',
        dailyTarget: 500,
        reminderTime: '09:00',
        startDate: new Date().toISOString().split('T')[0],
        progress: [],
    };
}

function getActiveProject() {
    return app.data.projects.find(p => p.id === app.data.activeProjectId);
}

// --- Logic for Modals (Create, Delete, Clear) ---

function showNewProjectModal() {
    document.getElementById('newProjectNameInput').value = `新项目 ${app.data.projects.length + 1}`;
    $('#newProjectModal').modal('show');
}

function handleCreateNewProject() {
    const projectName = document.getElementById('newProjectNameInput').value.trim();
    if (projectName) {
        const newProject = createProjectObject(projectName);
        app.data.projects.push(newProject);
        app.data.activeProjectId = newProject.id;
        
        saveAppData();
        renderProjectSelector();
        updateAllDisplaysForActiveProject();
        $('#newProjectModal').modal('hide');
    } else {
        showMessage("项目名称不能为空！", "warning");
    }
}

function showDeleteConfirmModal() {
    const project = getActiveProject();
    if (!project) return;
    document.getElementById('deleteProjectName').textContent = project.name;
    $('#deleteConfirmModal').modal('show');
}

function handleDeleteCurrentProject() {
    const projectToDeleteId = app.data.activeProjectId;
    app.data.projects = app.data.projects.filter(p => p.id !== projectToDeleteId);
    
    if (app.data.projects.length > 0) {
        app.data.activeProjectId = app.data.projects[0].id;
    } else {
        // If all projects are deleted, re-initialize to a clean state
        initializeEmptyData(); 
        const defaultProject = createProjectObject("我的第一个项目");
        app.data.projects.push(defaultProject);
        app.data.activeProjectId = defaultProject.id;
    }
    
    saveAppData();
    $('#deleteConfirmModal').modal('hide');
    renderProjectSelector();
    updateAllDisplaysForActiveProject();
}

function showClearDataConfirmModal() {
    $('#clearDataConfirmModal').modal('show');
}

function handleClearAllData() {
    // Remove the setting from the document
    Office.context.document.settings.remove('writingAppMultiTaskData');
    
    saveAppData(() => {
        // After saving (which now saves an empty setting), re-initialize the app state
        initializeEmptyData();
        const defaultProject = createProjectObject("我的第一个项目");
        app.data.projects.push(defaultProject);
        app.data.activeProjectId = defaultProject.id;
        
        // No need to save again, just update UI
        renderProjectSelector();
        updateAllDisplaysForActiveProject();
        
        $('#clearDataConfirmModal').modal('hide');
        showMessage("本文档中的所有插件数据已被清除。", "success");
    });
}

// (The rest of the functions: switchActiveProject, saveCurrentProjectPlan, etc. remain the same as the multi-task version you have)
// ... PASTE ALL OTHER FUNCTIONS FROM YOUR PREVIOUS JS FILE HERE ...
// Make sure to include:
// - switchActiveProject()
// - saveCurrentProjectPlan()
// - updateCurrentProjectProgress()
// - All UI Rendering functions (renderProjectSelector, updateAllDisplaysForActiveProject, clearAllDisplays, showMessage)
// - All Calculation functions (countWords, getCurrentTotalWords, getTodayWords, formatDate)
// - All Chart and History functions (initChart, updateChart, updateHistory)
// - The Reminder function (checkReminder)
// I am pasting them below for completeness.
function switchActiveProject() {
    const selector = document.getElementById("projectSelector");
    app.data.activeProjectId = selector.value;
    // No need to save on switch, it's just a view change. Save happens on action.
    updateAllDisplaysForActiveProject();
}

function saveCurrentProjectPlan() {
    const project = getActiveProject();
    if (!project) {
        showMessage("没有活动项目，无法保存。", "warning");
        return;
    }

    try {
        project.name = document.getElementById("projectName").value;
        project.targetWords = parseInt(document.getElementById("targetWords").value) || 0;
        project.deadline = document.getElementById("deadline").value;
        project.dailyTarget = parseInt(document.getElementById("dailyTarget").value) || 0;
        project.reminderTime = document.getElementById("reminderTime").value;

        if (!project.name || !project.targetWords || !project.deadline) {
            showMessage('请填写所有必填项！', 'warning');
            return;
        }

        saveAppData(() => {
            showMessage(`项目 "${project.name}" 的计划已保存！`, 'success');
        });
        renderProjectSelector();
        updateAllDisplaysForActiveProject();
    } catch (error) {
        console.error('Save plan error:', error);
        showMessage('保存计划时出错：' + error.message, 'danger');
    }
}

async function updateCurrentProjectProgress() {
    const project = getActiveProject();
    if (!project) {
        showMessage("请先选择一个项目再更新进度。", "warning");
        return;
    }

    try {
        await Word.run(async (context) => {
            const body = context.document.body;
            context.load(body, 'text');
            await context.sync();

            const wordCount = countWords(body.text);
            const today = new Date().toISOString().split('T')[0];
            
            if (!project.progress) project.progress = [];
            const todayProgress = project.progress.find(p => p.date === today);

            if (todayProgress) {
                todayProgress.words = wordCount;
            } else {
                project.progress.push({ date: today, words: wordCount });
            }

            saveAppData(() => {
                showMessage(`进度已更新！当前字数：${wordCount}`, 'success');
            });
            updateAllDisplaysForActiveProject();
        });
    } catch (error) {
        console.error('Update progress error:', error);
        showMessage('更新进度时出错：' + error.message, 'danger');
    }
}

function renderProjectSelector() {
    const selector = document.getElementById("projectSelector");
    selector.innerHTML = '';
    if (app.data.projects.length > 0) {
        app.data.projects.forEach(project => {
            const option = document.createElement('option');
            option.value = project.id;
            option.textContent = project.name;
            selector.appendChild(option);
        });
        if (app.data.activeProjectId) {
            selector.value = app.data.activeProjectId;
        }
    }
}

function updateAllDisplaysForActiveProject() {
    const project = getActiveProject();
    if (!project) {
        clearAllDisplays();
        return;
    }
    
    document.getElementById("projectName").value = project.name || '';
    document.getElementById("targetWords").value = project.targetWords || '';
    document.getElementById("deadline").value = project.deadline || '';
    document.getElementById("dailyTarget").value = project.dailyTarget || '';
    document.getElementById("reminderTime").value = project.reminderTime || '09:00';

    const currentWords = getCurrentTotalWords(project);
    const progress = Math.min(100, (currentWords / (project.targetWords || 1) * 100)).toFixed(1);
    
    const progressBar = document.getElementById("progressBar");
    progressBar.style.width = progress + '%';
    progressBar.textContent = progress + '%';
    
    document.getElementById("currentWords").textContent = currentWords.toLocaleString();
    document.getElementById("targetWordsDisplay").textContent = (project.targetWords || 0).toLocaleString();
    
    const today = new Date();
    const deadline = project.deadline ? new Date(project.deadline) : today;
    const daysLeft = Math.ceil((deadline - today) / (1000 * 60 * 60 * 24));
    document.getElementById("daysLeft").textContent = daysLeft >= 0 ? daysLeft : 0;
    
    document.getElementById("todayWords").textContent = getTodayWords(project).toLocaleString();

    updateHistory(project);
    updateChart(project);
}

function clearAllDisplays() {
    const formIds = ["projectName", "targetWords", "deadline", "dailyTarget", "reminderTime"];
    formIds.forEach(id => { document.getElementById(id).value = ''; });
    document.getElementById("reminderTime").value = '09:00';
    
    const pIds = ["currentWords", "targetWordsDisplay", "daysLeft", "todayWords"];
    pIds.forEach(id => { document.getElementById(id).textContent = '0'; });

    const progressBar = document.getElementById("progressBar");
    progressBar.style.width = '0%';
    progressBar.textContent = '0%';
    
    document.getElementById("historyList").innerHTML = '<div class="list-group-item">请创建一个新项目</div>';
    if(app.chartInstance) {
        app.chartInstance.data.labels = [];
        app.chartInstance.data.datasets.forEach(dataset => dataset.data = []);
        app.chartInstance.update();
    }
}

function showMessage(message, type) {
    try {
        const container = document.querySelector('.container-fluid');
        if (!container) return;

        const alertDiv = document.createElement('div');
        alertDiv.className = `alert alert-${type} alert-dismissible fade show`;
        alertDiv.style.position = 'absolute';
        alertDiv.style.top = '50px';
        alertDiv.style.left = '15px';
        alertDiv.style.right = '15px';
        alertDiv.style.zIndex = '1050';
        alertDiv.innerHTML = `
            ${message}
            <button type="button" class="close" onclick="this.parentElement.remove()">
                <span>×</span>
            </button>
        `;
        
        container.insertBefore(alertDiv, container.firstChild);
        setTimeout(() => {
            alertDiv.style.transition = 'opacity 0.5s ease';
            alertDiv.style.opacity = '0';
            setTimeout(() => alertDiv.remove(), 500);
        }, 3000);
    } catch (error) {
        console.error('Show message error:', error);
    }
}

function countWords(text) {
    if (!text) return 0;
    text = text.replace(/\s+/g, ' ').trim();
    const chineseChars = (text.match(/[\u4e00-\u9fa5]/g) || []).length;
    const englishWords = text.replace(/[\u4e00-\u9fa5]/g, ' ').split(/\s+/).filter(Boolean).length;
    return chineseChars + englishWords;
}

function getCurrentTotalWords(project) {
    if (!project || !project.progress || project.progress.length === 0) return 0;
    const latestProgress = project.progress.reduce((latest, current) => new Date(current.date) > new Date(latest.date) ? current : latest);
    return latestProgress.words || 0;
}

function getTodayWords(project) {
    if (!project || !project.progress) return 0;
    
    const todayStr = new Date().toISOString().split('T')[0];
    const yesterday = new Date();
    yesterday.setDate(yesterday.getDate() - 1);
    const yesterdayStr = yesterday.toISOString().split('T')[0];

    const todayProgress = project.progress.find(p => p.date === todayStr);
    const yesterdayProgress = project.progress.find(p => p.date === yesterdayStr);

    if (!todayProgress) return 0;
    const startOfDayWords = yesterdayProgress ? yesterdayProgress.words : 0;
    return Math.max(0, todayProgress.words - startOfDayWords);
}

function formatDate(dateString) {
    if (!dateString) return '';
    try {
        const date = new Date(dateString);
        return `${date.getMonth() + 1}月${date.getDate()}日`;
    } catch {
        return dateString;
    }
}

function initChart() {
    if (app.chartInstance) app.chartInstance.destroy();
    const ctx = document.getElementById('progressChart').getContext('2d');
    app.chartInstance = new Chart(ctx, {
        type: 'line',
        data: { labels: [], datasets: [
            { label: '累计字数', data: [], borderColor: '#2b579a', backgroundColor: 'rgba(43, 87, 154, 0.1)', tension: 0.1, yAxisID: 'y' },
            { label: '每日字数', data: [], borderColor: '#28a745', backgroundColor: 'rgba(40, 167, 69, 0.1)', tension: 0.1, yAxisID: 'y1' }
        ]},
        options: {
            responsive: true, maintainAspectRatio: false,
            scales: {
                y: { beginAtZero: true, position: 'left' },
                y1: { beginAtZero: true, position: 'right', grid: { drawOnChartArea: false } }
            }
        }
    });
}

function updateChart(project) {
    if (!app.chartInstance || !project || !project.progress) return;
    
    const sortedProgress = [...project.progress].sort((a, b) => new Date(a.date) - new Date(b.date));
    const labels = [], totalWords = [], dailyWords = [];
    
    sortedProgress.forEach((record, index) => {
        labels.push(formatDate(record.date));
        totalWords.push(record.words);
        const prevWords = index > 0 ? sortedProgress[index - 1].words : 0;
        dailyWords.push(Math.max(0, record.words - prevWords));
    });
    
    app.chartInstance.data.labels = labels;
    app.chartInstance.data.datasets[0].data = totalWords;
    app.chartInstance.data.datasets[1].data = dailyWords;
    app.chartInstance.update();
}

function updateHistory(project) {
    const historyList = document.getElementById("historyList");
    historyList.innerHTML = '';
    if (!project || !project.progress || project.progress.length === 0) {
        historyList.innerHTML = '<div class="list-group-item">暂无写作记录</div>';
        return;
    }

    const sortedProgress = [...project.progress].sort((a, b) => new Date(b.date) - new Date(a.date));
    
    sortedProgress.slice(0, 10).forEach((record, index) => {
        const prevRecord = sortedProgress[index + 1];
        const dailyWords = prevRecord ? record.words - prevRecord.words : record.words;
        
        const item = document.createElement('div');
        item.className = 'list-group-item';
        item.innerHTML = `
            <div class="d-flex justify-content-between align-items-center">
                <div><strong>${formatDate(record.date)}</strong><br><small>总字数: ${record.words.toLocaleString()}</small></div>
                <span class="badge badge-primary badge-pill">+${Math.max(0, dailyWords).toLocaleString()}</span>
            </div>`;
        historyList.appendChild(item);
    });
}

function checkReminder() {
    const project = getActiveProject();
    if (!project || !project.reminderTime) return;

    const now = new Date();
    const currentTime = `${now.getHours().toString().padStart(2, '0')}:${now.getMinutes().toString().padStart(2, '0')}`;

    if (currentTime === project.reminderTime) {
        const remaining = (project.dailyTarget || 0) - getTodayWords(project);
        if (remaining > 0) {
            showMessage(`写作提醒 ("${project.name}")：今天还需写 ${remaining} 字！`, 'warning');
            if ("Notification" in window && Notification.permission === "granted") {
                new Notification("写作提醒", {
                    body: `项目 "${project.name}" 今天还需要写 ${remaining} 字才能完成目标！`,
                    icon: 'assets/icon-128.png'
                });
            }
        }
    }
}