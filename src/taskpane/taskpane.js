/* global document, Office, Word, Chart, Notification */

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
            loadAppData();
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
    document.getElementById("newProjectBtn").onclick = createNewProject;
    document.getElementById("deleteProjectBtn").onclick = deleteCurrentProject;
    document.getElementById("savePlan").onclick = saveCurrentProjectPlan;
    document.getElementById("updateProgress").onclick = updateCurrentProjectProgress;

    if ("Notification" in window && Notification.permission === "default") {
        Notification.requestPermission().catch(err => console.error('Notification permission error:', err));
    }
}

// =================================================================
// Data Management (localStorage)
// =================================================================

function loadAppData() {
    const savedData = JSON.parse(localStorage.getItem('writingAppMultiTask') || '{}');
    app.data = {
        projects: savedData.projects || [],
        activeProjectId: savedData.activeProjectId || null,
    };

    if (app.data.projects.length === 0) {
        // 如果没有任何项目，创建一个默认的
        const defaultProject = createProjectObject("我的第一个项目");
        app.data.projects.push(defaultProject);
        app.data.activeProjectId = defaultProject.id;
    }

    if (!app.data.activeProjectId && app.data.projects.length > 0) {
        app.data.activeProjectId = app.data.projects[0].id;
    }
}

function saveAppData() {
    localStorage.setItem('writingAppMultiTask', JSON.stringify(app.data));
}

// =================================================================
// Project Management
// =================================================================

function createProjectObject(name) {
    return {
        id: 'proj_' + Date.now(), // Unique ID
        name: name,
        targetWords: 10000,
        deadline: '',
        dailyTarget: 500,
        reminderTime: '09:00',
        startDate: new Date().toISOString().split('T')[0],
        progress: [], // [{ date: 'YYYY-MM-DD', words: 123 }, ...]
    };
}

function getActiveProject() {
    return app.data.projects.find(p => p.id === app.data.activeProjectId);
}

function createNewProject() {
    const projectName = prompt("请输入新项目的名称：", `新项目 ${app.data.projects.length + 1}`);
    if (projectName) {
        const newProject = createProjectObject(projectName);
        app.data.projects.push(newProject);
        app.data.activeProjectId = newProject.id;
        
        saveAppData();
        renderProjectSelector();
        updateAllDisplaysForActiveProject();
    }
}

function deleteCurrentProject() {
    const project = getActiveProject();
    if (!project) return;

    if (confirm(`您确定要删除项目 "${project.name}" 吗？此操作无法撤销。`)) {
        app.data.projects = app.data.projects.filter(p => p.id !== project.id);
        
        if (app.data.projects.length > 0) {
            app.data.activeProjectId = app.data.projects[0].id;
        } else {
            app.data.activeProjectId = null;
        }

        saveAppData();

        if (app.data.projects.length === 0) {
            // 如果都删完了，新建一个默认的
            createNewProject();
        } else {
            renderProjectSelector();
            updateAllDisplaysForActiveProject();
        }
    }
}

function switchActiveProject() {
    const selector = document.getElementById("projectSelector");
    app.data.activeProjectId = selector.value;
    saveAppData();
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

        saveAppData();
        showMessage(`项目 "${project.name}" 的计划已保存！`, 'success');
        renderProjectSelector(); // 更新下拉列表中的项目名称
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

            saveAppData();
            updateAllDisplaysForActiveProject();
            showMessage(`进度已更新！当前字数：${wordCount}`, 'success');
        });
    } catch (error) {
        console.error('Update progress error:', error);
        showMessage('更新进度时出错：' + error.message, 'danger');
    }
}

// =================================================================
// UI Rendering and Display Updates
// =================================================================

function renderProjectSelector() {
    const selector = document.getElementById("projectSelector");
    selector.innerHTML = '';
    app.data.projects.forEach(project => {
        const option = document.createElement('option');
        option.value = project.id;
        option.textContent = project.name;
        selector.appendChild(option);
    });
    selector.value = app.data.activeProjectId;
}

function updateAllDisplaysForActiveProject() {
    const project = getActiveProject();
    if (!project) {
        // 如果没有活动项目（例如全部删除了），清空显示
        clearAllDisplays();
        return;
    }
    
    // 1. 更新计划设置表单
    document.getElementById("projectName").value = project.name || '';
    document.getElementById("targetWords").value = project.targetWords || '';
    document.getElementById("deadline").value = project.deadline || '';
    document.getElementById("dailyTarget").value = project.dailyTarget || '';
    document.getElementById("reminderTime").value = project.reminderTime || '09:00';

    // 2. 更新进度追踪面板
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
    document.getElementById("daysLeft").textContent = daysLeft > 0 ? daysLeft : 0;
    
    document.getElementById("todayWords").textContent = getTodayWords(project).toLocaleString();

    // 3. 更新统计分析面板
    updateHistory(project);
    updateChart(project);
}

function clearAllDisplays() {
    // 清空表单
    document.getElementById("projectName").value = '';
    document.getElementById("targetWords").value = '';
    document.getElementById("deadline").value = '';
    document.getElementById("dailyTarget").value = '';
    document.getElementById("reminderTime").value = '09:00';
    // 清空进度
    const progressBar = document.getElementById("progressBar");
    progressBar.style.width = '0%';
    progressBar.textContent = '0%';
    document.getElementById("currentWords").textContent = '0';
    document.getElementById("targetWordsDisplay").textContent = '0';
    document.getElementById("daysLeft").textContent = '0';
    document.getElementById("todayWords").textContent = '0';
    // 清空历史和图表
    document.getElementById("historyList").innerHTML = '<div class="list-group-item">请选择一个项目</div>';
    if(app.chartInstance) {
        app.chartInstance.data.labels = [];
        app.chartInstance.data.datasets.forEach(dataset => dataset.data = []);
        app.chartInstance.update();
    }
}


function showMessage(message, type) {
    try {
        const alertDiv = document.createElement('div');
        alertDiv.className = `alert alert-${type} alert-dismissible fade show`;
        alertDiv.style.position = 'fixed';
        alertDiv.style.top = '10px';
        alertDiv.style.left = '10px';
        alertDiv.style.right = '10px';
        alertDiv.style.zIndex = '9999';
        alertDiv.innerHTML = `
            ${message}
            <button type="button" class="close" onclick="this.parentElement.remove()">
                <span>×</span>
            </button>
        `;
        document.body.appendChild(alertDiv);
        setTimeout(() => {
            alertDiv.remove();
        }, 4000);
    } catch (error) {
        console.error('Show message error:', error);
    }
}

// =================================================================
// Calculation and Utility Functions
// =================================================================

function countWords(text) {
    if (!text) return 0;
    text = text.replace(/\s+/g, ' ').trim();
    const chineseChars = (text.match(/[\u4e00-\u9fa5]/g) || []).length;
    const englishWords = text.replace(/[\u4e00-\u9fa5]/g, ' ').split(/\s+/).filter(Boolean).length;
    return chineseChars + englishWords;
}

function getCurrentTotalWords(project) {
    if (!project || !project.progress || project.progress.length === 0) return 0;
    return project.progress.reduce((latest, current) => new Date(current.date) > new Date(latest.date) ? current : latest).words || 0;
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
    return Math.max(0, todayProgress.words - (yesterdayProgress ? yesterdayProgress.words : 0));
}

function formatDate(dateString) {
    try {
        const date = new Date(dateString);
        return `${date.getMonth() + 1}月${date.getDate()}日`;
    } catch {
        return dateString;
    }
}

// =================================================================
// Chart and History
// =================================================================

function initChart() {
    if (app.chartInstance) {
        app.chartInstance.destroy();
    }
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
    
    const labels = [];
    const totalWords = [];
    const dailyWords = [];
    
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
        const prevWords = (sortedProgress[index + 1] || { words: 0 }).words;
        const dailyWords = Math.max(0, record.words - prevWords);
        
        const item = document.createElement('div');
        item.className = 'list-group-item';
        item.innerHTML = `
            <div class="d-flex justify-content-between align-items-center">
                <div><strong>${formatDate(record.date)}</strong><br><small>总字数: ${record.words.toLocaleString()}</small></div>
                <span class="badge badge-primary badge-pill">+${dailyWords.toLocaleString()}</span>
            </div>`;
        historyList.appendChild(item);
    });
}

// =================================================================
// Reminder
// =================================================================

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
                    icon: 'assets/icon-128.png' // 确保路径正确
                });
            }
        }
    }
}