/* global document, Office, Word */

// 添加错误处理
window.addEventListener('error', (event) => {
    console.error('Global error:', event.error);
    event.preventDefault();
});

window.addEventListener('unhandledrejection', (event) => {
    console.error('Unhandled promise rejection:', event.reason);
    event.preventDefault();
});

Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        try {
            // 隐藏加载消息，显示应用主体
            const sideloadMsg = document.getElementById("sideload-msg");
            const appBody = document.getElementById("app-body");
            
            if (sideloadMsg) sideloadMsg.style.display = "none";
            if (appBody) appBody.style.display = "flex";
            
            // 绑定事件
            const savePlanBtn = document.getElementById("savePlan");
            const updateProgressBtn = document.getElementById("updateProgress");
            
            if (savePlanBtn) savePlanBtn.onclick = savePlan;
            if (updateProgressBtn) updateProgressBtn.onclick = updateProgress;
            
            // 加载已保存的数据
            loadSavedData();
            
            // 延迟初始化图表，确保 DOM 完全加载
            setTimeout(() => {
                if (typeof Chart !== 'undefined') {
                    initChart();
                } else {
                    console.warn('Chart.js not loaded, skipping chart initialization');
                }
            }, 100);
            
            // 设置定时器检查提醒
            setInterval(checkReminder, 60000); // 每分钟检查一次
            
        } catch (error) {
            console.error("Initialization error:", error);
        }
    }
});

// 保存写作计划
async function savePlan() {
    try {
        const plan = {
            projectName: document.getElementById("projectName").value,
            targetWords: parseInt(document.getElementById("targetWords").value) || 0,
            deadline: document.getElementById("deadline").value,
            dailyTarget: parseInt(document.getElementById("dailyTarget").value) || 0,
            reminderTime: document.getElementById("reminderTime").value,
            startDate: new Date().toISOString().split('T')[0],
            progress: []
        };
        
        // 验证输入
        if (!plan.projectName || !plan.targetWords || !plan.deadline) {
            showMessage('请填写所有必填项！', 'warning');
            return;
        }
        
        // 保存到本地存储
        localStorage.setItem('writingPlan', JSON.stringify(plan));
        
        // 显示成功消息
        showMessage('计划已保存！', 'success');
        
        // 更新显示
        updateDisplay();
    } catch (error) {
        console.error('Save plan error:', error);
        showMessage('保存计划时出错：' + error.message, 'danger');
    }
}

// 更新进度
async function updateProgress() {
    try {
        await Word.run(async (context) => {
            // 获取文档正文
            const body = context.document.body;
            context.load(body, 'text');
            
            await context.sync();
            
            // 计算字数
            const text = body.text;
            const wordCount = countWords(text);
            
            // 更新进度数据
            const plan = JSON.parse(localStorage.getItem('writingPlan') || '{}');
            if (!plan.progress) plan.progress = [];
            
            const today = new Date().toISOString().split('T')[0];
            const todayProgress = plan.progress.find(p => p.date === today);
            
            if (todayProgress) {
                todayProgress.words = wordCount;
            } else {
                plan.progress.push({
                    date: today,
                    words: wordCount
                });
            }
            
            localStorage.setItem('writingPlan', JSON.stringify(plan));
            
            // 更新显示
            updateDisplay();
            updateChart();
            
            showMessage(`进度已更新！当前字数：${wordCount}`, 'success');
        });
    } catch (error) {
        console.error('Update progress error:', error);
        showMessage('更新进度时出错：' + error.message, 'danger');
    }
}

// 计算字数
function countWords(text) {
    if (!text) return 0;
    
    // 移除多余空格和换行
    text = text.replace(/\s+/g, ' ').trim();
    
    // 计算中文字符
    const chineseChars = (text.match(/[\u4e00-\u9fa5]/g) || []).length;
    
    // 计算英文单词
    const englishWords = text.replace(/[\u4e00-\u9fa5]/g, ' ')
        .split(/\s+/)
        .filter(word => word.length > 0).length;
    
    return chineseChars + englishWords;
}

// 加载已保存的数据
function loadSavedData() {
    try {
        const plan = JSON.parse(localStorage.getItem('writingPlan') || '{}');
        
        if (plan.projectName) {
            const projectNameEl = document.getElementById("projectName");
            const targetWordsEl = document.getElementById("targetWords");
            const deadlineEl = document.getElementById("deadline");
            const dailyTargetEl = document.getElementById("dailyTarget");
            const reminderTimeEl = document.getElementById("reminderTime");
            
            if (projectNameEl) projectNameEl.value = plan.projectName;
            if (targetWordsEl) targetWordsEl.value = plan.targetWords;
            if (deadlineEl) deadlineEl.value = plan.deadline;
            if (dailyTargetEl) dailyTargetEl.value = plan.dailyTarget;
            if (reminderTimeEl) reminderTimeEl.value = plan.reminderTime;
            
            updateDisplay();
        }
    } catch (error) {
        console.error('Load saved data error:', error);
    }
}

// 更新显示
function updateDisplay() {
    try {
        const plan = JSON.parse(localStorage.getItem('writingPlan') || '{}');
        
        if (!plan.targetWords) return;
        
        // 计算当前总字数
        const currentWords = getCurrentTotalWords(plan);
        const progress = Math.min(100, (currentWords / plan.targetWords * 100)).toFixed(1);
        
        // 更新进度条
        const progressBar = document.getElementById("progressBar");
        if (progressBar) {
            progressBar.style.width = progress + '%';
            progressBar.textContent = progress + '%';
        }
        
        // 更新统计数据
        const currentWordsEl = document.getElementById("currentWords");
        const targetWordsDisplayEl = document.getElementById("targetWordsDisplay");
        const daysLeftEl = document.getElementById("daysLeft");
        const todayWordsEl = document.getElementById("todayWords");
        
        if (currentWordsEl) currentWordsEl.textContent = currentWords.toLocaleString();
        if (targetWordsDisplayEl) targetWordsDisplayEl.textContent = plan.targetWords.toLocaleString();
        
        // 计算剩余天数
        if (daysLeftEl && plan.deadline) {
            const today = new Date();
            const deadline = new Date(plan.deadline);
            const daysLeft = Math.ceil((deadline - today) / (1000 * 60 * 60 * 24));
            daysLeftEl.textContent = daysLeft > 0 ? daysLeft : 0;
        }
        
        // 计算今日字数
        if (todayWordsEl) {
            const todayWords = getTodayWords(plan);
            todayWordsEl.textContent = todayWords.toLocaleString();
        }
        
        // 更新历史记录
        updateHistory(plan);
    } catch (error) {
        console.error('Update display error:', error);
    }
}

// 获取当前总字数
function getCurrentTotalWords(plan) {
    if (!plan.progress || plan.progress.length === 0) return 0;
    
    // 返回最新的字数记录
    const latestProgress = plan.progress.reduce((latest, current) => {
        return new Date(current.date) > new Date(latest.date) ? current : latest;
    });
    
    return latestProgress.words || 0;
}

// 获取今日字数
function getTodayWords(plan) {
    if (!plan.progress) return 0;
    
    const today = new Date().toISOString().split('T')[0];
    const todayProgress = plan.progress.find(p => p.date === today);
    const yesterdayProgress = plan.progress.find(p => {
        const yesterday = new Date();
        yesterday.setDate(yesterday.getDate() - 1);
        return p.date === yesterday.toISOString().split('T')[0];
    });
    
    if (!todayProgress) return 0;
    
    const yesterdayWords = yesterdayProgress ? yesterdayProgress.words : 0;
    return Math.max(0, todayProgress.words - yesterdayWords);
}

// 更新历史记录
function updateHistory(plan) {
    try {
        const historyList = document.getElementById("historyList");
        if (!historyList) return;
        
        historyList.innerHTML = '';
        
        if (!plan.progress || plan.progress.length === 0) {
            historyList.innerHTML = '<div class="list-group-item">暂无记录</div>';
            return;
        }
        
        // 按日期降序排序
        const sortedProgress = [...plan.progress].sort((a, b) => 
            new Date(b.date) - new Date(a.date)
        );
        
        // 显示最近10条记录
        sortedProgress.slice(0, 10).forEach((record, index) => {
            const prevRecord = sortedProgress[index + 1];
            const dailyWords = prevRecord ? record.words - prevRecord.words : record.words;
            
            const item = document.createElement('div');
            item.className = 'list-group-item';
            item.innerHTML = `
                <div class="d-flex justify-content-between align-items-center">
                    <div>
                        <strong>${formatDate(record.date)}</strong>
                        <br>
                        <small>总字数: ${record.words.toLocaleString()}</small>
                    </div>
                    <div class="text-right">
                        <span class="badge badge-primary badge-pill">
                            +${Math.max(0, dailyWords).toLocaleString()}
                        </span>
                    </div>
                </div>
            `;
            historyList.appendChild(item);
        });
    } catch (error) {
        console.error('Update history error:', error);
    }
}

// 格式化日期
function formatDate(dateString) {
    try {
        const date = new Date(dateString);
        return `${date.getMonth() + 1}月${date.getDate()}日`;
    } catch (error) {
        return dateString;
    }
}

// 初始化图表
let progressChart;
function initChart() {
    try {
        const canvas = document.getElementById('progressChart');
        if (!canvas) {
            console.warn('Chart canvas not found');
            return;
        }
        
        const ctx = canvas.getContext('2d');
        
        // 检查 Chart.js 是否已加载
        if (typeof Chart === 'undefined') {
            console.warn('Chart.js is not loaded');
            return;
        }
        
        progressChart = new Chart(ctx, {
            type: 'line',
            data: {
                labels: [],
                datasets: [{
                    label: '累计字数',
                    data: [],
                    borderColor: '#2b579a',
                    backgroundColor: 'rgba(43, 87, 154, 0.1)',
                    tension: 0.1
                }, {
                    label: '每日字数',
                    data: [],
                    borderColor: '#28a745',
                    backgroundColor: 'rgba(40, 167, 69, 0.1)',
                    tension: 0.1,
                    yAxisID: 'y1'
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                scales: {
                    y: {
                        beginAtZero: true,
                        position: 'left'
                    },
                    y1: {
                        beginAtZero: true,
                        position: 'right',
                        grid: {
                            drawOnChartArea: false
                        }
                    }
                }
            }
        });
        
        updateChart();
    } catch (error) {
        console.error('Init chart error:', error);
    }
}

// 更新图表
function updateChart() {
    try {
        if (!progressChart) return;
        
        const plan = JSON.parse(localStorage.getItem('writingPlan') || '{}');
        if (!plan.progress || plan.progress.length === 0) return;
        
        // 按日期排序
        const sortedProgress = [...plan.progress].sort((a, b) => 
            new Date(a.date) - new Date(b.date)
        );
        
        const labels = [];
        const totalWords = [];
        const dailyWords = [];
        
        sortedProgress.forEach((record, index) => {
            labels.push(formatDate(record.date));
            totalWords.push(record.words);
            
            const prevRecord = index > 0 ? sortedProgress[index - 1] : null;
            const daily = prevRecord ? record.words - prevRecord.words : record.words;
            dailyWords.push(Math.max(0, daily));
        });
        
        progressChart.data.labels = labels;
        progressChart.data.datasets[0].data = totalWords;
        progressChart.data.datasets[1].data = dailyWords;
        progressChart.update();
    } catch (error) {
        console.error('Update chart error:', error);
    }
}

// 检查提醒
function checkReminder() {
    try {
        const plan = JSON.parse(localStorage.getItem('writingPlan') || '{}');
        if (!plan.reminderTime) return;
        
        const now = new Date();
        const currentTime = `${now.getHours().toString().padStart(2, '0')}:${now.getMinutes().toString().padStart(2, '0')}`;
        
        if (currentTime === plan.reminderTime) {
            const todayWords = getTodayWords(plan);
            const remaining = plan.dailyTarget - todayWords;
            
            if (remaining > 0) {
                showMessage(`写作提醒：今天还需要写 ${remaining} 字才能完成目标！`, 'warning');
                
                // 如果浏览器支持通知
                if ("Notification" in window && Notification.permission === "granted") {
                    new Notification("写作提醒", {
                        body: `今天还需要写 ${remaining} 字才能完成目标！`,
                        icon: '/assets/icon-128.png'
                    });
                }
            }
        }
    } catch (error) {
        console.error('Check reminder error:', error);
    }
}

// 显示消息
function showMessage(message, type) {
    try {
        const alertDiv = document.createElement('div');
        alertDiv.className = `alert alert-${type} alert-dismissible fade show`;
        alertDiv.innerHTML = `
            ${message}
            <button type="button" class="close" onclick="this.parentElement.remove()">
                <span>&times;</span>
            </button>
        `;
        
        const container = document.querySelector('.container-fluid');
        const navTabs = document.querySelector('.nav-tabs');
        
        if (container && navTabs) {
            container.insertBefore(alertDiv, navTabs);
        } else if (container) {
            container.insertBefore(alertDiv, container.firstChild);
        }
        
        // 3秒后自动消失
        setTimeout(() => {
            alertDiv.remove();
        }, 3000);
    } catch (error) {
        console.error('Show message error:', error);
    }
}

// 请求通知权限
if ("Notification" in window && Notification.permission === "default") {
    Notification.requestPermission().catch(err => console.log('Notification permission error:', err));
}
