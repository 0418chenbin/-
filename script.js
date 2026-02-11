// 全局变量
let nameList = []; // 抽奖名单
let winnerList = []; // 中奖记录
let isLotteryRunning = false; // 抽奖是否正在运行
let lotteryInterval = null; // 抽奖定时器
let currentIndex = 0; // 当前显示的名字索引
let currentAward = ''; // 当前选中的奖项
let specialWinner = ''; // 固定特等奖获得者

// DOM元素
const elements = {
    nameDisplay: document.getElementById('name-display'),
    statusDisplay: document.getElementById('status-display'),
    awardDisplay: document.getElementById('award-display'),
    startBtn: document.getElementById('start-btn'),
    stopBtn: document.getElementById('stop-btn'),
    excelFile: document.getElementById('excel-file'),
    importBtn: document.getElementById('import-btn'),
    exportBtn: document.getElementById('export-btn'),
    importStatus: document.getElementById('import-status'),
    nameCount: document.getElementById('name-count'),
    winnerList: document.getElementById('winner-list'),
    specialWinner: document.getElementById('special-winner'),
    specialWinnerLabel: document.getElementById('special-winner-label'),
    specialWinnerContainer: document.getElementById('special-winner-container'),
    setSpecialBtn: document.getElementById('set-special-btn')
};

// 初始化
function init() {
    // 禁用开始按钮
    elements.startBtn.disabled = true;
    elements.stopBtn.disabled = true;
    
    // 绑定事件
    bindEvents();
    
    // 加载本地存储的中奖记录
    loadWinnerList();
    
    // 更新中奖记录显示
    updateWinnerListDisplay();
}

// 绑定事件
function bindEvents() {
    // 开始抽奖按钮
    elements.startBtn.addEventListener('click', startLottery);
    
    // 停止抽奖按钮
    elements.stopBtn.addEventListener('click', stopLottery);
    
    // 导入Excel按钮
    elements.importBtn.addEventListener('click', importExcel);
    
    // 导出记录按钮
    elements.exportBtn.addEventListener('click', exportWinners);
    
    // 设置固定特等奖按钮
    elements.setSpecialBtn.addEventListener('click', setSpecialWinner);
    
    // 奖项按钮
    const awardBtns = document.querySelectorAll('.award-btn');
    awardBtns.forEach(btn => {
        btn.addEventListener('click', function() {
            selectAward(this.dataset.award);
        });
    });
}

// 选择奖项
function selectAward(award) {
    currentAward = award;
    elements.awardDisplay.textContent = `当前奖项：${award}`;
    
    // 更新奖项按钮样式
    const awardBtns = document.querySelectorAll('.award-btn');
    awardBtns.forEach(btn => {
        if (btn.dataset.award === award) {
            btn.classList.add('ring-2', 'ring-yellow-300', 'scale-105');
        } else {
            btn.classList.remove('ring-2', 'ring-yellow-300', 'scale-105');
        }
    });
    
    // 启用开始按钮
    if (nameList.length > 0) {
        elements.startBtn.disabled = false;
    }
}

// 设置固定特等奖
function setSpecialWinner() {
    const name = elements.specialWinner.value.trim();
    if (name) {
        specialWinner = name;
        alert(`固定特等奖已设置为：${name}`);
        
        // 隐藏整个固定特等奖设置容器
        elements.specialWinnerContainer.style.display = 'none';
    } else {
        alert('请输入固定特等奖获得者姓名');
    }
}

// 导入Excel文件
function importExcel() {
    const file = elements.excelFile.files[0];
    
    if (!file) {
        showImportStatus('请选择Excel文件', 'error');
        return;
    }
    
    if (!file.name.endsWith('.xlsx') && !file.name.endsWith('.xls')) {
        showImportStatus('请选择Excel文件(.xlsx或.xls)', 'error');
        return;
    }
    
    showImportStatus('正在导入...', 'info');
    
    const reader = new FileReader();
    
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            
            // 获取第一个工作表
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            
            // 转换为JSON
            const jsonData = XLSX.utils.sheet_to_json(worksheet);
            
            // 提取名字
            nameList = [];
            jsonData.forEach(row => {
                // 尝试更多列名
                const name = row['姓名'] || row['名字'] || row['name'] || row['Name'] || row['人员'] || row['员工'] || Object.values(row)[0];
                if (name) {
                    nameList.push(name.toString().trim());
                }
            });
            
            // 去重
            //nameList = [...new Set(nameList)];
            
            if (nameList.length === 0) {
                showImportStatus('未找到有效名字', 'error');
                elements.statusDisplay.textContent = '等待导入名单';
                elements.startBtn.disabled = true;
            } else {
                showImportStatus(`导入成功，共 ${nameList.length} 人`, 'success');
                elements.nameCount.textContent = `当前名单：${nameList.length} 人`;
                elements.statusDisplay.textContent = `就绪，共 ${nameList.length} 人`;
                elements.startBtn.disabled = false;
            }
            
        } catch (error) {
            console.error('导入失败:', error);
            showImportStatus('导入失败，请检查Excel文件格式', 'error');
        }
    };
    
    reader.onerror = function() {
        showImportStatus('文件读取失败', 'error');
    };
    
    reader.readAsArrayBuffer(file);
}

// 显示导入状态
function showImportStatus(message, type) {
    elements.importStatus.textContent = message;
    elements.importStatus.className = '';
    if (type === 'success' || type === 'error') {
        elements.importStatus.classList.add(type);
    }
    
    // 3秒后清除状态
    setTimeout(() => {
        elements.importStatus.textContent = '';
        elements.importStatus.className = '';
    }, 3000);
}

// 开始抽奖
function startLottery() {
    if (nameList.length === 0) {
        elements.statusDisplay.textContent = '请先导入名单';
        return;
    }
    
    if (!currentAward) {
        elements.statusDisplay.textContent = '请先选择奖项';
        return;
    }
    
    isLotteryRunning = true;
    elements.startBtn.disabled = true;
    elements.stopBtn.disabled = false;
    elements.statusDisplay.textContent = '抽奖中...';
    
    // 添加滚动动画
    elements.nameDisplay.classList.add('lottery-spinning');
    
    // 开始滚动名字
    lotteryInterval = setInterval(() => {
        if (currentAward === '特等奖' && specialWinner) {
            // 特等奖滚动，减少固定获奖者出现的频率，增加悬念
            if (Math.random() > 0.5) {
                elements.nameDisplay.textContent = specialWinner;
            } else {
                // 经常显示其他名字，增加悬念
                currentIndex = Math.floor(Math.random() * nameList.length);
                elements.nameDisplay.textContent = nameList[currentIndex];
            }
        } else {
            // 其他奖项正常滚动
            currentIndex = Math.floor(Math.random() * nameList.length);
            elements.nameDisplay.textContent = nameList[currentIndex];
        }
    }, 100);
}

// 停止抽奖
function stopLottery() {
    // 清除定时器
    if (lotteryInterval) {
        clearInterval(lotteryInterval);
        lotteryInterval = null;
    }
    
    isLotteryRunning = false;
    elements.startBtn.disabled = false;
    elements.stopBtn.disabled = true;
    
    // 移除滚动动画
    elements.nameDisplay.classList.remove('lottery-spinning');
    
    // 确保显示固定特等奖获得者
    let winnerName;
    if (currentAward === '特等奖' && specialWinner) {
        winnerName = specialWinner;
        elements.nameDisplay.textContent = winnerName;
    } else {
        winnerName = nameList[currentIndex];
    }
    
    // 添加中奖动画
    elements.nameDisplay.classList.add('winner-celebration');
    setTimeout(() => {
        elements.nameDisplay.classList.remove('winner-celebration');
    }, 500);
    
    elements.statusDisplay.textContent = '抽奖结束';
    
    // 记录中奖者
    addWinner(winnerName, currentAward);
    
    // 从名单中移除中奖者（固定特等奖除外）
    if (currentAward !== '特等奖' || !specialWinner) {
        nameList.splice(currentIndex, 1);
        elements.nameCount.textContent = `当前名单：${nameList.length} 人`;
        elements.statusDisplay.textContent = `就绪，共 ${nameList.length} 人`;
    }
    
    // 如果名单为空，禁用开始按钮
    if (nameList.length === 0) {
        elements.startBtn.disabled = true;
        elements.statusDisplay.textContent = '名单已抽完';
    }
}

// 添加中奖者
function addWinner(name, award) {
    const now = new Date();
    const winner = {
        name: name,
        award: award,
        time: now.toLocaleString(),
        id: Date.now()
    };
    
    // 添加到中奖记录
    winnerList.unshift(winner);
    
    // 保存到本地存储
    saveWinnerList();
    
    // 更新中奖记录显示
    updateWinnerListDisplay();
}

// 保存中奖记录到本地存储
function saveWinnerList() {
    localStorage.setItem('lotteryWinners', JSON.stringify(winnerList));
}

// 从本地存储加载中奖记录
function loadWinnerList() {
    const saved = localStorage.getItem('lotteryWinners');
    if (saved) {
        try {
            winnerList = JSON.parse(saved);
            // 为没有id的记录添加id
            winnerList.forEach(winner => {
                if (!winner.id) {
                    winner.id = Date.now() + Math.random();
                }
            });
            // 保存更新后的记录
            saveWinnerList();
        } catch (error) {
            console.error('加载中奖记录失败:', error);
            winnerList = [];
        }
    }
}

// 更新中奖记录显示
function updateWinnerListDisplay() {
    if (winnerList.length === 0) {
        elements.winnerList.innerHTML = '<div class="text-center text-gray-400 py-8">暂无中奖记录</div>';
        return;
    }
    
    elements.winnerList.innerHTML = winnerList.map((winner, index) => `
        <div class="winner-item">
            <div class="flex items-center">
                <div class="winner-number">${index + 1}</div>
                <div class="winner-name">${winner.name}</div>
                <div class="winner-award ml-4 px-2 py-1 rounded-full text-sm" style="background-color: ${getAwardColor(winner.award)}">
                    ${winner.award}
                </div>
            </div>
            <div class="flex items-center gap-4">
                <div class="winner-time">${winner.time}</div>
                <button class="delete-btn text-red-400 hover:text-red-600 transition-colors" data-id="${winner.id}">
                    <i class="fa fa-trash"></i>
                </button>
            </div>
        </div>
    `).join('');
    
    // 绑定删除按钮事件
    bindDeleteEvents();
}

// 获取奖项对应的颜色
function getAwardColor(award) {
    switch (award) {
        case '特等奖': return 'rgba(147, 51, 234, 0.8)'; // 紫色
        case '一等奖': return 'rgba(220, 38, 38, 0.8)'; // 红色
        case '二等奖': return 'rgba(249, 115, 22, 0.8)'; // 橙色
        case '三等奖': return 'rgba(34, 197, 94, 0.8)'; // 绿色
        default: return 'rgba(107, 114, 128, 0.8)'; // 灰色
    }
}

// 绑定删除按钮事件
function bindDeleteEvents() {
    const deleteBtns = document.querySelectorAll('.delete-btn');
    deleteBtns.forEach(btn => {
        btn.addEventListener('click', function() {
            const winnerId = parseInt(this.getAttribute('data-id'));
            deleteWinner(winnerId);
        });
    });
}

// 删除中奖者
function deleteWinner(id) {
    // 从中奖记录中移除
    winnerList = winnerList.filter(winner => winner.id !== id);
    
    // 保存到本地存储
    saveWinnerList();
    
    // 更新中奖记录显示
    updateWinnerListDisplay();
}

// 导出中奖记录
function exportWinners() {
    if (winnerList.length === 0) {
        alert('暂无中奖记录');
        return;
    }
    
    // 准备导出数据
    const exportData = winnerList.map((winner, index) => ({
        '序号': index + 1,
        '姓名': winner.name,
        '奖项': winner.award,
        '中奖时间': winner.time
    }));
    
    // 创建工作簿
    const wb = XLSX.utils.book_new();
    
    // 创建工作表
    const ws = XLSX.utils.json_to_sheet(exportData);
    
    // 添加工作表到工作簿
    XLSX.utils.book_append_sheet(wb, ws, '中奖记录');
    
    // 生成文件名
    const now = new Date();
    const fileName = `中奖记录_${now.getFullYear()}${(now.getMonth() + 1).toString().padStart(2, '0')}${now.getDate().toString().padStart(2, '0')}.xlsx`;
    
    // 导出文件
    XLSX.writeFile(wb, fileName);
}

// 初始化应用
init();
