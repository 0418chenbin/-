// 全局变量
let nameList = []; // 抽奖名单
let winnerList = []; // 中奖记录
let isLotteryRunning = false; // 抽奖是否正在运行
let lotteryInterval = null; // 抽奖定时器
let currentIndex = 0; // 当前显示的名字索引

// DOM元素
const elements = {
    nameDisplay: document.getElementById('name-display'),
    statusDisplay: document.getElementById('status-display'),
    startBtn: document.getElementById('start-btn'),
    stopBtn: document.getElementById('stop-btn'),
    excelFile: document.getElementById('excel-file'),
    importBtn: document.getElementById('import-btn'),
    importStatus: document.getElementById('import-status'),
    nameCount: document.getElementById('name-count'),
    winnerList: document.getElementById('winner-list')
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
    
    isLotteryRunning = true;
    elements.startBtn.disabled = true;
    elements.stopBtn.disabled = false;
    elements.statusDisplay.textContent = '抽奖中...';
    
    // 添加滚动动画
    elements.nameDisplay.classList.add('lottery-spinning');
    
    // 开始滚动名字
    lotteryInterval = setInterval(() => {
        currentIndex = Math.floor(Math.random() * nameList.length);
        elements.nameDisplay.textContent = nameList[currentIndex];
    }, 100);
}

// 停止抽奖
function stopLottery() {
    if (!isLotteryRunning) return;
    
    // 清除定时器
    clearInterval(lotteryInterval);
    lotteryInterval = null;
    
    isLotteryRunning = false;
    elements.startBtn.disabled = false;
    elements.stopBtn.disabled = true;
    
    // 移除滚动动画
    elements.nameDisplay.classList.remove('lottery-spinning');
    
    // 添加中奖动画
    elements.nameDisplay.classList.add('winner-celebration');
    setTimeout(() => {
        elements.nameDisplay.classList.remove('winner-celebration');
    }, 500);
    
    // 获取中奖者
    const winnerName = nameList[currentIndex];
    elements.statusDisplay.textContent = '抽奖结束';
    
    // 记录中奖者
    addWinner(winnerName);
    
    // 从名单中移除中奖者（可选）
     nameList.splice(currentIndex, 1);
     elements.nameCount.textContent = `当前名单：${nameList.length} 人`;
     elements.statusDisplay.textContent = `就绪，共 ${nameList.length} 人`;
    
    // 如果名单为空，禁用开始按钮
    if (nameList.length === 0) {
        elements.startBtn.disabled = true;
        elements.statusDisplay.textContent = '名单已抽完';
    }
}

// 添加中奖者
function addWinner(name) {
    const now = new Date();
    const winner = {
        name: name,
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

// 初始化应用
init();
