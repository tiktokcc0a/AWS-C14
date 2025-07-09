// ===================================================================================
// ### utils/network_watcher.js (V2.0 - 状态旗帜版) ###
// ===================================================================================
const axios = require('axios').default;

class NetworkWatcher {
    // 【核心修改】构造函数现在接收共享状态对象和实例ID
    constructor(sharedState, instanceId) {
        this.sharedState = sharedState;
        this.instanceId = instanceId;
        this.intervalId = null;
        this.isChecking = false;
        console.log(`[网络检测助手 ${this.instanceId}] 已初始化。`);
    }

    start() {
        if (this.intervalId) return;
        // 检查间隔可以按需调整
        this.intervalId = setInterval(() => this.checkNetwork(), 60000);
        console.log(`[网络检测助手 ${this.instanceId}] 已启动，每分钟检测一次网络。`);
    }

    stop() {
        if (this.intervalId) {
            clearInterval(this.intervalId);
            this.intervalId = null;
            console.log(`[网络检测助手 ${this.instanceId}] 已停止。`);
        }
    }

    async checkNetwork() {
        if (this.isChecking) return;
        this.isChecking = true;

        try {
            // 使用一个可靠的、低延迟的端点进行网络检查
            await axios.get('https://www.google.com/generate_204', { timeout: 15000 });
        } catch (error) {
            console.warn(`[网络检测助手 ${this.instanceId}] 检测到网络连接超时或失败: ${error.message}`);
            // 【核心修改】不再自己行动，而是升起“网络中断”的旗帜
            this.sharedState.networkInterrupted = true;
            console.log(`[网络检测助手 ${this.instanceId}] 已设置网络中断标志，等待主控处理。`);
            // 一旦检测到失败，可以暂时停止自身，避免在FIX期间重复报警
            this.stop(); 
        } finally {
            this.isChecking = false;
        }
    }
}

module.exports = { NetworkWatcher };