// ===================================================================================
// ### test_delete_browser.js (用于独立测试关闭与删除功能) ###
// ===================================================================================
const { setupBrowser, tearDownBrowser } = require('./shared/browser_setup');

async function runDeleteTest() {
    console.log("--- 开始独立测试：浏览器关闭与删除功能 ---");
    let browserId = null;

    try {
        // --- 步骤 1: 创建并打开一个测试浏览器 ---
        console.log("\n[测试步骤 1] 正在创建一个用于测试的浏览器实例...");
        const instanceId = 'DELETE-TEST'; // 使用一个特殊的ID，便于识别
        const proxyPort = 45999; // 使用一个不常用的端口，避免冲突
        const browserIndex = 999;
        
        // 我们只需要 browserId，所以解构时只取它
        ({ browserId } = await setupBrowser(instanceId, false, proxyPort, browserIndex));
        
        if (!browserId) {
            throw new Error("创建测试浏览器失败，无法获取 browserId。");
        }
        
        console.log(`[测试步骤 1] ✅ 成功创建并打开了测试浏览器，ID: ${browserId}`);
        console.log("浏览器窗口将保持打开5秒，然后自动开始清理流程...");
        await new Promise(resolve => setTimeout(resolve, 5000));

        // --- 步骤 2: 调用我们修复后的清理函数 ---
        console.log("\n[测试步骤 2] 开始调用 tearDownBrowser 函数进行清理...");
        await tearDownBrowser(browserId);
        
        console.log("\n[测试总结] ✅ tearDownBrowser 函数执行完毕，未抛出致命错误。");
        console.log("请检查比特浏览器后台，确认该窗口已被删除。");

    } catch (error) {
        console.error("\n[测试总结] ❌ 测试过程中发生严重错误！");
        console.error("错误详情:", error.message);
        console.error("请检查日志以确定问题所在。");
    }
}

// 运行测试
runDeleteTest();