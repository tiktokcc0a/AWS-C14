// ===================================================================================
// ### cleanup_manager.js (V1.0) ###
// 新增的批量清理脚本，用于高效地关闭和删除所有受控的浏览器窗口。
// ===================================================================================
const fs = require('fs').promises;
const path = require('path');
const axios = require('axios').default;

const BIT_API_BASE_URL = 'http://127.0.0.1:54345';
const BROWSER_ID_TRACKING_FILE = path.join(__dirname, 'data', 'managed_browser_ids.json');

/**
 * 从追踪文件中读取所有受控的浏览器ID。
 * @returns {Promise<string[]>} ID数组
 */
async function getManagedBrowserIds() {
    try {
        const content = await fs.readFile(BROWSER_ID_TRACKING_FILE, 'utf-8');
        const ids = JSON.parse(content);
        if (!Array.isArray(ids)) {
            throw new Error('追踪文件格式无效，期望是一个JSON数组。');
        }
        return ids;
    } catch (error) {
        if (error.code === 'ENOENT') {
            console.log('[管理器] 未找到浏览器ID追踪文件，无需清理。');
            return [];
        }
        console.error(`[管理器] 读取浏览器ID追踪文件时出错: ${error.message}`);
        return [];
    }
}

/**
 * 调用Bit浏览器API关闭所有窗口。
 */
async function closeAllBrowsers() {
    console.log('[管理器] 步骤 1/3: 发送命令关闭所有浏览器窗口...');
    try {
        const response = await axios.post(`${BIT_API_BASE_URL}/browser/close/all`);
        if (response.data.success) {
            console.log('[管理器] "关闭所有窗口" 命令已成功发送。等待进程退出...');
            // 等待几秒钟让进程有时间关闭
            await new Promise(resolve => setTimeout(resolve, 5000));
            return true;
        } else {
            console.error('[管理器] "关闭所有窗口" API调用失败:', response.data);
            return false;
        }
    } catch (error) {
        console.error(`[管理器] 调用 "关闭所有窗口" API时发生网络错误: ${error.message}`);
        return false;
    }
}

/**
 * 批量删除浏览器窗口配置。
 * @param {string[]} idsToDelete 要删除的ID数组
 */
async function bulkDeleteBrowsers(idsToDelete) {
    if (idsToDelete.length === 0) {
        console.log('[管理器] 步骤 2/3: 没有需要删除的浏览器ID，跳过删除。');
        return;
    }
    console.log(`[管理器] 步骤 2/3: 准备批量删除 ${idsToDelete.length} 个浏览器配置...`);

    const CHUNK_SIZE = 100; // API限制每次最多100个
    let successCount = 0;
    let failCount = 0;

    for (let i = 0; i < idsToDelete.length; i += CHUNK_SIZE) {
        const chunk = idsToDelete.slice(i, i + CHUNK_SIZE);
        console.log(`[管理器] 正在删除批次 ${Math.floor(i / CHUNK_SIZE) + 1} (数量: ${chunk.length})...`);
        try {
            const response = await axios.post(`${BIT_API_BASE_URL}/browser/delete/ids`, { ids: chunk });
            if (response.data.success) {
                console.log(`[管理器] 批次删除成功。`);
                successCount += chunk.length;
            } else {
                console.error(`[管理器] 批次删除API调用失败:`, response.data);
                failCount += chunk.length;
            }
        } catch (error) {
            console.error(`[管理器] 批次删除时发生网络错误: ${error.message}`);
            failCount += chunk.length;
        }
    }
    console.log(`[管理器] 删除操作完成。成功: ${successCount}, 失败: ${failCount}。`);
}

/**
 * 清空浏览器ID追踪文件。
 */
async function clearTrackingFile() {
    console.log('[管理器] 步骤 3/3: 清理浏览器ID追踪文件...');
    try {
        await fs.writeFile(BROWSER_ID_TRACKING_FILE, '[]');
        console.log('[管理器] 追踪文件已清空。');
    } catch (error) {
        console.error(`[管理器] 清理追踪文件时出错: ${error.message}`);
    }
}

/**
 * 主执行函数
 */
async function main() {
    console.log('--- 开始执行批量浏览器清理任务 ---');

    // 1. 获取所有需要管理的ID
    const allIds = await getManagedBrowserIds();
    if (allIds.length === 0) {
        console.log('--- 清理任务完成（无事可做） ---');
        return;
    }

    // 2. 关闭所有窗口
    await closeAllBrowsers();

    // 3. 批量删除所有受控的窗口
    await bulkDeleteBrowsers(allIds);

    // 4. 清理追踪文件
    await clearTrackingFile();

    console.log('--- 所有清理操作已执行完毕 ---');
}

main().catch(error => {
    console.error('清理脚本执行期间发生致命错误:', error);
    process.exit(1);
});