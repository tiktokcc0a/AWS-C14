/**
 * 模块 4.5: 选择账户计划
 * ------------------------------------------------
 * @description
 * 此模块处理账户计划选择页面，并根据配置选择付费计划。
 *
 * @precondition
 * - 当前页面为账户计划选择页。
 * - 主控制器已传入有效的 page 对象和 finalConfig 配置。
 *
 * @postcondition
 * - 成功点击“Choose paid plan”按钮。
 * - 页面将自动跳转至下一个流程。
 *
 */

/**
 * 在账户计划页面选择付费计划。
 * @param {import('puppeteer').Page} page - Puppeteer 的页面实例。
 * @param {object} signupData - 包含注册流程中所有数据的对象。
 * @param {object} finalConfig - 合并了静态与动态配置的最终配置对象。
 */
async function selectFreePlan(page, signupData, finalConfig) {
    // 从 finalConfig 中直接获取此模块所需的CSS选择器对象
    const selectors = finalConfig.ACCOUNT_PLAN_SELECTORS;
    const instanceId = signupData.instanceId || 'W_unknown'; // 从注册数据中获取实例ID用于日志

    console.log(`[${instanceId} 模块 4.5] 目标：选择付费计划。`);

    try {
        // 【已修改】直接执行最终目标：等待并点击“Choose paid plan”按钮
        console.log(`[${instanceId} 模块 4.5] 正在等待并点击 'Choose paid plan' 按钮...`);
        
        await page.waitForSelector(selectors.choosePlanButton, { visible: true, timeout: 30000 });
        
        await page.click(selectors.choosePlanButton);

        console.log(`[${instanceId} 模块 4.5] 已成功点击 'Choose paid plan' 按钮，模块执行完毕。`);
        
    } catch (error) {
        // 如果在任何步骤中发生错误（例如，元素未在指定时间内找到）
        console.error(`[${instanceId} 模块 4.5] 执行失败: ${error.message}`);
        // 向上抛出错误，以便主控制器中的错误处理逻辑（如重试、截图等）可以捕获并处理它
        throw error;
    }
}

// 导出模块
module.exports = {
    selectFreePlan
};