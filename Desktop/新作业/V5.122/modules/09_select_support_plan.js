// ===================================================================================
// ### modules/09_select_support_plan.js (V4 - 点击方式优化) ###
// ===================================================================================
const config = require('../shared/config');
const { humanLikeClick } = require('../shared/utils');

async function selectSupportPlan(page, data) {
    console.log("[模块9] 等待并选择支持计划...");
    
    // 【核心修改】将等待超时时间从180秒缩短为35秒
    await page.waitForSelector(config.SUPPORT_PLAN_SUBMIT_BUTTON, { visible: true, timeout: 35000 });
    
    // 【核心修改】将点击方式从humanLikeClick改为更直接的element.click
    console.log("[模块9] 准备使用 element.click() 方式直接点击 'Complete sign up' 按钮...");
    try {
        const buttonSelector = config.SUPPORT_PLAN_SUBMIT_BUTTON;
        // 等待按钮可见且可交互
        await page.waitForSelector(buttonSelector, { visible: true, timeout: 10000 }); 
        
        await page.evaluate((sel) => {
            // 由于::-p-text()是Puppeteer的伪类，不能在浏览器DOM中直接使用
            // 我们通过遍历所有按钮并检查其文本内容来找到目标按钮
            const buttons = Array.from(document.querySelectorAll('button'));
            // "Complete sign up" 是按钮的文本
            const targetButton = buttons.find(button => button.innerText.includes('Complete sign up'));
            
            if (targetButton) {
                targetButton.click();
            } else {
                throw new Error(`无法在DOM中找到文本为 "Complete sign up" 的按钮。`);
            }
        }, buttonSelector);
        
        console.log("[模块9 日志] 'Complete sign up' 按钮 element.click() 指令已成功发出。");

    } catch (error) {
        console.error(`[模块9 日志] 点击 'Complete sign up' 按钮时发生错误: ${error.message}`);
        // 将错误重新抛出，以便主控制器可以捕获并处理
        throw error;
    }
    
    console.log("[模块9] 支持计划选择完毕，注册完成。");
}

module.exports = { selectSupportPlan };