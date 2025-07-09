// ===================================================================================
// ### modules/9.5_handle_confirmation.js (V4 - 增加“未就绪”失败原因) ###
// ===================================================================================
const config = require('../shared/config');

async function handleConfirmation(page) {
    console.log("[模块9.5] 已进入注册确认(confirmation)页面。");
    
    await new Promise(resolve => setTimeout(resolve, 3000));
    
    const currentUrl = page.url();
    if (currentUrl.includes('portal.aws.amazon.com/billing/signup/incomplete')) {
        // --- 修改点: 在日志中附上明确的失败原因 ---
        console.error("[模块9.5] 判定为注册失败！URL包含 'incomplete'。附上失败原因:未就绪");
        // ------------------------------------
        throw new Error("REGISTRATION_FAILED_INCOMPLETE");
    }

    const waitTime = 50;
    console.log(`[模块9.5] 开始等待 ${waitTime} 秒...`);
    await new Promise(resolve => setTimeout(resolve, waitTime * 1000));
    
    console.log(`[模块9.5] 等待结束，准备跳转至IAM密钥创建页面...`);
    await page.goto(config.AWS_IAM_WIZARD_URL, { 
        waitUntil: 'load', 
        timeout: 180000 
    });
    console.log("[模块9.5] 已成功发起导航至IAM页面。");
}

module.exports = { handleConfirmation };