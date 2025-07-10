// ===================================================================================
// ### modules/07_enter_phone_number.js (V2.0 - 动态选择器版) ###
// ===================================================================================
const { humanLikeType, humanLikeClick } = require('../shared/utils');

async function enterPhoneNumber(page, data, config) { // 【修改】接收配置对象
    console.log("[模块7] 准备进入手机验证页面...");
    console.log("[模块7] 正在刷新页面以激活...");
    await page.reload({ waitUntil: 'networkidle0' });

    try {
        console.log("[模块7] 正在检测是否存在 'Ext' 分区失败标志...");
        const extLabelSelector = 'label::-p-text(Ext)';
        const extLabel = await page.$(extLabelSelector);
        if (extLabel) {
            console.error("[模块7] 检测到 'Ext' 标签，判定为注册失败！");
            throw new Error("出现分区");
        }
        console.log("[模块7] 未发现 'Ext' 分区，流程继续。");
    } catch (error) {
        if (error.message === "出现分区") {
            throw error;
        }
        console.warn(`[模块7] 在检测 'Ext' 标签时发生意外错误: ${error.message}`);
    }
    
    try {
        const cookieAcceptButtonSelector = 'button[data-id="awsccc-cb-btn-accept"]';
        console.log('[模块7] 正在检查Cookie横幅...');
        const acceptButton = await page.waitForSelector(cookieAcceptButtonSelector, { timeout: 5000 });
        if (acceptButton) {
            await acceptButton.click();
            console.log('[模块7] Cookie横幅已点击“Accept”。');
            await new Promise(resolve => setTimeout(resolve, 1000));
        }
    } catch (error) {
        console.log('[模块7] 未发现Cookie横幅，或已处理，继续执行...');
    }
    
    console.log("[模块7] 等待并填写身份验证手机号...");
    await page.waitForSelector(config.IDENTITY_PHONE_NUMBER_SELECTOR, { visible: true, timeout: 180000 });

    // 【核心修改】使用动态生成的国家/地区选择器
    await page.click(config.IDENTITY_PHONE_COUNTRY_TRIGGER_SELECTOR);
    await page.waitForSelector(config.dynamicIdentityPhoneOptionSelector, { visible: true });
    await page.click(config.dynamicIdentityPhoneOptionSelector);

    await humanLikeType(page, config.IDENTITY_PHONE_NUMBER_SELECTOR, data.phone_number);
    
    console.log("[模块7] 准备点击“发送短信”按钮...");
    await humanLikeClick(page, config.IDENTITY_SEND_SMS_BUTTON_SELECTOR);
    
    console.log("[模块7] 手机号提交完毕，主控制器将接管后续流程。");
}

module.exports = { enterPhoneNumber };