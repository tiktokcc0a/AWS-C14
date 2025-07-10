// ===================================================================================
// ### modules/03_verify_email.js (V2.0 - 集成Outlook API) ###
// ===================================================================================
const axios = require('axios').default;
const { humanLikeType, humanLikeClick } = require('../shared/utils');

/**
 * 动态获取邮箱验证码，支持多个API供应商
 * @param {object} page - Puppeteer page object.
 * @param {object} data - 包含所有任务信息的对象，如邮箱、密码、API供应商等
 * @param {object} config - 全局配置对象
 * @returns {Promise<void>}
 */
async function verifyEmail(page, data, config) {
    const emailAddress = data.account;
    const emailPassword = data.password;
    // 从传入的任务数据中获取API供应商，如果没有则默认为'Domain'
    const apiProvider = data.email_api_provider || 'Domain'; 
    const apiConfig = config.EMAIL_API_CONFIG[apiProvider];

    if (!apiConfig) {
        throw new Error(`[模块3] 错误: 邮箱API供应商 '${apiProvider}' 未在config.js中配置。`);
    }

    console.log(`[模块3] [${emailAddress}] 等待并填写邮箱验证码... (使用 ${apiProvider} 接口)`);
    await page.waitForSelector(config.OTP_INPUT_SELECTOR, { visible: true, timeout: 180000 });

    let verificationCode = null;

    // 循环获取验证码，总共尝试40次 (约200秒)
    for (let attempts = 0; attempts < 40; attempts++) {
        try {
            console.log(`[模块3] [${emailAddress}] 正在尝试获取验证码 (第 ${attempts + 1}/40 次)`);
            
            let response;
            // --- 动态API调用逻辑 ---
            if (apiProvider === 'Domain') {
                const url = `${apiConfig.url}?email=${emailAddress}&password=${emailPassword}&email_provider=domain`;
                response = await axios.get(url, { timeout: 20000 });

                if (response.data && response.data.verification_code) {
                    verificationCode = response.data.verification_code;
                } else {
                    console.warn(`[模块3] [${emailAddress}] Domain API响应格式无效: ${JSON.stringify(response.data)}`);
                }

            } else if (apiProvider === 'Outlook') {
                const url = `${apiConfig.url}${emailAddress}`;
                response = await axios.get(url, { timeout: 20000 });

                if (response.data && response.data.code === 0 && response.data.data && response.data.data.result) {
                    verificationCode = response.data.data.result;
                } else {
                    console.warn(`[模块3] [${emailAddress}] Outlook API响应格式或code无效: ${JSON.stringify(response.data)}`);
                }
            }
            // --- 动态API调用逻辑结束 ---

            if (verificationCode) {
                console.log(`[模块3] [${emailAddress}] 成功获取到邮箱验证码: ${verificationCode}`);
                await humanLikeType(page, config.OTP_INPUT_SELECTOR, verificationCode);
                await humanLikeClick(page, config.OTP_SUBMIT_BUTTON_SELECTOR);
                console.log("[模块3] 邮箱验证码提交完毕。");
                return; // 成功获取并提交后，直接退出函数
            }

        } catch (error) { 
            console.error(`[模块3] [${emailAddress}] 获取邮箱验证码时发生网络或请求错误 (尝试 ${attempts + 1}/40): ${error.message}`); 
        }

        // 尝试20次后，点击 "Resend Code" 按钮
        if (attempts === 19) {
            console.log("[模块3] 已尝试20次，准备点击 'Resend Code' 按钮...");
            try {
                const resendButtonSelector = 'button[data-testid="resend-otp-button"]';
                await page.waitForSelector(resendButtonSelector, { visible: true, timeout: 10000 });
                await humanLikeClick(page, resendButtonSelector);
                console.log("[模块3] 'Resend Code' 按钮已点击，将继续尝试获取验证码。");
            } catch (resendError) {
                console.error(`[模块3] 点击 'Resend Code' 按钮时出错: ${resendError.message}`);
            }
        }

        await new Promise(resolve => setTimeout(resolve, 5000)); // 每次尝试后等待5秒
    }

    // 如果循环40次后仍未获取到验证码，则抛出超时错误
    throw new Error("EMAIL_API_TIMEOUT:邮箱问题，无法获取到邮箱验证码");
}

module.exports = { verifyEmail };
