// ===================================================================================
// ### modules/03_verify_email.js (已集成“特殊超时失败”逻辑) ###
// ===================================================================================
const axios = require('axios').default;
const config = require('../shared/config');
const { humanLikeType, humanLikeClick } = require('../shared/utils');

async function verifyEmail(page, data) {
    console.log("[模块3] 等待并填写邮箱验证码...");
    await page.waitForSelector(config.OTP_INPUT_SELECTOR, { visible: true, timeout: 180000 });
    const [emailAddress, emailPassword] = data.mailbox.split('----');

    for (let attempts = 0; attempts < 40; attempts++) {
        try {
            const url = `${config.EMAIL_API_BASE_URL}?email=${emailAddress}&password=${emailPassword}&email_provider=domain`;
            const response = await axios.get(url, { timeout: 20000 });
            if (response.data && response.data.verification_code) {
                console.log(`[模块3] 成功获取到邮箱验证码: ${response.data.verification_code}`);
                await humanLikeType(page, config.OTP_INPUT_SELECTOR, response.data.verification_code);
                await humanLikeClick(page, config.OTP_SUBMIT_BUTTON_SELECTOR);
                console.log("[模块3] 邮箱验证码提交完毕。");
                return; 
            }
        } catch (error) { 
            console.error(`[模块3] 获取邮箱验证码时发生错误 (尝试 ${attempts + 1}/40): ${error.message}`); 
        }

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

        await new Promise(resolve => setTimeout(resolve, 5000));
    }

    // --- 修改点: 抛出带有特殊标识的错误，以便主控制器能够识别 ---
    throw new Error("EMAIL_API_TIMEOUT:邮箱问题，无法获取到邮箱验证码");
}

module.exports = { verifyEmail };