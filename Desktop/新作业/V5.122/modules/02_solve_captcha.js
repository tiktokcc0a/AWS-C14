// ===================================================================================
// ### modules/02_solve_captcha.js (V3.0 - 净化通用版) ###
// ===================================================================================
const config = require('../shared/config');
const { solveImageCaptchaWithTTShitu } = require('../shared/captcha_solver');

/**
 * 解决通用图片验证码。
 * 本模块现在是纯粹的执行单元，不再包含任何业务流程判断。
 * @param {import('puppeteer').Page} page - Puppeteer的Page对象。
 * @param {object} data - 包含账户信息的对象。
 */
async function solveCaptcha(page, data) {
    console.log('[模块2] 开始处理图形验证码...');

    try {
        // 定义一个函数，用于获取包含验证码的iframe上下文
        const getCaptchaFrame = async () => {
            console.log(`[模块2] 等待验证码iframe: "${config.COMMON_CAPTCHA_IFRAME_SELECTOR}" 加载...`);
            const iframeElement = await page.waitForSelector(config.COMMON_CAPTCHA_IFRAME_SELECTOR, { visible: true, timeout: 60000 });
            const frame = await iframeElement.contentFrame();
            if (!frame) throw new Error("无法获取验证码iframe的contentFrame。");
            return frame;
        };

        // 直接调用共享的打码平台函数
        // 这个函数内部包含了获取图片、调用API、填写、提交、检查错误等所有逻辑
        await solveImageCaptchaWithTTShitu(
            getCaptchaFrame,
            config.INITIAL_CAPTCHA_IMAGE_SELECTOR,
            config.INITIAL_CAPTCHA_INPUT_SELECTOR,
            config.INITIAL_CAPTCHA_SUBMIT_SELECTOR,
            config.INITIAL_CAPTCHA_ERROR_SELECTOR,
            config.COMMON_CAPTCHA_IFRAME_SELECTOR, // 成功标志是iframe消失
            3, // 内部重试次数
            config.INITIAL_CAPTCHA_EXPECTED_ERROR,
            true // 验证码是在iframe中
        );
        
        console.log('[模块2] 图形验证码流程成功完成！');

    } catch (error) {
        // 将任何发生的错误向上抛出，由主控统一处理
        console.error(`[模块2] 处理验证码时发生严重错误: ${error.message}`);
        throw new Error(`验证码模块执行失败: ${error.message}`);
    }
}

module.exports = { solveCaptcha };