// ===================================================================================
// ### captcha_solver.js (最终日志增强与超时优化版) ###
// ===================================================================================
const axios = require('axios').default;
const { URLSearchParams } = require('url');
const { TTSHITU_API_URL, TTSHITU_USERNAME, TTSHITU_PASSWORD, TTSHITU_TYPEID } = require('./config');
const { humanLikeType, getRandomDelay } = require('./utils');

async function bulletproofClick(frame, selector) {
    for (let i = 0; i < 10; i++) {
        const success = await frame.evaluate((sel) => { 
            const button = document.querySelector(sel); 
            if (button && !button.disabled && button.offsetHeight > 0) { 
                button.click(); 
                return true; 
            } 
            return false; 
        }, selector);
        if (success) return true;
        await new Promise(resolve => setTimeout(resolve, 1000));
    }
    return false;
}

/**
 * 使用天天识图解决图片验证码的核心函数。
 * @param {Function} getContextFunc - 一个用于获取当前page或frame对象的异步函数。
 * @param {string} imageSelector - 验证码图片的CSS选择器。
 * @param {string} inputSelector - 验证码输入框的CSS选择器。
 * @param {string} submitSelector - 提交按钮的CSS选择器。
 * @param {string} errorTextSelector - 错误信息提示的CSS选择器。
 * @param {string} successCheckSelector - 用于判断成功的选择器（例如，验证码iframe消失或下一个页面的元素出现）。
 * @param {number} maxRetries - 模块内部的最大重试次数。
 * @param {string} expectedErrorText - 预期的、可接受的错误文本（例如 "验证码错误"）。
 * @param {boolean} isCaptchaInIframe - 验证码是否在iframe中。
 * @param {number} timeout - 所有waitFor操作的超时时间（毫秒）。
 * @returns {Promise<boolean>} - 返回一个解析为true的Promise表示成功。
 * @throws {Error} - 如果所有尝试都失败，或发生致命错误，则抛出异常。
 */
// 【修改】将默认超时时间统一调整为60秒
async function solveImageCaptchaWithTTShitu(getContextFunc, imageSelector, inputSelector, submitSelector, errorTextSelector, successCheckSelector, maxRetries = 3, expectedErrorText = "...", isCaptchaInIframe = false, timeout = 60000) {
    for (let i = 0; i < maxRetries; i++) {
        let currentContext;
        try {
            currentContext = await getContextFunc();
            if (!currentContext) throw new Error("无法获取有效的 Page/Frame 对象。");

            console.log(`[打码模块] 尝试解决验证码 (第 ${i + 1} / ${maxRetries} 次)...`);
            
            await currentContext.waitForSelector(imageSelector, { visible: true, timeout });
            
            const imageUrl = await currentContext.$eval(imageSelector, img => img.src);
            // 【新增日志】输出获取到的图片URL
            console.log(`[打码模块] 成功获取图片URL: ${imageUrl}`);
            
            console.log('[打码模块] 正在下载图片并请求天天识图API...');
            const postData = new URLSearchParams({
                'username': TTSHITU_USERNAME, 'password': TTSHITU_PASSWORD, 'typeid': TTSHITU_TYPEID,
                'image': Buffer.from((await axios.get(imageUrl, { responseType: 'arraybuffer' })).data).toString('base64')
            }).toString();
            
            const solveResponse = await axios.post(TTSHITU_API_URL, postData, { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } });
            
            // 【新增日志】输出天天识图的完整响应
            console.log('[打码模块] 收到天天识图响应:', JSON.stringify(solveResponse.data));

            if (!solveResponse.data.success) throw new Error("天天识图打码失败: " + solveResponse.data.message);
            
            const captchaCode = solveResponse.data.data.result.toLowerCase();
            console.log(`[打码模块] 打码成功，识别结果: ${captchaCode}`);
            
            await humanLikeType(currentContext, inputSelector, captchaCode);
            if (!await bulletproofClick(currentContext, submitSelector)) throw new Error(`点击验证码Submit按钮失败。`);
            
            let successFlag = false;
            if (isCaptchaInIframe) {
                console.log(`[打码模块] 等待 iframe (${successCheckSelector}) 消失... (超时: ${timeout / 1000}秒)`);
                try {
                    await currentContext.page().waitForFunction(selector => !document.querySelector(selector), { timeout }, successCheckSelector);
                    successFlag = true;
                } catch (waitError) {
                    if (waitError.message.includes("detached Frame") || waitError.message.includes("Execution context was destroyed")) {
                        successFlag = true; // iframe消失导致上下文销毁，视为成功
                    } else {
                        throw waitError; // 其他等待错误，正常抛出
                    }
                }
            } else {
                console.log(`[打码模块] 等待成功标志 (${successCheckSelector}) 或错误出现... (超时: ${timeout / 1000}秒)`);
                const outcome = await Promise.race([
                    currentContext.waitForSelector(successCheckSelector, { visible: true, timeout }).then(() => 'success'),
                    currentContext.waitForFunction((sel, txt) => { const el = document.querySelector(sel); return el && el.offsetParent !== null && el.innerText.includes(txt); }, { visible: true, timeout }, errorTextSelector, expectedErrorText).then(() => 'error')
                ]).catch(() => 'timeout');
                if (outcome === 'success') successFlag = true;
            }

            if (successFlag) {
                console.log('[打码模块] 验证码提交通用成功。');
                return true;
            }
            
            console.warn(`[打码模块] 未检测到成功标志，准备模块内重试...`);
            await new Promise(resolve => setTimeout(resolve, getRandomDelay(2000, 4000)));

        } catch (error) {
            console.error(`[打码模块] 验证码尝试失败 (第 ${i + 1} 次): ${error.message}`);
            // 遇到致命错误，直接抛出，让主工作流处理
            if (error.message.includes("detached Frame") || error.message.includes("Execution context was destroyed") || error.message.includes("Target closed")) {
                throw new Error("浏览器上下文已失效，无法重试。");
            }
            // 如果是最后一次尝试，则抛出最终错误
            if (i >= maxRetries - 1) {
                 throw error;
            }
            // 否则，等待后继续下一次循环
            await new Promise(resolve => setTimeout(resolve, getRandomDelay(1500, 3500)));
        }
    }
    // 如果循环结束仍未成功，抛出最终错误
    throw new Error(`验证码重试 ${maxRetries} 次后仍未成功。`);
}

module.exports = { solveImageCaptchaWithTTShitu };