// ===================================================================================
// ### modules/10_create_iam_keys.js (V9.1 - FIX接入版) ###
// 基于 V9.0 版本修改，以接入主控的FIX流程
// ===================================================================================
const fs = require('fs').promises;
const path = require('path');
const config = require('../shared/config');
const { humanLikeClick } = require('../shared/utils');

async function createIamKeys(page, data) {
    const MAX_RETRIES = 3;
    const ELEMENT_WAIT_TIMEOUT = 70000;
    const NAVIGATION_TIMEOUT = 180000;

    // 主循环的重试现在主要由 main_controller 控制，此处的循环用于逻辑上的封装
    for (let attempt = 1; attempt <= MAX_RETRIES; attempt++) {
        console.log(`[模块10] 第 ${attempt}/${MAX_RETRIES} 次尝试...`);
        try {
            console.log("[模块10] 开始并行等待关键元素或封号提示...");

            const suspendedPromise = page.waitForSelector("span ::-p-text(Your AWS account was suspended because your account details couldn't be verified.)", { visible: true, timeout: ELEMENT_WAIT_TIMEOUT }).then(() => 'suspended');
            const elementsPromise = Promise.all([
                page.waitForSelector(config.IAM_UNDERSTAND_CHECKBOX, { visible: true, timeout: ELEMENT_WAIT_TIMEOUT }),
                page.waitForSelector(config.IAM_CREATE_KEY_BUTTON, { visible: true, timeout: ELEMENT_WAIT_TIMEOUT })
            ]).then(() => 'elements_ready');

            const raceResult = await Promise.race([suspendedPromise, elementsPromise]);

            if (raceResult === 'suspended') {
                console.error("[模块10] 检测到封号提示！");
                throw new Error("已被封号");
            }

            console.log("[模块10] 关键元素均已加载，准备执行操作...");

            try {
                const nextButtonSelector = 'button.awsui-button-variant-primary[data-testid="aws-onboarding-next-button"]';
                await page.click(nextButtonSelector, { timeout: 5000 });
                console.log('[模块10] 已成功点击新手引导框的"Next"按钮。');
                await new Promise(resolve => setTimeout(resolve, 1000));
            } catch (e) {
                console.log('[模块10] 未发现新手引导框，或已处理。');
            }

            await page.click(config.IAM_UNDERSTAND_CHECKBOX);
            await page.click(config.IAM_CREATE_KEY_BUTTON);

            console.log("[模块10] 等待 'Show' 按钮出现 (超时: 30秒)...");
            await page.waitForSelector(config.IAM_SHOW_SECRET_BUTTON, { visible: true, timeout: 30000 });
            console.log("[模块10] 'Show' 按钮已出现，准备点击。");

            try {
                const buttonSelector = config.IAM_SHOW_SECRET_BUTTON;
                await page.evaluate((sel) => {
                    const strongs = Array.from(document.querySelectorAll('strong'));
                    const targetButton = strongs.find(s => s.innerText.trim() === 'Show');
                    if (targetButton) {
                        targetButton.click();
                    } else {
                        throw new Error(`无法在DOM中找到文本为 "Show" 的strong元素按钮。`);
                    }
                }, buttonSelector);
                console.log("[模块10 日志] 'Show' 按钮 element.click() 指令已发出。");
            } catch(error) {
                console.error(`[模块10 日志] 点击 'Show' 按钮时发生错误: ${error.message}`);
                throw error;
            }

            await page.waitForSelector(config.IAM_SECRET_KEY_VALUE, { visible: true, timeout: 5000 });

            const accessKey = await page.$eval(config.IAM_ACCESS_KEY_VALUE, el => el.textContent.trim());
            const secretKey = await page.$eval(config.IAM_SECRET_KEY_VALUE, el => el.textContent.trim());

            if (!accessKey || !secretKey) throw new Error("未能提取到Access Key或Secret Key。");
            console.log(`[模块10] 密钥提取成功！`);

            const contentToSave = [data.account, data.password, accessKey, secretKey, data.country_full_name].join('\t');
            const saveDir = config.KEY_SAVE_PATH;
            const filePath = path.join(saveDir, `${data.account}.txt`);

            await fs.mkdir(saveDir, { recursive: true });
            await fs.writeFile(filePath, contentToSave, 'utf-8');
            console.log(`[模块10] ✅✅✅ 最终成功！账号信息已保存到: ${filePath}`);

            return { status: 'final_success' };

        } catch (error) {
            console.error(`[模块10] 第 ${attempt} 次尝试失败: ${error.message}`);

            // ===============================================================================
            // ### 核心修改：将错误向上抛出以触发FIX ###
            // 捕获到任何错误后，不再自行刷新页面重试。
            // 而是将错误直接抛出，由 main_controller.js 来捕获。
            // main_controller.js 会判断错误类型，如果是超时等网络问题，就会执行 FIX 流程（更换IP并刷新）。
            // ===============================================================================
            throw error;
        }
    }
     // 此处代码理论上不会执行，因为任何错误都会在catch中被抛出，从而中断循环。
    throw new Error(`模块10在 ${MAX_RETRIES} 次尝试后仍未成功，最终失败。`);
}

module.exports = { createIamKeys };