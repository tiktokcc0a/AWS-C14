// ===================================================================================
// ### modules/06_fill_payment_info.js (V3.2 - 修复版) ###
// ===================================================================================
const { humanLikeType, humanLikeClick, getMonthName } = require('../shared/utils');

async function fillPaymentInfo(page, data, config) {
    console.log("[模块6] 进入模块，等待并填写支付信息...");
    
    // 【已修复】在模块开始时增加页面刷新，确保页面状态干净
    console.log("[模块6] 正在刷新页面以确保状态纯净...");
    await page.reload({ waitUntil: 'networkidle0' });

    try {
        const cookieAcceptButtonSelector = 'button[data-id="awsccc-cb-btn-accept"]';
        console.log('[模块6] 正在检查Cookie横幅...');
        const acceptButton = await page.waitForSelector(cookieAcceptButtonSelector, { timeout: 5000 });
        if (acceptButton) {
            await acceptButton.click();
            console.log('[模块6] Cookie横幅已点击“Accept”。');
            await new Promise(resolve => setTimeout(resolve, 1000)); // 等待横幅消失
        }
    } catch (error) {
        console.log('[模块6] 未发现Cookie横幅，或已处理，继续执行...');
    }
    
    await page.waitForSelector(config.PAYMENT_CARD_NUMBER_SELECTOR, { visible: true, timeout: 180000 });
    await humanLikeType(page, config.PAYMENT_CARD_NUMBER_SELECTOR, data['1step_number']);
    await humanLikeType(page, config.PAYMENT_CARD_HOLDER_NAME_SELECTOR, data.real_name);
    await humanLikeType(page, config.PAYMENT_CVV_SELECTOR, data['1step_code']);
    
    const month = parseInt(data['1step_month'], 10);
    const year = parseInt(data['1step_year'], 10);
    const formattedMonth = month < 10 ? `0${month}` : `${month}`;
    const targetMonthName = getMonthName(month);

    await page.click(config.PAYMENT_MONTH_TRIGGER_SELECTOR);
    const monthOptionSelector = `div.awsui-select-option[data-value="${formattedMonth}"][title="${targetMonthName}"]`;
    await page.waitForSelector(monthOptionSelector, { visible: true });
    await page.click(monthOptionSelector);

    await page.click(config.PAYMENT_YEAR_TRIGGER_SELECTOR);
    const yearOptionSelector = `div.awsui-select-option[data-value="${year}"][title="${year}"]`;
    await page.waitForSelector(yearOptionSelector, { visible: true });
    await page.click(yearOptionSelector);

    // --- 遵照总监指示修改点击方式 ---
    console.log("[模块6] 准备使用 element.click() 方式直接点击支付按钮...");
    try {
        const buttonSelector = config.PAYMENT_SUBMIT_BUTTON_SELECTOR;
        // 等待按钮可见且可交互
        await page.waitForSelector(buttonSelector, { visible: true, timeout: 30000 }); 
        // 使用 page.evaluate 在浏览器上下文中执行最直接的 click
        await page.evaluate((sel) => {
            // puppeteer-extra的伪元素选择器::-p-text()不能直接在querySelector中使用
            // 因此，我们需要找到所有可能的按钮，然后筛选文本内容
            const buttons = Array.from(document.querySelectorAll('button'));
            const targetButton = buttons.find(button => button.innerText.includes('Verify and continue (step 3 of 5)'));
            if (targetButton) {
                targetButton.click();
            } else {
                // 如果找不到，就抛出一个错误，这样外面的catch可以捕获到
                throw new Error(`无法在DOM中找到文本为 "Verify and continue (step 3 of 5)" 的按钮。`);
            }
        }, buttonSelector); // buttonSelector 在这里主要是为了语义清晰
        console.log("[模块6 日志] 支付按钮 element.click() 指令已成功发出。");
    } catch (error) {
        console.error(`[模块6 日志] 使用 element.click() 点击支付按钮时发生错误: ${error.message}`);
        throw error; // 重新抛出错误，让主控制器处理
    }
    // --- 修改结束 ---

    console.log("[模块6] 支付信息填写完毕。");
}

module.exports = { fillPaymentInfo };