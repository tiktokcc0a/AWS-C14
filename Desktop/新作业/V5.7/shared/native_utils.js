// ===================================================================================
// ### shared/native_utils.js ###
// ===================================================================================
async function simulateClickNative(page, selector, timeout = 180000) {
    await page.waitForSelector(selector, { visible: true, timeout });
    await page.evaluate((sel) => {
        const element = document.querySelector(sel);
        if (element) {
            var event = new MouseEvent('mousedown', { bubbles: true, cancelable: true, view: window });
            element.dispatchEvent(event);
            event = new MouseEvent('mouseup', { bubbles: true, cancelable: true, view: window });
            element.dispatchEvent(event);
            event = new MouseEvent('click', { bubbles: true, cancelable: true, view: window });
            element.dispatchEvent(event);
        } else {
            throw new Error(`simulateClickNative: 元素 ${sel} 未在DOM中找到进行点击。`);
        }
    }, selector);
}
async function waitForElementNative(page, selector, timeout = 180000) {
    let attempts = 0;
    const maxAttempts = timeout / 500;
    return new Promise((resolve, reject) => {
        const interval = setInterval(async () => {
            const element = await page.$(selector);
            if (element) {
                const isVisible = await element.isVisible();
                if (isVisible) {
                    clearInterval(interval);
                    resolve(element);
                }
            }
            if (attempts >= maxAttempts) {
                clearInterval(interval);
                reject(new Error(`waitForElementNative: 元素 ${selector} 在 ${timeout / 1000} 秒内未找到或不可见。`));
            }
            attempts++;
        }, 500); 
    });
}
module.exports = {
    simulateClickNative,
    waitForElementNative,
};