// ===================================================================================
// ### modules/05_fill_contact_info.js (V3.0 - 增加Cookie横幅处理) ###
// ===================================================================================
const { humanLikeClick, humanLikeType, getRandomDelay } = require('../shared/utils');

async function fillContactInfo(page, data, config) {
    console.log("[模块5] 等待并填写联系人信息页面...");
    await page.waitForSelector(config.CONTACT_FULL_NAME_SELECTOR, { visible: true, timeout: 180000 });
    
    // 【新增】在模块开头加入Cookie横幅处理逻辑
    try {
        const cookieAcceptButtonSelector = 'button[data-id="awsccc-cb-btn-accept"]';
        console.log('[模块5] 正在检查Cookie横幅...');
        const acceptButton = await page.waitForSelector(cookieAcceptButtonSelector, { timeout: 5000 });
        if (acceptButton) {
            await acceptButton.click();
            console.log('[模块5] Cookie横幅已点击“Accept”。');
            await new Promise(resolve => setTimeout(resolve, 1000)); // 等待横幅消失
        }
    } catch (error) {
        console.log('[模块5] 未发现Cookie横幅，或已处理，继续执行...');
    }
    
    await humanLikeClick(page, config.PERSONAL_ACCOUNT_RADIO_SELECTOR);

    await page.click(config.CONTACT_PHONE_COUNTRY_TRIGGER_SELECTOR);
    await page.waitForSelector(config.dynamicContactPhoneOptionSelector, { visible: true });
    await page.click(config.dynamicContactPhoneOptionSelector);
    
    await page.click(config.CONTACT_ADDRESS_COUNTRY_TRIGGER_SELECTOR);
    await page.waitForSelector(config.dynamicContactAddressOptionSelector, { visible: true });
    await page.click(config.dynamicContactAddressOptionSelector);
    
    await humanLikeType(page, config.CONTACT_FULL_NAME_SELECTOR, data.real_name);
    await humanLikeType(page, config.CONTACT_STREET_SELECTOR, data.street);
    await humanLikeType(page, config.CONTACT_CITY_SELECTOR, data.city);
    await humanLikeType(page, config.CONTACT_STATE_SELECTOR, data.state);
    await humanLikeType(page, config.CONTACT_POSTCODE_SELECTOR, data.postcode);
    await humanLikeType(page, config.CONTACT_PHONE_NUMBER_SELECTOR, data.phone_number);
    
    const checkboxElement = await page.waitForSelector(config.CONTACT_AGREEMENT_CHECKBOX_SELECTOR, { timeout: 180000 });
    await checkboxElement.evaluate(el => el.scrollIntoView());
    await page.evaluate(sel => document.querySelector(sel).click(), config.CONTACT_AGREEMENT_CHECKBOX_SELECTOR);
    
    await new Promise(resolve => setTimeout(resolve, getRandomDelay(1500, 2500)));
    
    await page.waitForSelector(config.CONTACT_SUBMIT_BUTTON_SELECTOR, { visible: true });
    await page.evaluate((sel) => document.querySelector(sel)?.click(), config.CONTACT_SUBMIT_BUTTON_SELECTOR);
    console.log("[模块5] 联系人信息填写完毕。");
}

module.exports = { fillContactInfo };