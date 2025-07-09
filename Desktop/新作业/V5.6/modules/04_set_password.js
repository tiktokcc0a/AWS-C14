// ===================================================================================
// ### modules/04_set_password.js (V2.0 - 增加服务器错误判定) ###
// ===================================================================================
const config = require('../shared/config');
const { humanLikeType, humanLikeClick } = require('../shared/utils');

async function setPassword(page, data) {
    console.log("[模块4] 等待并设置密码...");
    await page.waitForSelector(config.PASSWORD_INPUT_SELECTOR, { visible: true, timeout: 180000 });
    await humanLikeType(page, config.PASSWORD_INPUT_SELECTOR, data.password);
    await humanLikeType(page, config.RE_PASSWORD_INPUT_SELECTOR, data.password);
    await humanLikeClick(page, config.CREATE_PASSWORD_SUBMIT_SELECTOR);
    console.log("[模块4] 密码设置并提交完毕。");

    // 【新增判定】检查是否存在服务器错误
    try {
        console.log("[模块4] 正在检查是否存在服务器错误提示 (等待5秒)...");
        const errorSelector = 'div ::-p-text(Sorry. There is something wrong with our server)';
        await page.waitForSelector(errorSelector, { visible: true, timeout: 5000 });
        // 如果上面的行没有抛出错误，说明元素被找到了
        throw new Error("密码创建服务错误");
    } catch (error) {
        if (error.message.includes("密码创建服务错误")) {
            // 如果是我们主动抛出的错误，就再次抛出
            throw error;
        }
        // 如果是超时错误，说明没找到错误提示，这是正常流程
        console.log("[模块4] 未检测到服务器错误提示，流程正常。");
    }
}

module.exports = { setPassword };