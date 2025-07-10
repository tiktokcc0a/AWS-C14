// ===================================================================================
// ### shared/browser_setup.js (V5.0 - 增加启动重试机制) ###
// ===================================================================================
const puppeteer = require('puppeteer-extra');
const StealthPlugin = require('puppeteer-extra-plugin-stealth');
const { openBrowser, closeBrowser, createBrowser, deleteBrowser } = require('../request');
puppeteer.use(StealthPlugin());

const FINGERPRINT_OPTIONS = [
    { coreVersion: '136', os: 'Win32', osVersion: '11', platform: 'Win32' }
];

async function setupBrowser(instanceId, isHeadless = false, proxyPort, browserIndex) {
    const fingerprint = FINGERPRINT_OPTIONS[browserIndex % FINGERPRINT_OPTIONS.length];
    console.log(`[浏览器设置 ${instanceId}] 使用端口 ${proxyPort} 和指纹:`, fingerprint);

    let browserId = null;
    // 【修改点3】创建浏览器增加3次重试
    for (let attempt = 1; attempt <= 3; attempt++) {
        try {
            console.log(`[浏览器设置 ${instanceId}] 正在创建浏览器配置文件 (第 ${attempt}/3 次尝试)...`);
            const createRes = await createBrowser({
                name: `AWS-WF-${instanceId}`,
                remark: `AWS工作流实例 ${instanceId}`,
                proxyMethod: 2,
                proxyType: 'socks5',
                host: '127.0.0.1',
                port: proxyPort.toString(),
                browserFingerPrint: {
                    ...fingerprint,
                    "isIpCreateLanguage": false,
                    "displayLanguages": "en-US",
                    "languages": "en-US"
                }
            });

            if (!createRes.success) throw new Error(`创建浏览器API返回失败: ${createRes.msg}`);
            browserId = createRes.data.id;
            console.log(`[浏览器设置 ${instanceId}] 配置文件创建成功，ID: ${browserId}`);
            break; // 成功则跳出循环
        } catch (error) {
            console.error(`[浏览器设置 ${instanceId}] 创建浏览器时出错 (尝试 ${attempt}/3): ${error.message}`);
            if (attempt >= 3) {
                throw new Error(`[${instanceId}] 创建浏览器失败，已达最大重试次数: ${error.message}`);
            }
            await new Promise(resolve => setTimeout(resolve, 2000)); // 等待2秒后重试
        }
    }
    
    // 【修改点3】打开浏览器增加3次重试
    let openRes = null;
    for (let attempt = 1; attempt <= 3; attempt++) {
        try {
            console.log(`[浏览器设置 ${instanceId}] 正在打开浏览器窗口 (第 ${attempt}/3 次尝试)...`);
            openRes = await openBrowser({ id: browserId, args: isHeadless ? ['--headless=new'] : [] });
            if (!openRes.success) {
                throw new Error(`打开浏览器API返回失败: ${openRes.msg}`);
            }
            console.log(`[浏览器设置 ${instanceId}] 打开浏览器API调用成功。`);
            break; // 成功则跳出循环
        } catch (error) {
            console.error(`[浏览器设置 ${instanceId}] 打开浏览器时出错 (尝试 ${attempt}/3): ${error.message}`);
            if (attempt >= 3) {
                try { await deleteBrowser(browserId); } catch (e) {} // 最终失败前尝试清理已创建的配置文件
                throw new Error(`[${instanceId}] 打开浏览器失败，已达最大重试次数: ${error.message}`);
            }
            await new Promise(resolve => setTimeout(resolve, 2000)); // 等待2秒后重试
        }
    }

    const browser = await puppeteer.connect({ browserWSEndpoint: openRes.data.ws, defaultViewport: null });
    const page = (await browser.pages())[0] || await browser.newPage();
    console.log(`[浏览器设置 ${instanceId}] 浏览器 (ID: ${browserId}) 连接成功。`);
    
    return { browser, page, browserId };
}


async function tearDownBrowser(browserId) {
    if (!browserId) return;
    console.log(`[清理 ${browserId}] 开始执行清理流程...`);

    try {
        const closeRes = await closeBrowser(browserId);
        if (closeRes && closeRes.success) {
            console.log(`[清理 ${browserId}] 浏览器窗口已成功关闭。`);
        } else {
            console.warn(`[清理 ${browserId}] 关闭浏览器窗口时API返回失败或未成功: ${closeRes?.msg || '未知错误'}`);
        }
    } catch (error) {
        console.error(`[清理 ${browserId}] 调用关闭API时发生网络错误: ${error.message}`);
    }

    console.log(`[清理 ${browserId}] 等待2秒，确保浏览器进程完全退出...`);
    await new Promise(resolve => setTimeout(resolve, 2000));

    try {
        console.log(`[清理 ${browserId}] 正在发送删除命令...`);
        const deleteRes = await deleteBrowser(browserId);
        if (deleteRes && deleteRes.success) {
            console.log(`[清理 ${browserId}] ✅ 浏览器配置文件已成功删除。清理完成！`);
        } else {
            throw new Error(`删除浏览器配置文件时API返回失败: ${deleteRes?.msg || '未知错误'}`);
        }
    } catch (error) {
        console.error(`[清理 ${browserId}] ❌ 清理浏览器时发生最终错误: ${error.message}`);
        throw error;
    }
}

module.exports = { setupBrowser, tearDownBrowser };