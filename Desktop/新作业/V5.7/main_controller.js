// ===================================================================================
// ### main_controller.js (V21.1 - 漏洞修复与批量清理支持版) ###
// ### MODIFIED: 根据讨论修改模块9.5的触发逻辑 ###
// ===================================================================================
const fs = require('fs').promises;
const path = require('path');
const axios =require('axios').default;
const { setupBrowser, tearDownBrowser } = require('./shared/browser_setup');
const staticConfig = require('./shared/config');
const { NetworkWatcher } = require('./utils/network_watcher');

// --- 全局状态与配置 ---
const args = process.argv.slice(2);
const windowArg = args.find(arg => arg.startsWith('--window='));
const instanceId = windowArg ? windowArg.split('=')[1] : 'W0';

const proxyPortArg = args.find(arg => arg.startsWith('--proxy_port='));
const PROXY_PORT = proxyPortArg ? parseInt(proxyPortArg.split('=')[1], 10) : 45000;

const headlessArg = args.find(arg => arg === '--headless=new' || arg === '--headless');
const IS_HEADLESS = !!headlessArg;

const pauseState = {};
let currentPage = null;

const KNOWN_FAILURE_MESSAGES = [
    "出现分区", "死卡", "红窗", "EMAIL_API_TIMEOUT", "REGISTRATION_FAILED_INCOMPLETE", "红窗ES",
    "密码创建服务错误", "已被封号", "模块10在 3 次尝试后仍未成功，最终失败。", "验证码重试 3 次后仍未成功。"
];

const RECOVERABLE_NETWORK_ERRORS = [
    'timeout', 'err_timed_out', 'err_socks_connection_failed', 'err_proxy_connection_failed',
    'err_connection_reset', 'err_connection_timed_out', 'err_internet_disconnected',
    'err_address_unreachable', 'err_connection_refused', 'net::err_empty_response',
    'net::err_network_changed', 'net::err_cert_authority_invalid'
];


// --- 模块与工作流定义 ---
const modules = {
    '01_fillSignupForm': require('./modules/01_fill_signup_form').fillSignupForm,
    '02_solveCaptcha': require('./modules/02_solve_captcha').solveCaptcha,
    '03_verifyEmail': require('./modules/03_verify_email').verifyEmail,
    '04_setPassword': require('./modules/04_set_password').setPassword,
    '05_fillContactInfo': require('./modules/05_fill_contact_info').fillContactInfo,
    '06_fillPaymentInfo': require('./modules/06_fill_payment_info').fillPaymentInfo,
    '07_enterPhoneNumber': require('./modules/07_enter_phone_number').enterPhoneNumber,
    '08_verifySms': require('./modules/08_verify_sms').verifySms,
    '09_selectSupportPlan': require('./modules/09_select_support_plan').selectSupportPlan,
    '9.5_handleConfirmation': require('./modules/9.5_handle_confirmation').handleConfirmation,
    '10_createIamKeys': require('./modules/10_create_iam_keys').createIamKeys,
};

// ======================= 修改点 1 of 2: 移除旧的 'confirmation' 规则 =======================
const WORKFLOWS = {
    'signup?request_type=register': ['01_fillSignupForm', '02_solveCaptcha', '03_verifyEmail', '04_setPassword'],
    '#/account': ['05_fillContactInfo'],
    '#/paymentinformation': ['06_fillPaymentInfo'],
    '#/identityverification': ['07_enterPhoneNumber', '02_solveCaptcha', '08_verifySms'],
    '#/support': ['09_selectSupportPlan'],
    'security_credentials': ['10_createIamKeys']
};
// =========================================================================================


// --- 配置生成器 ---
async function generateDynamicConfig(countryCode, countryFullName) {
    console.log(`[配置生成器 ${instanceId}] 正在为国家代码 "${countryCode}" 生成动态配置...`);
    const countryDataPath = path.join(__dirname, 'shared', 'combined_country_data.json');
    const countryData = JSON.parse(await fs.readFile(countryDataPath, 'utf-8'));

    const countryInfo = Object.entries(countryData).find(([, data]) => data.country_code === countryCode);

    if (!countryInfo) {
        throw new Error(`[配置生成器 ${instanceId}] 无法在 a中找到国家代码为 "${countryCode}" 的条目。`);
    }
    const [countryNameFound, countryDetails] = countryInfo;
    const { dialing_code } = countryDetails;

    const resolvedCountryFullName = countryFullName || countryNameFound;

    const dynamicConfig = {
        dynamicContactPhoneOptionSelector: `div[data-value="${countryCode}"][title="${resolvedCountryFullName} (+${dialing_code})"]`,
        dynamicContactAddressOptionSelector: `div[data-value="${countryCode}"][title="${resolvedCountryFullName}"]`,
        dynamicIdentityPhoneOptionSelector: `div[data-value="${countryCode}"][title="${resolvedCountryFullName} (+${dialing_code})"]`,
        countryCode: countryCode,
        countryFullName: resolvedCountryFullName
    };
    console.log(`[配置生成器 ${instanceId}] 动态配置生成成功:`, dynamicConfig);
    return dynamicConfig;
}


// --- 监听来自Python GUI的命令 ---
process.stdin.on('data', async (data) => {
    const command = data.toString().trim();
    if (command.startsWith("PAUSE::")) {
        const targetInstanceId = command.split("::")[1];
        if (targetInstanceId === instanceId) {
            pauseState[instanceId] = true;
            console.log(`[主控 ${instanceId}] 收到命令: 暂停 ${instanceId}`);
        }
    } else if (command.startsWith("RESUME::")) {
        const targetInstanceId = command.split("::")[1];
        if (targetInstanceId === instanceId) {
            delete pauseState[instanceId];
            console.log(`[主控 ${instanceId}] 收到命令: 恢复 ${instanceId}`);
        }
    } else if (command.startsWith("TERMINATE::")) {
        const targetInstanceId = command.split("::")[1];
        if (targetInstanceId === instanceId) {
            console.log(`[主控 ${instanceId}] 收到命令: TERMINATE，准备退出。`);
            process.exit(0);
        }
    } else if (command.startsWith("SCREENSHOT::")) {
        const targetInstanceId = command.split("::")[1];
        if (targetInstanceId === instanceId && currentPage) {
            console.log(`[主控 ${instanceId}] 收到命令: SCREENSHOT，正在截取实时图像...`);
            try {
                const screenshotDir = path.join(__dirname, 'screenshot');
                await fs.mkdir(screenshotDir, { recursive: true });
                const timestamp = new Date().toISOString().replace(/[:.-]/g, '');
                const screenshotPath = path.join(screenshotDir, `manual_screenshot_${instanceId}_${timestamp}.png`);

                await currentPage.screenshot({ path: screenshotPath, fullPage: true });
                console.log(`[主控 ${instanceId}] 截图成功！已保存至: ${screenshotPath}`);
            } catch (screenshotError) {
                console.error(`[主控 ${instanceId}] 截取手动截图时发生错误: ${screenshotError.message}`);
            }
        }
    }
});


// --- 核心辅助函数 ---
async function executeFixProcess(port, page, reason) {
    console.log(`[${instanceId} FIX] 触发原因: ${reason}. 开始执行FIX流程...`);
    try {
        const countryCodeForProxy = global.signupData.country_code;
        console.log(`[${instanceId} FIX] 正在为端口 ${port} 请求更换IP (国家: ${countryCodeForProxy})...`);
        const response = await axios.post('http://localhost:8080/api/proxy/start', {
            line: "Line A (AS Route)", country_code: countryCodeForProxy, start_port: port, count: 1, time: 30
        }, { headers: { 'Content-Type': 'application/json' }, timeout: 25000 });
        console.log(`[${instanceId} FIX] IP更换API响应:`, response.data);
        console.log(`[${instanceId} FIX] IP更换成功，准备刷新页面...`);
        try {
            await page.reload({ waitUntil: 'load', timeout: 120000 });
            console.log(`[${instanceId} FIX] 页面刷新成功。`);
            return true;
        } catch (reloadError) {
            console.error(`[${instanceId} FIX] FIX中刷新页面超时: ${reloadError.message}`);
            return false;
        }
    } catch (error) {
        console.error(`[${instanceId} FIX] FIX流程执行失败! 错误: ${error.message}`);
        if (page) {
            try {
                const screenshotDir = path.join(__dirname, 'screenshot');
                await fs.mkdir(screenshotDir, { recursive: true });
                const screenshotPath = path.join(screenshotDir, `fix_failed_screenshot_${instanceId}_${Date.now()}.png`);
                await page.screenshot({ path: screenshotPath, fullPage: true });
                console.log(`[${instanceId} FIX] FIX失败截图已保存至: ${screenshotPath}`);
            } catch (screenshotError) {
                console.error(`[${instanceId} FIX] 截取FIX失败截图时发生错误: ${screenshotError.message}`);
            }
        }
        return false;
    }
}

async function saveFailedCardInfo(data, reason) {
    try {
        const info = [data['1step_number'], `${data['1step_month']}/${data['1step_year']}`, data.real_name, reason].join('|');
        const filePath = path.join(__dirname, 'data', 'Not used cards.txt');
        await fs.appendFile(filePath, info + '\n', 'utf-8');
        console.log(`[${instanceId} 错误处理] 已将失败的卡信息及原因保存至 ${path.basename(filePath)}`);
    } catch (error) {
        console.error(`[${instanceId} 错误处理] 保存失败的卡信息时发生错误: ${error.message}`);
    }
}

// --- 主工作流函数 ---
async function runWorkflow(signupData, finalConfig) {
    global.signupData = signupData;
    const MAX_MODULE_RETRIES = 3;
    const NAVIGATION_TIMEOUT = 120000;
    const START_NAVIGATION_TIMEOUT = 60000;
    const MAX_STANDBY_TIME = 70000;
    const STANDBY_CHECK_INTERVAL = 5000;

    const MAX_CONSECUTIVE_FIXES = 3;
    let consecutiveFixes = 0;

    const BROWSER_ID_TRACKING_FILE = path.join(__dirname, 'data', 'managed_browser_ids.json');

    let page, browserId = null;
    currentPage = null;
    let networkWatcher = null;
    const sharedState = { networkInterrupted: false };

    let reachedModule7 = false;
    let lastActiveWorkflowKey = 'signup?request_type=register';

    const reportStatus = (status, details = "") => {
        const account = signupData.account || 'N/A';
        console.log(`STATUS_UPDATE::${JSON.stringify({ instanceId, account, status, details: details.substring(0, 150) })}`);
    };

    const isRecoverableError = (error) => {
        const errorMessage = error.message.toLowerCase();
        const isNetworkError = RECOVERABLE_NETWORK_ERRORS.some(errSig => errorMessage.includes(errSig));
        const isModule2CaptchaTimeout = error.message.includes("Waiting for selector `iframe#core-container` failed:") && error.message.includes("ms exceeded");
        return isNetworkError || isModule2CaptchaTimeout;
    };


    try {
        let browserSetupSuccess = false;
        for (let attempt = 1; attempt <= 3; attempt++) {
            try {
                reportStatus("初始化", `启动浏览器，端口: ${PROXY_PORT} (第 ${attempt}/3 次尝试)`);
                ({ page, browserId } = await setupBrowser(instanceId, IS_HEADLESS, PROXY_PORT, 0));
                currentPage = page;

                // =======================================================================
                // ### 新增功能：记录已创建的浏览器ID ###
                // 将成功创建的 browserId 添加到追踪文件中，供批量清理脚本使用。
                // =======================================================================
                try {
                    await fs.mkdir(path.dirname(BROWSER_ID_TRACKING_FILE), { recursive: true });
                    let existingIds = [];
                    try {
                        const content = await fs.readFile(BROWSER_ID_TRACKING_FILE, 'utf-8');
                        existingIds = JSON.parse(content);
                    } catch (e) { /* 文件不存在或为空，属正常情况 */ }
                    if (!existingIds.includes(browserId)) {
                        existingIds.push(browserId);
                        await fs.writeFile(BROWSER_ID_TRACKING_FILE, JSON.stringify(existingIds, null, 2));
                        console.log(`[${instanceId} 追踪] 已记录浏览器ID ${browserId} 用于集中管理。`);
                    }
                } catch (trackingError) {
                    console.error(`[${instanceId} 追踪] 记录浏览器ID时出错: ${trackingError.message}`);
                }

                browserSetupSuccess = true;
                break;
            } catch (error) {
                console.error(`[${instanceId} 工作流启动失败] 浏览器设置第 ${attempt} 次尝试失败: ${error.message}`);
                reportStatus("失败", `[${instanceId}] 浏览器启动失败 (尝试 ${attempt}/3): ${error.message}`);
                if (attempt >= 3) {
                    throw error;
                }
                await new Promise(resolve => setTimeout(resolve, 3000));
            }
        }

        networkWatcher = new NetworkWatcher(sharedState, instanceId);
        networkWatcher.start();
        const workflowState = {};
        let standbyTime = 0;

        let initialNavigationSuccess = false;
        while (!initialNavigationSuccess) {
            try {
                const navigationPromise = page.goto(finalConfig.AWS_SIGNUP_URL, { waitUntil: 'load', timeout: NAVIGATION_TIMEOUT });
                const watchdogPromise = new Promise((_, reject) =>
                    setTimeout(() => reject(new Error('START_NAVIGATION_TIMEOUT')), START_NAVIGATION_TIMEOUT)
                );

                await Promise.race([navigationPromise, watchdogPromise]);
                initialNavigationSuccess = true;

            } catch (error) {
                if (isRecoverableError(error) || error.message.includes('START_NAVIGATION_TIMEOUT')) {
                    const reason = error.message.includes('START_NAVIGATION_TIMEOUT') ? "启动导航超时(60秒)" : "初始页面加载网络错误";
                    reportStatus("错误", `${reason}，执行FIX...`);
                    consecutiveFixes++;
                    if (consecutiveFixes > MAX_CONSECUTIVE_FIXES) {
                        throw new Error(`初始页面加载连续FIX失败 ${MAX_CONSECUTIVE_FIXES} 次。`);
                    }
                    const fixSuccess = await executeFixProcess(PROXY_PORT, page, reason);
                    if (!fixSuccess) {
                         console.error(`[${instanceId}] 初始导航的FIX失败。`);
                         reportStatus("错误", `${reason}，且FIX流程也失败了。`);
                    }
                } else { throw error; }
            }
        }

        consecutiveFixes = 0;

        page.on('load', () => {
            const loadedUrl = page.url();
            console.log(`[${instanceId} 事件] 页面加载: ${loadedUrl.substring(0, 80)}...`);
            for (const urlPart in WORKFLOWS) {
                if (loadedUrl.includes(urlPart)) {
                    workflowState[urlPart] = 0;
                    lastActiveWorkflowKey = urlPart;
                }
            }
        });

        let allWorkflowsComplete = false;
        mainLoop: while (!allWorkflowsComplete) {
            if (pauseState[instanceId]) {
                reportStatus("暂停中", "用户手动暂停");
                while (pauseState[instanceId]) { await new Promise(resolve => setTimeout(resolve, 2000)); }
                reportStatus("运行中", "已从暂停中恢复...");
            }
            if (sharedState.networkInterrupted) {
                consecutiveFixes++;
                if (consecutiveFixes > MAX_CONSECUTIVE_FIXES) { throw new Error(`网络中断连续FIX ${MAX_CONSECUTIVE_FIXES} 次。`); }
                reportStatus("网络中断", "检测到中断，执行FIX...");
                const fixSuccess = await executeFixProcess(PROXY_PORT, page, "网络观察员检测到中断");
                sharedState.networkInterrupted = false;
                if(fixSuccess) {
                    reportStatus("运行中", "网络FIX完成...");
                    networkWatcher.start();
                } else {
                    console.error(`[${instanceId}] 网络中断FIX失败。`);
                    reportStatus("错误", "网络中断FIX失败。");
                }
            }
            await new Promise(resolve => setTimeout(resolve, STANDBY_CHECK_INTERVAL));
            const currentUrl = page.url();
            if (currentUrl.includes('/signup/incomplete')) { throw new Error("REGISTRATION_FAILED_INCOMPLETE"); }

            let activeWorkflowKey = null;
            for (const urlPart in WORKFLOWS) {
                const isComplete = (workflowState[urlPart] || 0) >= WORKFLOWS[urlPart].length;
                if (currentUrl.includes(urlPart) && !isComplete) { activeWorkflowKey = urlPart; break; }
            }
            
            // ======================= 修改点 2 of 2: 添加新的触发逻辑 =======================
            if (activeWorkflowKey) {
                standbyTime = 0;

                if (activeWorkflowKey !== lastActiveWorkflowKey) {
                    console.log(`[主控 ${instanceId}] 流程进展: 从 ${lastActiveWorkflowKey} 切换到 ${activeWorkflowKey}。重置FIX计数器。`);
                    consecutiveFixes = 0;
                    lastActiveWorkflowKey = activeWorkflowKey;
                }

                let currentIndex = workflowState[activeWorkflowKey] || 0;
                const moduleName = WORKFLOWS[activeWorkflowKey][currentIndex];

                if (activeWorkflowKey === '#/identityverification' && moduleName === '07_enterPhoneNumber') {
                    reachedModule7 = true;
                }

                if (activeWorkflowKey === 'signup?request_type=register' && moduleName === '02_solveCaptcha') {
                    console.log(`[主控 ${instanceId}] 进入模块2前置判断...`);
                    try {
                        await page.waitForSelector(finalConfig.OTP_INPUT_SELECTOR, { visible: true, timeout: 6000 });
                        console.log(`[主控 ${instanceId}] 检测到OTP输入框，跳过模块2！`);
                        workflowState[activeWorkflowKey]++;
                        reportStatus("流程优化", "跳过图形验证码");
                        continue mainLoop;
                    } catch (e) {
                        console.log(`[主控 ${instanceId}] 未发现OTP输入框，执行模块2。`);
                    }
                }
                let moduleRetries = 0;
                while (moduleRetries < MAX_MODULE_RETRIES) {
                    try {
                        reportStatus("运行中", `模块: ${moduleName} (尝试 ${moduleRetries + 1})`);
                        const result = await modules[moduleName](page, signupData, finalConfig);
                        console.log(`[${instanceId} 成功] 模块 ${moduleName} 执行完毕。`);
                        workflowState[activeWorkflowKey]++;
                        if (result?.status === 'final_success') allWorkflowsComplete = true;
                        break;
                    } catch (error) {
                        console.error(`[${instanceId} 失败] 模块 ${moduleName} 第 ${moduleRetries + 1} 次尝试出错: ${error.message.substring(0, 200)}`);
                        reportStatus("错误", `模块 ${moduleName} 出错: ${error.message}`);

                        const isKnownFailure = KNOWN_FAILURE_MESSAGES.some(msg => error.message.includes(msg));
                        if (isKnownFailure) {
                            throw error;
                        }

                        if (isRecoverableError(error) || error.message.includes("PHONE_NUMBER_UPDATED_AND_RELOADED")) {
                            if (!error.message.includes("PHONE_NUMBER_UPDATED_AND_RELOADED")) {
                                consecutiveFixes++;
                                if (consecutiveFixes > MAX_CONSECUTIVE_FIXES) { throw new Error(`模块 ${moduleName} 连续FIX ${MAX_CONSECUTIVE_FIXES} 次失败。`); }
                                const fixSuccess = await executeFixProcess(PROXY_PORT, page, `模块 ${moduleName} 发生网络错误`);
                                if (!fixSuccess) {
                                    console.error(`[${instanceId}] 模块错误FIX失败。`);
                                    reportStatus("错误", `模块 ${moduleName} FIX失败。`);
                                }
                            }
                            continue mainLoop;
                        }

                        moduleRetries++;
                        if (moduleRetries >= MAX_MODULE_RETRIES) { throw new Error(`[${instanceId}] 模块 ${moduleName} 已达最大重试次数。`); }
                        console.log(`[${instanceId} 重试] 准备刷新页面...`);
                        try {
                            await page.reload({ waitUntil: 'load', timeout: NAVIGATION_TIMEOUT });
                            continue mainLoop;
                        } catch (reloadError) {
                            if (isRecoverableError(reloadError)) {
                                consecutiveFixes++;
                                if (consecutiveFixes > MAX_CONSECUTIVE_FIXES) { throw new Error(`重试刷新时连续FIX ${MAX_CONSECUTIVE_FIXES} 次失败。`); }
                                const fixSuccess = await executeFixProcess(PROXY_PORT, page, `重试时刷新页面发生网络错误`);
                                if (!fixSuccess) {
                                    console.error(`[${instanceId}] 重试刷新页面FIX失败。`);
                                    reportStatus("错误", "重试刷新FIX失败。");
                                }
                                continue mainLoop;
                            } else {
                                throw reloadError;
                            }
                        }
                    }
                }
            } else if (!currentUrl.includes('#/support') && lastActiveWorkflowKey === '#/support') {
                // 新增的逻辑：如果没有任何工作流匹配，但我们刚从 support 页面跳转过来，
                // 那么就认为这是确认步骤，执行模块 9.5
                standbyTime = 0;
                console.log(`[主控 ${instanceId}] 检测到从 support 页面跳转，触发宽泛确认逻辑，执行模块 9.5...`);
                reportStatus("运行中", `模块: 9.5_handleConfirmation (宽泛触发)`);
                try {
                    await modules['9.5_handleConfirmation'](page, signupData, finalConfig);
                    // 模块9.5会导航到新页面，主循环将在下一次迭代中处理新页面的URL
                    lastActiveWorkflowKey = '9.5_handleConfirmation'; // 更新最后活动key，防止重复触发
                } catch (error) {
                     console.error(`[${instanceId} 失败] 宽泛触发的模块 9.5_handleConfirmation 执行出错: ${error.message}`);
                     throw error; // 抛出错误以激活主错误处理流程
                }
            } else {
                standbyTime += STANDBY_CHECK_INTERVAL;
                reportStatus("待机", `等待页面跳转 (已待机 ${standbyTime / 1000}秒)`);
                if (standbyTime >= MAX_STANDBY_TIME) {
                    const screenshotDir = path.join(__dirname, 'screenshot');
                    await fs.mkdir(screenshotDir, { recursive: true });
                    const screenshotPath = path.join(screenshotDir, `standby_timeout_screenshot_${instanceId}_${Date.now()}.png`);
                    console.log(`[主控 ${instanceId}] 待机超时！正在截图...`);
                    try {
                        await currentPage.screenshot({ path: screenshotPath, fullPage: true });
                        console.log(`[主控 ${instanceId}] 截图已保存: ${screenshotPath}`);
                    } catch (screenshotError) {
                        console.error(`[主控 ${instanceId}] 截取待机超时截图失败: ${screenshotError.message}`);
                    }
                    consecutiveFixes++;
                    if (consecutiveFixes > MAX_CONSECUTIVE_FIXES) { throw new Error(`待机超时连续FIX ${MAX_CONSECUTIVE_FIXES} 次失败。`); }
                    const fixSuccess = await executeFixProcess(PROXY_PORT, page, `待机超时 (${standbyTime / 1000}秒)`);
                    if (fixSuccess) {
                        standbyTime = 0;
                        continue mainLoop;
                    } else {
                        console.error(`[${instanceId}] 待机超时FIX失败。`);
                        reportStatus("错误", "待机超时FIX失败。");
                        standbyTime = 0;
                    }
                }
            }
            // ===================================================================================
        }
        reportStatus("成功", "所有工作流执行完毕！");
        console.log(`\n🎉🎉🎉 [${instanceId} 任务完成] 工作流成功！ 🎉🎉🎉`);
        await tearDownBrowser(browserId);
        process.exit(0);

    } catch (error) {
        const errorMessage = error.message;
        console.error(`\n[${instanceId} 工作流失败] 发生严重错误:`, errorMessage);

        const isKnownFailure = KNOWN_FAILURE_MESSAGES.some(msg => errorMessage.includes(msg));
        const finalErrorMessage = isKnownFailure
            ? KNOWN_FAILURE_MESSAGES.find(msg => errorMessage.includes(msg))
            : errorMessage;

        reportStatus("失败", `[${instanceId}] ` + finalErrorMessage);

        const isBeforeModule7Failure = !reachedModule7;

        if (isBeforeModule7Failure && !errorMessage.includes("REGISTRATION_FAILED_INCOMPLETE")) {
            await saveFailedCardInfo(signupData, finalErrorMessage);
        }

        // ===============================================================================
        // ### 漏洞修复：统一错误处理逻辑 ###
        // 无论失败类型是已知还是未知，都先尝试截图，然后无条件进入清理流程。
        // 这将从根本上解决因“未知失败”导致浏览器窗口残留的问题。
        // ===============================================================================
        console.log(`[${instanceId} 清理] 流程因错误终止，准备截图并关闭浏览器...`);
        if (page) {
            const screenshotDir = path.join(__dirname, 'screenshot');
            await fs.mkdir(screenshotDir, { recursive: true });
            const screenshotPath = path.join(screenshotDir, `error_screenshot_${instanceId}_${Date.now()}.png`);
            try {
                await page.screenshot({ path: screenshotPath, fullPage: true });
                console.log(`[${instanceId}] 错误现场截图已保存: ${screenshotPath}`);
            } catch (screenshotError) {
                console.error(`[${instanceId}] 截取错误截图时失败: ${screenshotError.message}`);
            }
        }

        // 无条件关闭浏览器
        if (browserId) {
             await tearDownBrowser(browserId);
             console.log(`[${instanceId} 清理] 浏览器实例 ${browserId} 已关闭。`);
        } else {
            console.log(`[${instanceId} 清理] browserId 无效，无法执行关闭操作。`);
        }

        process.exit(1);

    } finally {
        if (networkWatcher) { networkWatcher.stop(); }
        currentPage = null;
        if (browserId) {
            try {
                // 最终保障性清理，以防万一有未捕获的路径跳过了 catch 块中的关闭逻辑
                await tearDownBrowser(browserId);
            } catch (e) {
                // 忽略错误，因为它很可能已经被关闭了
            }
        }
    }
}

// --- 脚本主入口 ---
async function main() {
    console.log(`[主控 ${instanceId}] 准备启动任务...`);

    const dataFilePath = path.join(__dirname, 'data', instanceId, 'data.json');
    let allSignupData;

    try {
        const dataContent = await fs.readFile(dataFilePath, 'utf-8');
        allSignupData = JSON.parse(dataContent);
        if (!allSignupData || allSignupData.length === 0) {
            console.log(`[主控 ${instanceId}] 数据文件为空。`);
            process.exit(1);
        }
        if (allSignupData.length > 1) {
            console.warn(`[主控 ${instanceId}] 警告：数据文件包含多条数据，只处理第一条。`);
            allSignupData = [allSignupData[0]];
        }
        console.log(`[主控 ${instanceId}] 从数据文件中加载了 1 条任务。`);
    } catch (error) {
        console.error(`[主控 ${instanceId}] 加载数据文件失败: ${error.message}`);
        process.exit(1);
    }

    const signupData = allSignupData[0];
    const dynamicConfig = await generateDynamicConfig(signupData.country_code, signupData.country_full_name);
    const finalConfig = { ...staticConfig, ...dynamicConfig };

    signupData.country_code = finalConfig.countryCode;
    signupData.country_full_name = finalConfig.countryFullName;

    const account = signupData.account || 'N/A';
    console.log(`STATUS_UPDATE::${JSON.stringify({ instanceId, account, status: "启动中", details: "Node.js进程启动..." })}`);

    try {
        await runWorkflow(signupData, finalConfig);
    } catch (error) {
        console.error(`[主控 ${instanceId}] 工作流最终失败: ${error.message}`);
        process.exit(1);
    }
}

main();