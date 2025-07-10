/**
 * 模块 4.5: 选择免费支持计划
 * ------------------------------------------------
 * @description
 * 此模块处理在模块4后可能出现的“账户计划选择”页面（彩蛋）。
 * 它将自动选择免费的基础套餐，并点击按钮进入下一阶段。
 *
 * @precondition
 * - 当前页面URL为 https://portal.aws.amazon.com/billing/signup?type=register#/accountplan
 *
 * @postcondition
 * - 成功点击“基础支持 - 免费”选项。
 * - 成功点击“选择套餐”按钮。
 * - 页面跳转至模块5（联系人信息填写）。
 *
 */
// 假设项目上下文中已存在以下辅助函数和配置对象
// const { waitForElement, clickElement, log } = require('./utils');
// const config = require('./shared/config');

class SelectFreePlanModule {
    constructor(config, utils) {
        this.config = config;
        this.utils = utils; // 假设的辅助函数库
        this.selectors = this.config.SELECTORS.ACCOUNT_PLAN_PAGE;
    }

    async run() {
        this.utils.log('模块 4.5: 检测到账户计划选择页面，开始执行...');

        try {
            // 步骤 1: 等待并选择“免费套餐”选项
            // 根据我的假设，这里需要先点击包含“Free”文本的元素来选中它
            this.utils.log('正在等待并选择免费套餐...');
            const freePlanElement = await this.utils.waitForElement(this.selectors.freePlanOption);
            await this.utils.clickElement(freePlanElement, '选择免费基础套餐');

            // 步骤 2: 等待并点击提交按钮
            // 这对应您提供的HTML <button...aria-label="Choose paid plan">
            this.utils.log('正在等待并点击“选择套餐”按钮...');
            const choosePlanButton = await this.utils.waitForElement(this.selectors.choosePlanButton);
            await this.utils.clickElement(choosePlanButton, '点击“选择套餐”按钮');

            this.utils.log('模块 4.5: 执行完毕，预计将跳转到模块 5。');
            return true;

        } catch (error) {
            console.error('模块 4.5 执行失败:', error);
            this.utils.log(`错误: 在选择免费套餐过程中发生问题。错误信息: ${error.message}`, 'error');
            // 可以根据需要增加错误处理和重试逻辑
            return false;
        }
    }
}

// 导出模块以供主控制器使用
// module.exports = SelectFreePlanModule;