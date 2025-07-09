// ===================================================================================
// ### shared/api_helper.js (V5 - 适配 numeric_id) ###
// ===================================================================================
const axios = require('axios').default;
// 1. 引入并加载 JSON 数据文件
// Node.js 会自动解析 JSON 文件为一个 JavaScript 对象
const countryData = require('./combined_country_data.json');

/**
 * @typedef {object} NewPhoneInfo
 * @property {string} phone_number_id
 * @property {string} phone_number
 * @property {string} phone_number_url
 */

/**
 * 从主API获取一个新的手机号信息。
 * @param {string} countryCode - 目标国家的两字母代码, e.g., 'GB'
 * @returns {Promise<NewPhoneInfo>}
 */
async function fetchNewPhoneNumber(countryCode) {
    console.log(`[API助手] 接收到国家代码 ${countryCode}，正在查找其 numeric_id...`);

    // 2. 在 JSON 数据中查找匹配的国家信息
    // Object.values(countryData) 获取所有国家对象组成的数组
    // .find() 遍历数组，找到 country_code 匹配的第一项
    const countryInfo = Object.values(countryData).find(country => country.country_code === countryCode);

    // 如果没有找到对应的国家，则抛出错误
    if (!countryInfo) {
        const errorMessage = `[API助手] 错误：无法在 combined_country_data.json 中找到国家代码为 "${countryCode}" 的条目。`;
        console.error(errorMessage);
        throw new Error(errorMessage);
    }

    // 3. 提取 numeric_id 并替换 countryCode
    const numericId = countryInfo.numeric_id;
    console.log(`[API助手] 找到 ${countryCode} 对应的 numeric_id: ${numericId}。`);

    // 使用 numericId 构建 API URL
    const apiUrl = `https://api.small5.co/hub/en/proxy.php?action=getNumber&service=am&country=${numericId}&platform=sms`;
    
    try {
        console.log(`[API助手] 正在使用 numeric_id ${numericId} 请求新的手机号...`);
        const response = await axios.get(apiUrl, { timeout: 30000 });
        console.log("[API助手] 已收到API的原始响应数据:", JSON.stringify(response.data));

        // 【核心修改】处理字符串格式的响应
        if (typeof response.data === 'string' && response.data.startsWith('ACCESS_NUMBER:')) {
            const parts = response.data.split(':');
            
            if (parts.length >= 3) {
                const phone_number_id = parts[1];
                const phone_number = parts[2];
                // 根据ID自动组装检查短信的URL
                const phone_number_url = `https://api.small5.co/hub/en/proxy.php?action=getStatus&id=${phone_number_id}&platform=sms`;

                console.log(`[API助手] 成功解析手机号: ${phone_number}, ID: ${phone_number_id}`);
                
                return {
                    phone_number_id,
                    phone_number,
                    phone_number_url
                };
            }
        }
        
        // 如果不是预期的字符串格式，则抛出错误
        throw new Error('API响应不是预期的 "ACCESS_NUMBER:..." 字符串格式。');

    } catch (error) {
        console.error(`[API助手] 请求新手机号时发生错误: ${error.message}`);
        throw error;
    }
}

module.exports = { fetchNewPhoneNumber };