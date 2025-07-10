// ===================================================================================
// ### shared/api_helper.js (V6 - 增加手机号去区号功能) ###
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
 * 从主API获取一个新的、已去除区号的手机号信息。
 * @param {string} countryCode - 目标国家的两字母代码, e.g., 'GB'
 * @returns {Promise<NewPhoneInfo>}
 */
async function fetchNewPhoneNumber(countryCode) {
    console.log(`[API助手] 接收到国家代码 ${countryCode}，正在查找其 numeric_id 和 dialing_code...`);

    // 2. 在 JSON 数据中查找匹配的国家信息
    const countryInfo = Object.values(countryData).find(country => country.country_code === countryCode);

    if (!countryInfo || !countryInfo.dialing_code) {
        const errorMessage = `[API助手] 错误：无法在 combined_country_data.json 中找到国家代码为 "${countryCode}" 的有效条目或其区号。`;
        console.error(errorMessage);
        throw new Error(errorMessage);
    }

    // 3. 提取 numeric_id 和 dialing_code
    const numericId = countryInfo.numeric_id;
    const dialingCode = countryInfo.dialing_code;
    console.log(`[API助手] 找到 ${countryCode} -> numeric_id: ${numericId}, dialing_code: ${dialingCode}。`);

    const apiUrl = `https://api.small5.co/hub/en/proxy.php?action=getNumber&service=am&country=${numericId}&platform=sms`;
    
    try {
        console.log(`[API助手] 正在使用 numeric_id ${numericId} 请求新的手机号...`);
        const response = await axios.get(apiUrl, { timeout: 30000 });
        console.log("[API助手] 已收到API的原始响应数据:", JSON.stringify(response.data));

        if (typeof response.data === 'string' && response.data.startsWith('ACCESS_NUMBER:')) {
            const parts = response.data.split(':');
            
            if (parts.length >= 3) {
                const phone_number_id = parts[1];
                const raw_phone_number = parts[2]; // 这是原始的、可能带区号的号码
                let cleaned_phone_number;

                // =================================================================
                // ### 核心修改：增加去除区号的逻辑 ###
                // =================================================================
                if (raw_phone_number.startsWith(dialingCode)) {
                    cleaned_phone_number = raw_phone_number.substring(dialingCode.length);
                    console.log(`[API助手] 区号匹配成功。原始号码: ${raw_phone_number} -> 去除区号后: ${cleaned_phone_number}`);
                } else {
                    cleaned_phone_number = raw_phone_number;
                    console.warn(`[API助手] 警告: 获取的手机号 ${raw_phone_number} 不以区号 ${dialingCode} 开头，将直接使用。`);
                }
                // =================================================================
                
                const phone_number_url = `https://api.small5.co/hub/en/proxy.php?action=getStatus&id=${phone_number_id}&platform=sms`;

                console.log(`[API助手] 成功返回处理后的手机号: ${cleaned_phone_number}, ID: ${phone_number_id}`);
                
                return {
                    phone_number_id,
                    phone_number: cleaned_phone_number, // 返回处理过的号码
                    phone_number_url
                };
            }
        }
        
        throw new Error('API响应不是预期的 "ACCESS_NUMBER:..." 字符串格式。');

    } catch (error) {
        console.error(`[API助手] 请求新手机号时发生错误: ${error.message}`);
        throw error;
    }
}

module.exports = { fetchNewPhoneNumber };