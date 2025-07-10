// ===================================================================================
// ### test_api.js (用于独立测试API助手) ###
// ===================================================================================

// 1. 引入我们需要测试的模块
const { fetchNewPhoneNumber } = require('./shared/api_helper.js');

// 2. 定义一个主函数来执行测试
async function runTest() {
    console.log("准备开始独立API测试...");
    
    // 您可以在这里更改想要测试的国家代码
    const targetCountryCode = 'SE'; 
    
    try {
        console.log(`正在为国家 ${targetCountryCode} 请求新的手机号...`);
        
        // 3. 调用核心函数
        const newPhoneInfo = await fetchNewPhoneNumber(targetCountryCode);
        
        // 4. 如果成功，打印出获取到的信息
        console.log("\n✅✅✅ API请求成功！ ✅✅✅");
        console.log("获取到的手机号信息如下:");
        console.log(JSON.stringify(newPhoneInfo, null, 2)); // 使用格式化的JSON输出，更清晰

    } catch (error) {
        // 5. 如果失败，打印出错误信息
        console.error("\n❌❌❌ API请求失败！ ❌❌❌");
        console.error("错误详情:", error.message);
    }
}

// 6. 运行测试
runTest();