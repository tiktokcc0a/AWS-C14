// ===================================================================================
// ### shared/config.js (V5.3 - 修正模块4.5目标) ###
// ===================================================================================

module.exports = {
    // --- URL ---
    AWS_SIGNUP_URL: "https://signin.aws.amazon.com/signup?request_type=register",
    AWS_IAM_WIZARD_URL: "https://us-east-1.console.aws.amazon.com/iam/home?region=us-east-1#/security_credentials/access-key-wizard",

    ACCOUNT_PLAN_PAGE_URL_PART: "/billing/signup?type=register#/accountplan",

    // --- 文件保存路径 ---
    KEY_SAVE_PATH: "C:\\Users\\Administrator\\Desktop\\AWS-ACCOUNT",

    // --- 天天识图 API ---
    TTSHITU_API_URL: 'http://api.ttshitu.com/predict',
    TTSHITU_USERNAME: 'tao194250',
    TTSHITU_PASSWORD: 'Too194250',
    TTSHITU_TYPEID: '7',

    // --- 邮箱API配置 ---
    EMAIL_API_CONFIG: {
        'Domain': {
            url: 'https://mail.32v.us/api/index.php',
        },
        'Outlook': {
            url: 'https://authhk.bhdata.com:30015/9W7ZJ41/',
        }
    },
    EMAIL_API_BASE_URL: 'https://mail.32v.us/api/index.php',

    // ===============================================================================
    // ### CSS 选择器 ###
    // ===============================================================================
    EMAIL_INPUT_SELECTOR: '#emailAddress',
    ACCOUNT_NAME_INPUT_SELECTOR: '#accountName',
    CAPTCHA_TRIGGER_SELECTOR: 'button[data-testid="collect-email-submit-button"]',
    COMMON_CAPTCHA_IFRAME_SELECTOR: 'iframe#core-container',
    COMMON_CAPTCHA_IFRAME_URL_PART: 'threat-mitigation.aws.amazon.com',
    INITIAL_CAPTCHA_IMAGE_SELECTOR: 'img[alt="captcha"]',
    INITIAL_CAPTCHA_INPUT_SELECTOR: 'input[name="captchaGuess"]',
    INITIAL_CAPTCHA_SUBMIT_SELECTOR: 'button[type="submit"]',
    INITIAL_CAPTCHA_ERROR_SELECTOR: 'span[data-testid="error-message"], div[data-testid="error-message"], p.awsui-form-field-error-message',
    INITIAL_CAPTCHA_EXPECTED_ERROR: "That wasn't quite right, please try again.",
    OTP_INPUT_SELECTOR: '#otp',
    OTP_SUBMIT_BUTTON_SELECTOR: 'button[data-testid="verify-email-submit-button"]',
    PASSWORD_INPUT_SELECTOR: '#password',
    RE_PASSWORD_INPUT_SELECTOR: '#rePassword',
    CREATE_PASSWORD_SUBMIT_SELECTOR: 'button[data-testid="create-password-submit-button"]',

    // 【已修改】模块4.5的选择器，只保留付费计划按钮
    ACCOUNT_PLAN_SELECTORS: {
        choosePlanButton: 'button.awsui-button-variant-primary[type="submit"][aria-label="Choose paid plan"]'
    },

    PERSONAL_ACCOUNT_RADIO_SELECTOR: '#awsui-radio-button-2',
    CONTACT_FULL_NAME_SELECTOR: 'input[name="address.fullName"]',
    CONTACT_PHONE_COUNTRY_TRIGGER_SELECTOR: '#awsui-select-0',
    CONTACT_ADDRESS_COUNTRY_TRIGGER_SELECTOR: '#awsui-select-1',
    IDENTITY_PHONE_COUNTRY_TRIGGER_SELECTOR: 'div.awsui-select-trigger ::-p-text(United States (+1))',
    CONTACT_STREET_SELECTOR: 'input[name="address.addressLine1"]',
    CONTACT_CITY_SELECTOR: 'input[name="address.city"]',
    CONTACT_STATE_SELECTOR: 'input[name="address.state"]',
    CONTACT_POSTCODE_SELECTOR: 'input[name="address.postalCode"]',
    CONTACT_PHONE_NUMBER_SELECTOR: 'input[name="address.phoneNumber"]',
    CONTACT_AGREEMENT_CHECKBOX_SELECTOR: 'input[type="checkbox"][name="agreement"]',
    CONTACT_SUBMIT_BUTTON_SELECTOR: 'button.awsui-button-variant-primary[type="submit"][aria-label="Agree and Continue (step 2 of 5)"]',
    PAYMENT_CARD_NUMBER_SELECTOR: '#awsui-input-1',
    PAYMENT_CARD_HOLDER_NAME_SELECTOR: '#awsui-input-3',
    PAYMENT_CVV_SELECTOR: '#awsui-input-2',
    PAYMENT_MONTH_TRIGGER_SELECTOR: '#awsui-select-1',
    PAYMENT_YEAR_TRIGGER_SELECTOR: '#awsui-select-2',
    PAYMENT_SUBMIT_BUTTON_SELECTOR: 'button ::-p-text(Verify and continue (step 3 of 5))',
    PAYMENT_PAGE_FAQ_SELECTOR: 'span.LinkButton_linkButton__eGLo',
    IDENTITY_PHONE_NUMBER_SELECTOR: 'input[name="phoneNumber"]',
    IDENTITY_SEND_SMS_BUTTON_SELECTOR: 'button.awsui-button-variant-primary[type="submit"]',
    IDENTITY_CAPTCHA_IMAGE_SELECTOR: 'img[alt="captcha"]',
    IDENTITY_CAPTCHA_INPUT_SELECTOR: 'input[name="captchaGuess"]',
    IDENTITY_CAPTCHA_SUBMIT_SELECTOR: 'button.awsui_button_vjswe_1379u_157.awsui_variant-primary_vjswe_1379u_230[type="submit"]',
    IDENTITY_CAPTCHA_ERROR_SELECTOR: 'div.awsui_error_1i0s3_1goap_185#form-error-\\\:r0\\\:',
    IDENTITY_SMS_PIN_INPUT_SELECTOR: 'input#awsui-input-2',
    IDENTITY_CONTINUE_BUTTON_SELECTOR: 'button ::-p-text(Continue (step 4 of 5))',
    FINAL_PHONE_VERIFY_COUNTRY_TRIGGER_SELECTOR: 'div[role="button"][aria-haspopup="listbox"]',
    SUPPORT_PLAN_SUBMIT_BUTTON: 'button ::-p-text(Complete sign up)',
    IAM_UNDERSTAND_CHECKBOX: 'input[name="ack-risk"]', 
    IAM_CREATE_KEY_BUTTON: 'button ::-p-text(Create access key)',
    IAM_SHOW_SECRET_BUTTON: 'strong ::-p-text(Show)',
    IAM_ACCESS_KEY_VALUE: 'span[data-testid="inner-text"]',
    IAM_SECRET_KEY_VALUE: 'span[data-testid="shown-inner-text"]',
    IAM_DOWNLOAD_BUTTON: 'button ::-p-text(Download .csv file)'
};