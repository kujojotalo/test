
from appium import webdriver

desired_caps = {
  "platformName": "Android",
  "appium:platformVersion": "12.0.0",
  "appium:deviceName": "6a2c02ad",
  "appium:automationName": "UiAutomator2",
  "appium:appPackage": "com.tencent.mm",
  "appium:appActivity": ".ui.LauncherUI",
  "appium:ignoreHiddenApiPolicyError": "true",
  "appium:unicodeKeyboard": False,
  "appium:noReset": True,
  "appium:forceAppLaunch": True
}
# 连接Appium Server，初始化自动化环境
driver = webdriver.Remote('http://127.0.0.1:4723', desired_caps)

driver.implicitly_wait(5)

driver.quit()
