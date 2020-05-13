from selenium import webdriver

driver = webdriver.Chrome()
driver.get("'https://google.com")
#找到輸入框
element = driver.find_element_by_name("q");
#輸入內容
element.send_keys("python is awesome");
#提交表單
element.submit();
driver.close()

print('hello python')