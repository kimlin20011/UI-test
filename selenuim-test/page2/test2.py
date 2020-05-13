from selenium import webdriver

driver = webdriver.Chrome()
driver.get("https://www.google.com")
#找到輸入框
element = driver.find_element_by_name("q")
#輸入內容
element.send_keys("python is awesome")

#element.submit();
element.submit()
# 將該網頁程式碼抓取下來
htmltext = driver.page_source

# 抓取完後關閉網頁
driver.close()


print('hello python')
